"""
ios_nativo.py
=============
Llena el form en iPhone Safari DENTRO de la landing, por la capa nativa del iPhone.

Por qué así: el form vive en un iframe de otro dominio y Safari en iPhone bloquea que
la automatización entre a ese iframe (switch_to.frame -> "Blocked a frame ...
cross-origin"). La capa nativa (NATIVE_APP) ve la pantalla como la ve el sistema del
iPhone: el form aparece sin importar el dominio, y se toca y tipea como una persona.

La capa nativa no ve ids, pero cada control trae su etiqueta visible ("Nombre",
"Cédula", "Ciudad"...). Con esa etiqueta se detecta qué es cada campo, igual que la app
lo hace por id en los demás navegadores: el que decide qué valor va es quien llama
(callback valor_para), con la lógica de siempre del proyecto.

Cómo ve el iPhone cada control:
  texto     -> XCUIElementTypeTextField / TextView, con un StaticText (su label) adentro
  select    -> XCUIElementTypeOther cuyo value es la opción elegida ("Seleccionar" al
               principio), con un StaticText (su label) adentro. Al tocarlo, iOS 16 abre un
               menú (CollectionView con un botón por opción); versiones viejas, una ruedita.
  checkbox  -> XCUIElementTypeSwitch, value "0"/"1", con su texto al lado
  enviar    -> XCUIElementTypeButton "Enviar"

Detalles de la capa nativa:
  - No hay JavaScript: nada de is_displayed()/get_attribute() de Selenium (inyectan JS).
    Se lee page_source (XML, una llamada por pantalla) y se ubica cada elemento por XPath
    con su posición actual.
  - Al tipear, Safari hace zoom y cambian posiciones y tamaños. Por eso no se calcula nada
    de antemano: después de cada acción se vuelve a leer la pantalla.
  - Para scrollear se vuelve a la capa web (WEBVIEW): ahí la landing sí acepta JS.
"""

import random
import time
import unicodedata
import xml.etree.ElementTree as ET
from typing import Callable, Dict, List, Optional

from selenium.webdriver.common.by import By

_CMD_CONTEXTOS = ("getContexts", "GET", "/session/$sessionId/contexts")
_CMD_CONTEXTO = ("switchToContext", "POST", "/session/$sessionId/context")

_T = "XCUIElementType"
_BOTONES_CERRAR = ("Done", "Listo", "OK", "Aceptar")

def _kw_botones():
    """Textos de botón Siguiente / Enviar: los mismos que usa la app (lt_runner)."""
    from osocio.core.campos_inferencia import _NEXT_KEYWORDS, _SUBMIT_KEYWORDS
    return (tuple(_norm(k) for k in _NEXT_KEYWORDS), tuple(_norm(k) for k in _SUBMIT_KEYWORDS))
# Checkbox que se marcan: términos, privacidad, autorizaciones (mismo criterio que
# _handle_terms_checkboxes de la app, por texto en vez de por name).
_CHECKBOX_A_MARCAR = ("acepto", "termin", "autoriz", "privacidad", "politica", "consent",
                      "concordo", "aceito", "declaro", "li e")
# Respaldo de la detección de TY (la principal es estructural, ver esperar_ty).
_TEXTOS_TY = ("gracias por", "solicitud enviada", "se pondra en contacto", "hemos recibido",
              "nos pondremos en contacto", "obrigado", "recebemos sua", "enviado com sucesso",
              "thank you for")
# Cartel de error de event_id (el form no llegó a enviarse y se recarga solo): la misma
# lista que la app (BaseFormFiller._EVENT_ID_PATTERNS), única fuente.
def _patrones_event_id():
    from osocio.core.base_form_filler import BaseFormFiller
    return tuple(dict.fromkeys(_norm(p) for p in BaseFormFiller._EVENT_ID_PATTERNS))


# valor_para() devuelve esto cuando el Excel pidió NO completar el campo ('-', 'omitir'...).
OMITIR = object()

# Franja de pantalla donde se actúa: arriba la tapa la barra de estado, abajo la barra de
# Safari y el teclado / menú de los selects.
_Y_MIN, _Y_MAX = 60, 620


def _texto(e) -> str:
    # Según el control, el texto visible viene en label, name o value.
    return (e.get("label") or e.get("name") or e.get("value") or "").strip()


def _norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", (s or "").lower())
    return " ".join("".join(c for c in s if not unicodedata.combining(c)).split())


def _num(e, k) -> int:
    try:
        return int(float(e.get(k) or 0))
    except ValueError:
        return 0


def _adentro(hijo, padre) -> bool:
    cx = _num(hijo, "x") + _num(hijo, "width") / 2
    cy = _num(hijo, "y") + _num(hijo, "height") / 2
    return (_num(padre, "x") <= cx <= _num(padre, "x") + _num(padre, "width")
            and _num(padre, "y") <= cy <= _num(padre, "y") + _num(padre, "height"))


class FormNativoIOS:
    def __init__(self, driver, iframe, log: Callable = print):
        self.d = driver
        self.iframe = iframe
        self.log = log
        for nombre, metodo, ruta in (_CMD_CONTEXTOS, _CMD_CONTEXTO):
            driver.command_executor._commands[nombre] = (metodo, ruta)
        contextos = driver.execute("getContexts")["value"] or []
        if "NATIVE_APP" not in contextos:
            raise RuntimeError(f"El iPhone no expone la capa nativa: {contextos}")
        self.web = next(c for c in contextos if c != "NATIVE_APP")
        # En la capa nativa el iframe aparece como un contenedor con su title.
        self.titulo_iframe = (iframe.get_attribute("title") or "").strip()

    # ── contextos y scroll ───────────────────────────────────────────────────
    def _nativo(self):
        self.d.execute("switchToContext", {"name": "NATIVE_APP"})

    def _webview(self):
        self.d.execute("switchToContext", {"name": self.web})

    def _offset(self) -> int:
        return int(self.d.execute_script("return window.pageYOffset") or 0)

    def _scroll_por(self, dy: int) -> bool:
        """Scrollea la landing. False si ya no se movió (fin de la página)."""
        antes = self._offset()
        self.d.execute_script("window.scrollBy(0, arguments[0]);", dy)
        time.sleep(0.7)
        return self._offset() != antes

    def ir_al_form(self):
        self._webview()
        self.d.execute_script("arguments[0].scrollIntoView({block:'start'});", self.iframe)
        time.sleep(1.2)

    # ── lectura de la pantalla ───────────────────────────────────────────────
    def _fuente(self):
        """La pantalla nativa como árbol XML.

        Sin el atributo "visible": calcularlo para cada elemento es lo más caro de leer la
        pantalla del iPhone (~4-5 s por lectura con él). La visibilidad se deduce de la
        posición (ver _visible). Si el device no acepta la opción, se lee completa.
        """
        try:
            xml = self.d.execute_script("mobile: source",
                                        {"format": "xml", "excludedAttributes": "visible,accessible"})
            if xml:
                return ET.fromstring(xml)
        except Exception:
            pass
        return ET.fromstring(self.d.page_source)

    def _alto_pantalla(self) -> int:
        if not getattr(self, "_alto", None):
            try:
                self._alto = int(self.d.get_window_size().get("height") or 932)
            except Exception:
                self._alto = 932
        return self._alto

    def _visible(self, e) -> bool:
        """Visible = con tamaño y dentro de la pantalla (o el atributo, si vino)."""
        if e.get("visible") is not None:
            return e.get("visible") == "true"
        y, h, w = _num(e, "y"), _num(e, "height"), _num(e, "width")
        return w > 0 and h > 0 and y + h > 0 and y < self._alto_pantalla()

    def escanear(self) -> Dict:
        """Controles del form a la vista, de arriba a abajo. Hay que estar en NATIVE_APP."""
        raiz = self._fuente()
        web = raiz.find(f".//{_T}WebView")
        todos = list(web.iter()) if web is not None else []
        estaticos = [e for e in todos if e.tag == f"{_T}StaticText" and _texto(e)]
        # Los checkboxes que cuentan son los del form (la landing puede tener otros).
        contenedor = self._contenedor_del_form(raiz)

        def etiqueta_de(ctrl, excluir=""):
            for s in estaticos:
                t = _texto(s)
                if t != excluir and _adentro(s, ctrl):
                    return t
            return ""

        kw_siguiente, kw_enviar = _kw_botones()
        controles, enviar, siguiente = [], None, None
        for e in todos:
            tag = e.tag.replace(_T, "")
            visible = self._visible(e)
            if tag in ("TextField", "TextView", "SecureTextField") and visible:
                controles.append({"tipo": "texto", "e": e, "valor": (e.get("value") or "").strip(),
                                  "etiqueta": etiqueta_de(e) or (e.get("placeholderValue") or "")})
            elif tag == "Other" and visible and (e.get("value") or "").strip() \
                    and not (e.get("label") or "").strip() and 30 <= _num(e, "height") <= 100:
                valor = e.get("value").strip()
                etiqueta = etiqueta_de(e, excluir=valor)
                if etiqueta:   # sin label adentro no es un select del form
                    controles.append({"tipo": "select", "e": e, "valor": valor, "etiqueta": etiqueta,
                                      "habilitado": e.get("enabled") == "true"})
            elif tag == "Switch" and (contenedor is None or _adentro(e, contenedor)):
                controles.append(self._checkbox(e, estaticos))
            elif tag == "Button" and visible and _norm(_texto(e)).startswith(kw_enviar):
                enviar = e
            elif tag == "Button" and visible and _norm(_texto(e)).startswith(kw_siguiente):
                siguiente = e
        # Duplicados (el mismo control anidado dos veces). Un checkbox se ubica por SU fila:
        # los de un grupo empiezan todos en la misma línea (el título del grupo).
        vistos, out = set(), []
        for c in sorted(controles, key=lambda c: c.get("fila", _num(c["e"], "y"))):
            c["y"] = c.get("fila", _num(c["e"], "y"))
            clave = (c["tipo"], c["y"] // 8)
            if clave in vistos:
                continue
            vistos.add(clave)
            out.append(c)
        textos = [{"texto": _texto(s), "x": _num(s, "x"), "y": _num(s, "y")} for s in estaticos]
        return {"controles": out, "enviar": enviar, "siguiente": siguiente, "textos": textos}

    @staticmethod
    def _checkbox(e, estaticos) -> Dict:
        """Un checkbox (XCUIElementTypeSwitch) con su opción, su grupo y dónde tocarlo.

        En los grupos ("Preferencia de contacto": E-mail / Llamada / WhatsApp) cada opción es
        un Switch que EMPIEZA en el título del grupo y se estira hasta su fila: su opción es
        el texto que termina donde termina el Switch, no el primero que tiene adentro. Se
        toca el cuadradito (entre el borde del Switch y el texto), no el centro: el centro
        cae en el título del grupo, o en el link de "términos y condiciones".
        """
        x, y = _num(e, "x"), _num(e, "y")
        abajo = y + _num(e, "height")
        adentro = [s for s in estaticos if _adentro(s, e)]
        opcion = min(adentro, key=lambda s: abs(_num(s, "y") + _num(s, "height") - abajo), default=None)
        if opcion is not None and abs(_num(opcion, "y") + _num(opcion, "height") - abajo) > 20:
            opcion = None
        etiqueta = _texto(opcion) if opcion is not None else ""
        # El título del grupo está ARRIBA de la opción (un superíndice "1" de la misma
        # línea, como en "términos y condiciones¹", no es un título).
        tope = _num(opcion, "y") - 8 if opcion is not None else y + _num(e, "height")
        titulo = next((_texto(s) for s in adentro
                       if s is not opcion and abs(_num(s, "y") - y) <= 4 and _num(s, "y") < tope
                       and len(_texto(s)) > 2 and not _texto(s).isdigit()), "")
        if opcion is not None:
            fila = _num(opcion, "y") + _num(opcion, "height") // 2
            tap_x = (x + _num(opcion, "x")) // 2 if _num(opcion, "x") > x + 8 else x + 12
        else:
            fila, tap_x = y + _num(e, "height") // 2, x + 12
        return {"tipo": "checkbox", "e": e, "valor": (e.get("value") or "0"), "etiqueta": etiqueta,
                "grupo": titulo, "fila": fila, "abajo": abajo, "tap": (tap_x, fila)}

    def _elemento(self, e):
        """El WebElement de un nodo del page_source, por tipo y posición actual."""
        xp = f"//{_T}WebView//{e.tag}[@x='{_num(e, 'x')}' and @y='{_num(e, 'y')}']"
        encontrados = self.d.find_elements(By.XPATH, xp)
        return encontrados[0] if encontrados else None

    def _tocar_centro(self, e):
        self.d.execute_script("mobile: tap", {"x": _num(e, "x") + _num(e, "width") // 2,
                                              "y": _num(e, "y") + _num(e, "height") // 2})

    def _cerrar_teclado(self):
        for nombre in _BOTONES_CERRAR:
            for b in self.d.find_elements(By.XPATH, f"//{_T}Button[@name='{nombre}']"):
                try:
                    b.click()
                    time.sleep(0.5)
                    return
                except Exception:
                    continue

    # ── acciones ─────────────────────────────────────────────────────────────
    def _llenar_texto(self, c, valor: str) -> str:
        el = self._elemento(c["e"])
        if el is None:
            return "no encontrado"
        el.click()
        time.sleep(0.7)
        el.send_keys(valor)
        time.sleep(0.3)
        quedo = el.get_dom_attribute("value") or ""
        self._cerrar_teclado()
        return quedo

    def _elegir(self, c, opcion: Optional[str]) -> str:
        el = self._elemento(c["e"])
        if el is None:
            return "no encontrado"
        el.click()
        time.sleep(1.2)
        ruedas = self.d.find_elements(By.XPATH, f"//{_T}PickerWheel")
        if ruedas:
            return self._elegir_en_ruedita(ruedas[0], opcion)
        return self._elegir_en_menu(opcion)

    def _elegir_en_ruedita(self, rueda, opcion: Optional[str]) -> str:
        try:
            if not opcion:
                raise ValueError
            rueda.send_keys(opcion)
        except Exception:
            self.d.execute_script("mobile: selectPickerWheelValue",
                                  {"elementId": rueda.id, "order": "next", "offset": 0.15})
        time.sleep(0.6)
        valor = rueda.get_dom_attribute("value") or ""
        self._cerrar_teclado()
        time.sleep(1.0)
        return valor

    def _elegir_en_menu(self, opcion: Optional[str]) -> str:
        """Elige en el menú de iOS la opción pedida o, si no hay, una real al azar."""
        from osocio.core.campos_inferencia import _is_placeholder

        raiz = self._fuente()
        menu = raiz.find(f".//{_T}CollectionView")
        etiquetas = [] if menu is None else [_texto(b) for b in menu.iter(f"{_T}Button") if _texto(b)]
        if not etiquetas:
            return "no abrió ni ruedita ni menú"
        reales = [t for t in etiquetas if not _is_placeholder(t)]
        if not reales:
            self._cerrar_menu()
            return "sin opciones"
        elegida = None
        if opcion:
            o = _norm(opcion)
            elegida = next((t for t in reales if _norm(t) == o), None) \
                or next((t for t in reales if o in _norm(t) or _norm(t) in o), None)
        elegida = elegida or random.choice(reales)
        self.log(f"      menú: {len(reales)} opciones reales {reales[:6]}{'...' if len(reales) > 6 else ''}")
        comillas = "'" if '"' in elegida else '"'
        boton = self.d.find_element(
            By.XPATH, f"//{_T}CollectionView//{_T}Button[@label={comillas}{elegida}{comillas}]")
        if boton.get_dom_attribute("visible") != "true":
            self.d.execute_script("mobile: scroll", {"elementId": boton.id, "toVisible": True})
            time.sleep(0.4)
        boton.click()
        time.sleep(1.2)  # el select siguiente (ej. concesionario) carga opciones al cambiar
        return elegida

    def _cerrar_menu(self):
        # Tocar arriba de todo, fuera del menú, lo cierra sin elegir.
        try:
            self.d.execute_script("mobile: tap", {"x": 20, "y": 120})
            time.sleep(0.5)
        except Exception:
            pass

    def _marcar_checkbox(self, c) -> str:
        """Hay que estar en NATIVE_APP y con el checkbox a la vista."""
        if c["valor"] == "1":
            return "ya estaba marcado"
        nuevo = c
        for _intento in range(2):   # a veces el primer toque después de scrollear no entra
            tap_x, tap_y = nuevo.get("tap") or (_num(nuevo["e"], "x") + 12, nuevo["y"])
            self.d.execute_script("mobile: tap", {"x": tap_x, "y": tap_y})
            time.sleep(0.6)
            nuevo = self._mismo_checkbox(c)
            if nuevo is None:
                return "tocado"
            if nuevo["valor"] == "1":
                return "marcado"
        return f"sigue en {nuevo['valor']}"

    def checkboxes_pendientes(self) -> List[Dict]:
        """Los checkboxes sin marcar que la regla de la app manda marcar (términos, *)."""
        self._nativo()
        try:
            return [c for c in self.escanear()["controles"]
                    if c["tipo"] == "checkbox" and self._pendiente(c)]
        finally:
            self._webview()

    def _mismo_checkbox(self, c):
        """El checkbox c en la pantalla actual (por etiqueta y grupo; si no, por posición)."""
        cbs = [n for n in self.escanear()["controles"] if n["tipo"] == "checkbox"]
        if c.get("etiqueta"):
            for n in cbs:
                if n["etiqueta"] == c["etiqueta"] and n.get("grupo") == c.get("grupo"):
                    return n
        return next((n for n in cbs if abs(n["y"] - c["y"]) < 15), None)

    def checkboxes_rechazados(self) -> List[Dict]:
        """Después de tocar Enviar: los checkboxes sin marcar que el form exige.

        Es la lógica de la app para los checkboxes obligatorios por JS (#27), con lo que se
        ve en el iPhone: el form frena el envío y pone un aviso justo DEBAJO del checkbox o
        de su grupo ("Seleccione al menos una opción.", "Acepta términos y condiciones.").
        Si el aviso es de un grupo, se devuelven todas sus opciones sin marcar (como en
        Chrome, que marca todos los checkboxes con error).
        """
        self._nativo()
        try:
            pantalla = self.escanear()
        finally:
            self._webview()
        # Sin texto propio ni de grupo no es una casilla del form que se pueda pedir
        # (en Chile hay una así, invisible, que no se deja marcar).
        cbs = [c for c in pantalla["controles"]
               if c["tipo"] == "checkbox" and (c["etiqueta"] or c.get("grupo"))]
        etiquetas = {c["etiqueta"] for c in cbs} | {c.get("grupo", "") for c in cbs}
        grupos: Dict[str, List[Dict]] = {}
        for c in cbs:
            grupos.setdefault(c.get("grupo") or c["etiqueta"] or str(c["y"]), []).append(c)
        rechazados = []
        for opciones in grupos.values():
            abajo = max(c.get("abajo", c["y"]) for c in opciones)
            x_min = min(_num(c["e"], "x") for c in opciones)
            aviso = next((t for t in pantalla["textos"]
                          if t["texto"] not in etiquetas and len(t["texto"]) > 3
                          and abajo - 2 <= t["y"] <= abajo + 45 and abs(t["x"] - x_min) <= 60), None)
            if aviso is not None:
                rechazados += [c for c in opciones if c["valor"] != "1"]
        return rechazados

    def marcar(self, c) -> str:
        """Marca un checkbox desde la capa web/nativa: lo trae a la vista si hace falta."""
        try:
            self._webview()
            if not (_Y_MIN <= c["y"] <= _Y_MAX):
                self._scroll_por(int(c["y"] - 300))
            self._nativo()
            actual = self._mismo_checkbox(c) or c
            return self._marcar_checkbox(actual)
        finally:
            self._webview()

    # ── recorrido completo ───────────────────────────────────────────────────
    def _pendiente(self, c) -> bool:
        """Un control falta llenar según lo que MUESTRA (no según una lista de hechos):
        así el zoom de Safari o un re-render no hacen que se llene dos veces."""
        from osocio.core.campos_inferencia import _is_placeholder
        if c["tipo"] == "texto":
            return not c["valor"].strip()
        if c["tipo"] == "select":
            return _is_placeholder(c["valor"])
        # Checkbox: se marca si es de términos/privacidad o si es obligatorio (* en el label),
        # igual que la app ("se marca si es obligatorio o de términos").
        texto = f"{c['etiqueta']} {c.get('grupo', '')}"
        return c["valor"] != "1" and ("*" in texto
                                      or any(k in _norm(texto) for k in _CHECKBOX_A_MARCAR))

    def llenar(self, valor_para: Callable[[str, str], Optional[str]],
               limite_seg: int = 360,
               al_llenar: Optional[Callable[[Dict], None]] = None) -> List[Dict]:
        """Recorre el form de arriba a abajo y llena cada control que falte.

        valor_para(tipo, etiqueta) -> valor de texto, opción de select (None = al azar),
        u OMITIR para no tocar el campo.
        al_llenar(registro) se llama después de cada campo (para registrar lo que quedó).
        Devuelve un registro por acción: tipo, etiqueta, pedido, quedó.
        """
        self.ir_al_form()
        intentos: Dict[tuple, int] = {}
        registro = []
        pasos = 0
        fin = time.time() + limite_seg
        while time.time() < fin:
            self._nativo()
            try:
                pantalla = self.escanear()
                pendientes = []
                for c in pantalla["controles"]:
                    if not (_Y_MIN <= c["y"] <= _Y_MAX) or not self._pendiente(c):
                        continue
                    clave = (c["tipo"], _norm(c["etiqueta"]))
                    if intentos.get(clave, 0) >= 2:   # ya se intentó dos veces: no insistir
                        continue
                    pendientes.append((clave, c))
                if pendientes:
                    clave, c = pendientes[0]
                    intentos[clave] = intentos.get(clave, 0) + 1
                    registro.append(self._accion(c, valor_para))
                    r = registro[-1]
                    if r["quedo"] in ("(sin dato)", "(omitido)"):
                        intentos[clave] = 99   # omitido a propósito: no volver a intentarlo
                    if al_llenar is not None:
                        try:
                            al_llenar(r)
                        except Exception:
                            pass
                    self.log(f"    · {r['tipo']:8} {r['etiqueta'][:28]:28} "
                             f"pedido='{r['pedido']}' quedó='{r['quedo']}'")
                    continue
                enviar, siguiente = pantalla["enviar"], pantalla["siguiente"]
                # Form en pasos (wizard): con el paso completo a la vista, avanzar.
                if enviar is None and siguiente is not None and _Y_MIN <= _num(siguiente, "y") <= _Y_MAX                         and pasos < 6:
                    el = self._elemento(siguiente)
                    if el is not None:
                        el.click()
                        pasos += 1
                        self.log(f"    → paso siguiente ('{_texto(siguiente)}')")
                        time.sleep(1.5)
                        continue
            finally:
                self._webview()
            # Nada pendiente a la vista: si ya se ve Enviar, terminó el form; si no, bajar.
            if enviar is not None and _num(enviar, "y") <= _Y_MAX:
                break
            if not self._scroll_por(250):
                break
        else:
            self.log(f"    ⚠ se cortó el recorrido a los {limite_seg}s")
        return registro

    # Mismos tiempos que la app para dropdowns dependientes (desplegables.py,
    # _get_dependency_wait_settings): hasta 8 s, revisando cada 0,2 s, 2 reintentos.
    _DEP_TIMEOUT, _DEP_POLL, _DEP_REINTENTOS = 8.0, 0.2, 2

    def _esperar_habilitado(self, c):
        """Hijo de una dependencia (ej. concesionario <- ciudad): esperar a que se habilite."""
        limite = time.time() + self._DEP_TIMEOUT
        while time.time() < limite:
            nuevo = next((n for n in self.escanear()["controles"]
                          if n["tipo"] == "select" and n["etiqueta"] == c["etiqueta"]), None)
            if nuevo is not None and nuevo.get("habilitado"):
                return nuevo
            time.sleep(self._DEP_POLL)
        return None

    def _accion(self, c, valor_para) -> Dict:
        reg = {"tipo": c["tipo"], "etiqueta": c["etiqueta"], "pedido": "", "quedo": ""}
        try:
            if c["tipo"] in ("texto", "select"):
                pedido = valor_para(c["tipo"], c["etiqueta"])
                if pedido is OMITIR:
                    reg["pedido"], reg["quedo"] = "-", "(omitido)"
                    return reg
            if c["tipo"] == "texto":
                reg["pedido"] = pedido or ""
                reg["quedo"] = self._llenar_texto(c, reg["pedido"]) if reg["pedido"] else "(sin dato)"
            elif c["tipo"] == "select":
                if not c.get("habilitado", True):
                    listo = self._esperar_habilitado(c)
                    if listo is None:
                        reg["quedo"] = f"no se habilitó en {self._DEP_TIMEOUT:.0f}s (¿falta el campo padre?)"
                        return reg
                    c = listo
                reg["pedido"] = pedido or "(al azar)"
                for intento in range(self._DEP_REINTENTOS + 1):
                    reg["quedo"] = self._elegir(c, pedido)
                    if reg["quedo"] != "sin opciones":
                        break
                    # Habilitado pero todavía sin opciones reales: el padre recién cambió.
                    time.sleep(1.5)
                    c = self._esperar_habilitado(c) or c
                if not reg["quedo"].startswith(("no ", "sin ", "ERROR")):
                    # Lo que quedó es lo que MUESTRA el select ahora, no lo que se tocó.
                    elegida = reg["quedo"]
                    ahora = next((n for n in self.escanear()["controles"]
                                  if n["tipo"] == "select" and n["etiqueta"] == c["etiqueta"]), None)
                    reg["quedo"] = ahora["valor"] if ahora is not None else elegida
                    if ahora is not None and _norm(ahora["valor"]) != _norm(elegida):
                        self.log(f"      ⚠ se tocó '{elegida}' pero el select muestra '{ahora['valor']}'")
            elif c["tipo"] == "checkbox":
                reg["pedido"] = "marcar"
                reg["quedo"] = self._marcar_checkbox(c)
        except Exception as e:
            reg["quedo"] = f"ERROR {str(e).splitlines()[0][:120]}"
        return reg

    def tocar_enviar(self) -> bool:
        self._webview()
        for _ in range(8):
            self._nativo()
            try:
                enviar = self.escanear()["enviar"]
                if enviar is not None and _Y_MIN <= _num(enviar, "y") <= _Y_MAX:
                    el = self._elemento(enviar)
                    if el is not None:
                        el.click()
                        return True
            finally:
                self._webview()
            if not self._scroll_por(200):
                break
        return False

    def esperar_ty(self, timeout: int = 15):
        """Espera el resultado del envío. Devuelve (estado, texto):
          ("ty", texto)        apareció la thank you page
          ("event_id", texto)  apareció el cartel de error de event_id (hay que reintentar)
          ("nada", "")         ninguna de las dos (se deja diagnóstico en el log)

        TY con el mismo criterio que la app (lt_runner, _ty_visible): el form REEMPLAZADO
        por otro contenido (div.rp-wrapper / div#thank-you), no un texto fijo, porque cada
        form dice otra cosa. En la capa nativa el iframe es un contenedor con su title: si
        adentro ya no quedan campos ni botones de avanzar/enviar y hay texto, es la TY.
        Como respaldo, textos de agradecimiento conocidos.
        """
        patrones_eid = _patrones_event_id()
        limite = time.time() + timeout
        while time.time() < limite:
            self._nativo()
            try:
                raiz = self._fuente()
                for e in raiz.iter(f"{_T}StaticText"):
                    t = _texto(e)
                    if any(p in _norm(t) for p in patrones_eid):
                        return "event_id", t
                caja = self._contenedor_del_form(raiz)
                if caja is not None:
                    hay_campos = any(e.tag.replace(_T, "") in ("TextField", "TextView", "SecureTextField")
                                     for e in caja.iter())
                    kw_siguiente, kw_enviar = _kw_botones()
                    hay_enviar = any(e.tag == f"{_T}Button"
                                     and _norm(_texto(e)).startswith(kw_enviar + kw_siguiente)
                                     for e in caja.iter())
                    textos = [_texto(e) for e in caja.iter(f"{_T}StaticText") if _texto(e)]
                    if not hay_campos and not hay_enviar and textos:
                        return "ty", " / ".join(textos[:2])
                for e in raiz.iter(f"{_T}StaticText"):
                    t = _texto(e)
                    if any(m in _norm(t) for m in _TEXTOS_TY):
                        return "ty", t
            finally:
                self._webview()
            time.sleep(1.5)
        self._diagnostico_sin_ty()
        return "nada", ""

    def esperar_recarga(self, timeout: int = 12) -> bool:
        """Tras el error de event_id el form se recarga solo. Igual que la app
        (_esperar_recarga_tras_event_id): esperar a que los campos que se llenaron vuelvan
        a estar vacíos antes de rellenar; si se rellena antes, la recarga borra todo."""
        fin = time.time() + timeout
        time.sleep(1)
        while time.time() < fin:
            self.ir_al_form()
            self._nativo()
            try:
                textos = [c for c in self.escanear()["controles"] if c["tipo"] == "texto"]
                if textos and all(not c["valor"].strip() for c in textos):
                    self.log("    ✓ Formulario recargado tras error event_id")
                    time.sleep(1)  # que terminen de montar dropdowns/validaciones
                    return True
            finally:
                self._webview()
            time.sleep(0.5)
        self.log(f"    ⚠ No se detectó la recarga del formulario en {timeout}s; se rellena igual")
        return False

    def _diagnostico_sin_ty(self):
        """Sin TY: dejar en el log qué muestra el form, para saber si rebotó o si la TY es
        de un tipo que no se reconoce."""
        self._nativo()
        try:
            raiz = self._fuente()
            caja = self._contenedor_del_form(raiz)
            if caja is None:
                self.log(f"    (diagnóstico) no se encontró el contenedor del form '{self.titulo_iframe}'")
                return
            campos = sum(1 for e in caja.iter() if e.tag.replace(_T, "") in ("TextField", "TextView"))
            botones = [_texto(e) for e in caja.iter(f"{_T}Button") if _texto(e)]
            textos = [_texto(e) for e in caja.iter(f"{_T}StaticText")
                      if _texto(e) and self._visible(e)]
            self.log(f"    (diagnóstico) campos de texto: {campos} | botones: {botones[:5]}")
            self.log(f"    (diagnóstico) textos visibles: {[t[:60] for t in textos[:12]]}")
        except Exception as e:
            self.log(f"    (diagnóstico) {str(e).splitlines()[0][:120]}")
        finally:
            self._webview()

    def _contenedor_del_form(self, raiz):
        """El nodo nativo que corresponde al iframe: un Other con el title del iframe."""
        if not self.titulo_iframe:
            return None
        for e in raiz.iter(f"{_T}Other"):
            if (e.get("label") or "").strip() == self.titulo_iframe:
                return e
        return None

    def mostrar_ty(self):
        """Deja la TY a la vista unos segundos para que quede en el video (igual que la app)."""
        self.ir_al_form()
        time.sleep(6)
