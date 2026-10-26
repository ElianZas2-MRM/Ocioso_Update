"""
iphone_runner.py
================
El motor de los navegadores (BaseFormFiller), adaptado al iPhone de BrowserStack.

Safari en iPhone no deja entrar al iframe del form (otro dominio), así que el llenado
va por la capa nativa del dispositivo (ios_nativo.FormNativoIOS). Pero TODO lo demás es
el motor de la app, heredado tal cual:
  - lectura del Excel, nombres y carpetas de resultados (t1/t3, CDC/T3), colores,
    "Form coincide", estado de las URLs, motivo, resumen y fail_rows (Reintentar fallidos)
  - el valor de cada campo: _resolve_field_value (Excel), '-' para omitir, IDs dinámicos
    sólo si la celda está vacía, documentos de Brasil normalizados (nunca generados),
    documento de Perú según el tipo elegido, campos no mapeados (obligatorio -> valor
    plausible; opcional -> 50%)
  - el reintento por error de event_id del bucle principal

Lo único que se reemplaza son los métodos que tocan el DOM del iframe, y sólo cuando el
form está en un iframe: si está embebido en la página no hay bloqueo, y llena el motor
normal por DOM, igual que en cualquier navegador.
"""

import random
import time
from typing import Dict, Optional

from osocio.core.campos_inferencia import ids_para_etiqueta
from osocio.core.generic_country_base import GenericCountryBase
from osocio.providers.browserstack.ios_nativo import OMITIR, FormNativoIOS


def _casillas_por_http(url, timeout=15) -> Dict:
    """Radios, checkboxes y <select multiple> del HTML que devuelve el servidor, con el
    mismo formato que FormFillerIphone._JS_CASILLAS. {} si no hay (o no se pudo bajar)."""
    import urllib.request
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        html = urllib.request.urlopen(req, timeout=timeout).read().decode("utf-8", "replace")
    except Exception:
        return {}
    return casillas_del_html(html)


def casillas_del_html(html) -> Dict:
    from html.parser import HTMLParser

    class _Lector(HTMLParser):
        def __init__(self):
            super().__init__()
            self.inputs = []            # [name, tipo, value, id, texto del label que lo envuelve]
            self.labels = {}            # for -> texto
            self.multis = {}            # name -> [[value, texto], ...]
            self._label = None          # [texto, for, índices de inputs adentro]
            self._multi = None
            self._opcion = None

        def handle_starttag(self, tag, attrs):
            a = dict(attrs)
            tipo = (a.get("type") or "").lower()
            if tag == "input" and tipo in ("radio", "checkbox"):
                nombre = (a.get("name") or a.get("id") or "").lower()
                if nombre:
                    self.inputs.append([nombre, tipo, a.get("value") or "", a.get("id") or "", ""])
                    if self._label is not None:
                        self._label[2].append(len(self.inputs) - 1)
            elif tag == "label":
                self._label = ["", a.get("for"), []]
            elif tag == "select" and "multiple" in a:
                self._multi = (a.get("name") or a.get("id") or "").lower()
                self.multis[self._multi] = []
            elif tag == "option" and self._multi and "disabled" not in a and (a.get("value") or "").strip():
                self._opcion = [a.get("value"), ""]

        def handle_endtag(self, tag):
            if tag == "label" and self._label is not None:
                texto = " ".join(self._label[0].split())
                if self._label[1]:
                    self.labels[self._label[1]] = texto
                for i in self._label[2]:
                    self.inputs[i][4] = texto
                self._label = None
            elif tag == "option" and self._opcion is not None:
                self.multis[self._multi].append([self._opcion[0], " ".join(self._opcion[1].split())])
                self._opcion = None
            elif tag == "select":
                self._multi = None

        def handle_data(self, data):
            if self._label is not None:
                self._label[0] += data
            if self._opcion is not None:
                self._opcion[1] += data

    lector = _Lector()
    try:
        lector.feed(html or "")
    except Exception:
        return {}
    out: Dict[str, Dict] = {}
    for nombre, tipo, valor, iid, envuelve in lector.inputs:
        g = out.setdefault(nombre, {"tipo": tipo, "titulo": "", "opciones": []})
        g["opciones"].append({"valor": valor, "texto": (lector.labels.get(iid) if iid else "") or envuelve})
    for nombre, opciones in lector.multis.items():
        if nombre and opciones:
            out[nombre] = {"tipo": "multi", "titulo": "",
                           "opciones": [{"valor": v, "texto": t} for v, t in opciones]}
    return out


class FormFillerIphone(GenericCountryBase):

    _nativo: Optional[FormNativoIOS] = None

    def begin_row_tracking(self):
        super().begin_row_tracking()
        self._nativo = None  # cada fila es otra landing: el iframe se vuelve a buscar
        self._casillas_form = {}
        self._nombres_ios = {}          # (tipo, etiqueta en pantalla) -> name del form
        self._checkbox_ios_marcadas = {}

    # ── traductor name <-> texto (radios, checkboxes, listas múltiples) ──────────
    # La capa nativa ve los textos ("WhatsApp", "¿Tenés un Chevrolet?") pero no el name
    # del campo, que es lo que trae el Excel. Antes de abrir la landing se abre la URL del
    # form sola (ahí sí hay DOM) y se leen los name con sus opciones, tal como están HOY:
    # si el form agrega una opción, aparece sola. Después se llena en la landing como siempre.
    _JS_CASILLAS = r"""
        function txt(e) { return e ? e.textContent.replace(/\s+/g, ' ').trim() : ''; }
        function label(c) {
            return txt((c.id && document.querySelector('label[for="' + c.id + '"]')) || c.closest('label'));
        }
        function titulo(c) {
            var g = c.closest('fieldset, .form-group, [role=group], [role=radiogroup]');
            var t = g && g.querySelector('legend, label:not(.form-check-label)');
            return t && !t.contains(c) ? txt(t) : '';
        }
        var out = {};
        document.querySelectorAll('input[type=radio], input[type=checkbox]').forEach(function (c) {
            var n = (c.getAttribute('name') || c.id || '').toLowerCase();
            if (!n) return;
            var g = out[n] = out[n] || {tipo: c.type, titulo: titulo(c), opciones: []};
            g.opciones.push({valor: c.value || '', texto: label(c)});
        });
        document.querySelectorAll('select[multiple]').forEach(function (s) {
            var n = (s.getAttribute('name') || s.id || '').toLowerCase();
            if (!n) return;
            out[n] = {tipo: 'multi', titulo: txt(s.id && document.querySelector('label[for="' + s.id + '"]')) || titulo(s),
                      opciones: Array.from(s.options).filter(function (o) { return !o.disabled && o.value.trim(); })
                                     .map(function (o) { return {valor: o.value, texto: txt(o)}; })};
        });
        return out;"""

    def _columnas_de_casillas(self):
        """Columnas del Excel con valor que no son campos del mapping: pueden ser el name de
        un radio, de un grupo de checkboxes o de una lista múltiple."""
        from osocio.utils.field_id_aliases import alias_ids_for
        conocidos = {str(c).lower() for c in (getattr(self, "data_columns", None) or [])}
        for fc in self.field_mapping or []:
            fids = fc.get("id")
            for f in (fids if isinstance(fids, list) else [fids]):
                conocidos.add(str(f or "").lower())
                conocidos.update(str(a).lower() for a in alias_ids_for(f))
        pedidas = {**(self.radio_prefs or {}), **{k: "" for k in (self.checkbox_prefs or {})}}
        return [k for k, v in pedidas.items()
                if k not in conocidos and not str(v).lower().startswith(("http", "/content/"))]

    def process_landing_page(self, landing_url, ss_counter, take_screenshot=True):
        form_url = getattr(self, "expected_form_url", "") or ""
        # Una vez por fila (process_landing_page se repite en los reintentos).
        if form_url and not self._casillas_form and self._columnas_de_casillas():
            self._casillas_form = self._leer_casillas_del_form(form_url)
        return super().process_landing_page(landing_url, ss_counter, take_screenshot)

    def _leer_casillas_del_form(self, form_url):
        """{name: {tipo, titulo, opciones: [{valor, texto}]}} del form. Primero por HTTP (los
        forms gm_forms traen los inputs en el HTML): no se navega en el iPhone, que con
        landings pesadas (CL embajadores) se queda sin memoria. Si el HTML no los trae (se
        pintan por JS, ej. gm_front), se abre el form solo en el iPhone y se lee el DOM."""
        casillas = _casillas_por_http(form_url)
        if casillas:
            print(f"🔎 Radios / checkboxes del form (HTML): "
                  f"{ {n: [o['texto'] or o['valor'] for o in g['opciones']] for n, g in casillas.items()} }")
            return casillas
        try:
            self.driver.get(form_url)
            fin = time.time() + 10
            casillas = {}
            while time.time() < fin:   # forms que se pintan por JS: esperar a que aparezcan
                casillas = self.driver.execute_script(self._JS_CASILLAS) or {}
                if casillas:
                    break
                time.sleep(1)
            print(f"🔎 Radios / checkboxes del form: "
                  f"{ {n: [o['texto'] or o['valor'] for o in g['opciones']] for n, g in casillas.items()} }")
            return casillas
        except Exception as e:
            print(f"  ⚠ No se pudieron leer los radios / checkboxes del form: {e}")
            return {}

    def _clave_texto(self, texto):
        return "".join(self._palabras(str(texto or "").replace("*", "")))

    @staticmethod
    def _textos_visibles(opciones):
        """Lo que la pantalla muestra de cada opción: el texto del label o, si el label es
        el mismo para todas (CL envolve: el título del grupo) o no hay, el value."""
        textos = [o["texto"] for o in opciones]
        return [o["texto"] if o["texto"] and textos.count(o["texto"]) == 1 else o["valor"] for o in opciones]

    def _titulo_coincide(self, titulo_form, etiqueta):
        a, b = self._clave_texto(titulo_form), self._clave_texto(etiqueta)
        return bool(a and b and (a == b or a in b or b in a))

    def _grupo_del_form(self, tipo, opciones_pantalla=None, titulo_pantalla="", opcion=""):
        """El name del form que corresponde a lo que se ve en pantalla. Radios: por el
        conjunto de opciones (y el título si hay dos grupos iguales, ej. dos "Sí / No").
        Checkbox: por el texto de su opción (y el título del grupo si se repite)."""
        grupos = {n: g for n, g in (self._casillas_form or {}).items() if g.get("tipo") == tipo}
        if opcion:
            o = self._clave_texto(opcion)
            cands = [n for n, g in grupos.items()
                     if any(self._clave_texto(t) == o for t in self._textos_visibles(g["opciones"]))]
        elif opciones_pantalla is not None:
            # Todas las opciones de la pantalla están en el grupo. Una pregunta con opciones
            # largas, una debajo de otra (PE libro de reclamaciones), llega en pedazos de una
            # opción: cada pedazo se reconoce igual.
            pant = {self._clave_texto(t) for t in opciones_pantalla} - {""}
            cands = [n for n, g in grupos.items()
                     if pant and pant <= {self._clave_texto(t) for t in self._textos_visibles(g["opciones"])}]
        else:
            cands = list(grupos)
        if len(cands) > 1:
            cands = [n for n in cands if self._titulo_coincide(grupos[n].get("titulo"), titulo_pantalla)]
        return cands[0] if len(cands) == 1 else None

    def _radio_pedido_ios(self, etiqueta, opciones_pantalla):
        """La opción (texto en pantalla) que pide el Excel para esta pregunta de radio, OMITIR,
        o None (al azar)."""
        n = self._grupo_del_form("radio", opciones_pantalla, etiqueta)
        if n is None:
            return None
        self._nombres_ios[("radio", etiqueta)] = n
        pedido = (self.radio_prefs or {}).get(n)
        if not pedido:
            return None
        if self._pide_omitir_campo(n, pedido):
            return OMITIR
        opciones = self._casillas_form[n]["opciones"]
        marcar, _ = self.opciones_a_marcar([pedido], [(o["valor"], o["texto"]) for o in opciones])
        if not marcar:
            self._avisar_opciones_no_encontradas(n, [pedido])
            return None
        elegida = self._textos_visibles(opciones)[min(marcar)]
        if self._clave_texto(elegida) not in {self._clave_texto(t) for t in opciones_pantalla}:
            # Pedazo de la pregunta sin la opción pedida: no se toca (se marca en el suyo).
            print(f"      radio '{n}': '{elegida}' no está en este pedazo {opciones_pantalla}; no se toca")
            return OMITIR
        return elegida

    def _preferencia_checkbox_ios(self, c):
        """True / False según el Excel (SI/NO, o la lista de opciones de su grupo); None si
        el Excel no dice nada de este checkbox."""
        n = self._grupo_del_form("checkbox", titulo_pantalla=c.get("grupo", ""), opcion=c["etiqueta"])
        if n is None:
            return None
        lista = (self.checkbox_opciones or {}).get(n)
        opciones = self._casillas_form[n]["opciones"]
        if lista is not None and self._pide_omitir_campo(n, ",".join(lista)):
            return c["valor"] == "1"      # '-': se deja como está
        if lista is not None and len(opciones) > 1:
            marcar, faltan = self.opciones_a_marcar(lista, [(o["valor"], o["texto"]) for o in opciones])
            self._avisar_opciones_no_encontradas(n, faltan)
            o = self._clave_texto(c["etiqueta"])
            visibles = self._textos_visibles(opciones)
            si = any(self._clave_texto(visibles[i]) == o for i in marcar)
            elegidas = self._checkbox_ios_marcadas.setdefault(n, [])
            if si and c["etiqueta"] not in elegidas:
                elegidas.append(c["etiqueta"])
            return si
        return (self.checkbox_prefs or {}).get(n)

    def _opciones_multi_ios(self, c):
        """Las filas que pide el Excel en una lista múltiple (por su orden en el form) y
        cuántas opciones tiene hoy; None si el form no tiene una lista que coincida."""
        n = self._grupo_del_form("multi", titulo_pantalla=c["etiqueta"])
        if n is None:
            return None
        self._nombres_ios[("multi", c["etiqueta"])] = n
        opciones = self._casillas_form[n]["opciones"]
        plan = {"total": len(opciones), "indices": None}
        lista = (self.checkbox_opciones or {}).get(n)
        if lista:
            if self._pide_omitir_campo(n, ",".join(lista)):
                return {"omitir": True}
            marcar, faltan = self.opciones_a_marcar(lista, [(o["valor"], o["texto"]) for o in opciones])
            self._avisar_opciones_no_encontradas(n, faltan)
            if marcar:
                plan["indices"] = sorted(marcar)
                plan["pedido"] = " | ".join(self._textos_visibles(opciones)[i] for i in plan["indices"])
        return plan

    # ── entrar al form: quedarse en la landing ───────────────────────────────
    def _entrar_al_iframe(self, iframe):
        """No se entra (Safari lo bloquea): el form se maneja desde la capa nativa."""
        self._nativo = FormNativoIOS(self.driver, iframe, log=print)

    def wait_for_form_ready_in_iframe(self, timeout=3):
        return True if self._nativo else super().wait_for_form_ready_in_iframe(timeout)

    def _resolve_submit_button(self, wait_seconds=3):
        # Con el form en iframe no hay envío vacío inicial: su único fin es la captura de
        # errores, y acá no hay capturas (la evidencia es el video).
        return (None, None) if self._nativo else super()._resolve_submit_button(wait_seconds)

    def _detectar_checkboxes_con_error(self):
        return None if self._nativo else super()._detectar_checkboxes_con_error()

    def _verificar_cta_eletricos_br(self, landing_url, form_url):
        return None if self._nativo else super()._verificar_cta_eletricos_br(landing_url, form_url)

    def _seguir_cta(self, cta):
        """iPhone: un CTA que es un link se sigue abriendo su href. Con el click, en landings
        pesadas (BR sonic/acessorios, después del pre-scroll) Safari se quedaba sin memoria
        al navegar y se caía la sesión. El href es el mismo destino: igual se verifica que
        el CTA lleve a la página correcta."""
        href = cta.get_attribute("href") or ""
        if not href:
            return super()._seguir_cta(cta)
        self.driver.get(href)

    def _maybe_click_raq_cta(self, *urls):
        """RAQ de Brasil: antes del form hay una pantalla con cards y el CTA "Solicitar
        Contato" (#contact-by-form) que hay que tocar. En los navegadores se clickea por DOM
        adentro del iframe; en el iPhone se toca por la capa nativa, con las mismas
        condiciones (Brasil + RAQ, no eléctricos)."""
        if not self._nativo:
            return super()._maybe_click_raq_cta(*urls)
        blob = " ".join((u or "").lower() for u in urls)
        if str(self.config.get("pais", "")).lower() not in ("brasil", "brazil", "br"):
            return
        if "raq" not in blob or "eletricos" in blob:
            return
        if not ("solicitar-contato" in blob or "gm_forms/raq" in blob or "raq-revamp" in blob):
            return
        if self._nativo.tocar_texto("Solicitar Contato"):
            print("Tocado 'Solicitar Contato' (RAQ Brasil)")
        else:
            print("⚠ No se encontró el CTA 'Solicitar Contato' del RAQ de Brasil")

    def _rellenar_si_form_se_reseteo(self, form_data, es_libro_reclamaciones=False):
        if self._nativo:
            return None
        return super()._rellenar_si_form_se_reseteo(form_data, es_libro_reclamaciones)

    def _refill_brasil_doc_sendkeys(self, form_data) -> bool:
        return False if self._nativo else super()._refill_brasil_doc_sendkeys(form_data)

    # ── llenado ──────────────────────────────────────────────────────────────
    def fill_form_fields(self, form_data):
        """Llena el form dentro de la landing con los valores que decide la app."""
        if not self._nativo:
            return super().fill_form_fields(form_data)  # form sin iframe: DOM normal
        self._current_form_data = form_data
        self._snapshot_ios: Dict[str, Dict] = {}
        self._tipo_doc_ios = ""
        def valor_para(tipo, etiqueta, opciones=None):
            if tipo == "radio":
                return self._radio_pedido_ios(etiqueta, opciones or [])
            valor = self._valor_para_etiqueta(form_data, tipo, etiqueta)
            # Nunca "test" en un campo de texto (IDs Dinámicos trae "Este texto es un
            # test"): lo mismo que _fill_and_dispatch en los navegadores.
            if tipo == "texto" and isinstance(valor, str):
                valor = self._sanitize_field_value(valor)
            return valor

        self._nativo.elegir_opcion = self._elegir_opcion_ios
        if self._casillas_form:
            self._nativo.preferencia_checkbox = self._preferencia_checkbox_ios
            self._nativo.opciones_multi = self._opciones_multi_ios
        self._rotacion_excluir = {}
        registro = self._nativo.llenar(valor_para, al_llenar=self._registrar_campo_ios)
        self._evitar_concesionario_repetido_ios(valor_para)
        # "Términos" deshabilitado hasta que todo es válido: si no se dejó marcar, algún
        # opcional que quedó vacío por el 50% en realidad es obligatorio por JS (en la capa
        # nativa no se ve), o un checkbox lo habilita (acdelco Chile: `terms` está disabled
        # hasta marcar `product_info`, igual que la segunda pasada de checkboxes en Chrome).
        # Se completan los opcionales, se marcan los demás checkboxes y se vuelve a recorrer.
        if any(r["tipo"] == "checkbox" and r["quedo"] == "sigue en 0" for r in registro):
            print("  'Términos' no se dejó marcar: se completan también los opcionales vacíos "
                  "y los demás checkboxes")
            for clave in self._decision_opcionales:
                if clave.startswith("ios::"):
                    self._decision_opcionales[clave] = True
            self._nativo.marcar_todos = True
            try:
                self._nativo.llenar(valor_para, al_llenar=self._registrar_campo_ios)
            finally:
                self._nativo.marcar_todos = False
        for n, marcadas in self._checkbox_ios_marcadas.items():
            self._record_field_value(n, " | ".join(marcadas))
        return self.screenshot_manager.fname("form", "completado", 1)

    # ── rotación del día (región / ciudad / concesionario) ──────────────────────
    def _elegir_opcion_ios(self, etiqueta, opciones):
        """La opción al azar de un select del iPhone: región, ciudad y concesionario rotan
        (las menos usadas hoy); el resto, None (al azar común)."""
        from osocio.utils.rotacion_dealers import tipo_de_campo
        tipo = tipo_de_campo(etiqueta)
        if not tipo or not opciones:
            return None
        excluir = (getattr(self, "_rotacion_excluir", None) or {}).get(tipo, ())
        i = self._rotacion().elegir(self._pais_rotacion(), tipo, opciones, excluir=excluir)
        self._rotacion_anotar(tipo, etiqueta, opciones[i], al_azar=True)
        return opciones[i]

    def _evitar_concesionario_repetido_ios(self, valor_para):
        """Como _evitar_concesionario_repetido de los navegadores: si el concesionario al azar
        ya se usó hoy, se vuelve a elegir la ciudad (y si hace falta la región) y se completa
        de nuevo lo que dependa de ella, hasta dar con uno sin usar."""
        from osocio.utils.rotacion_dealers import CIUDAD, DEALER, REGION
        fila = getattr(self, "_rotacion_fila", None) or {}
        rot, pais = self._rotacion(), self._pais_rotacion()

        def repetido():
            d = fila.get(DEALER)
            return bool(d and d["al_azar"] and rot.veces(pais, DEALER, d["texto"]) > 0)

        if not repetido():
            return
        for tipo_padre in (CIUDAD, REGION):
            padre = fila.get(tipo_padre)
            if not padre or not padre["al_azar"]:
                break
            probados = {padre["texto"]}
            for _ in range(5):
                self._rotacion_excluir = {tipo_padre: set(probados)}
                try:
                    nuevo = self._nativo.elegir_de_nuevo(padre["id"])
                finally:
                    self._rotacion_excluir = {}
                if not nuevo or nuevo in probados:
                    break
                probados.add(nuevo)
                # Lo que dependía (ciudad / concesionario) quedó vacío: se completa de nuevo.
                self._nativo.llenar(valor_para, al_llenar=self._registrar_campo_ios)
                if not repetido():
                    print(f"🔁 Concesionario sin usar hoy: {fila[DEALER]['texto']}")
                    return
        d = fila.get(DEALER) or {}
        self._aviso_rotacion = (f"concesionario repetido hoy: {d.get('texto', '?')} "
                                "(ya se usaron hoy todos los que ofrece el form, o es el único)")
        print(f"⚠️ {self._aviso_rotacion}")

    def _config_para_ids(self, ids):
        """El field_config del mapping del país que corresponde a esos ids (con alias)."""
        from osocio.utils.field_id_aliases import alias_ids_for
        buscados = set()
        for i in ids:
            buscados.add(i)
            buscados.update(alias_ids_for(i))
        for fc in self.field_mapping or []:
            fids = fc.get("id")
            fids = fids if isinstance(fids, list) else [fids]
            if any(str(f or "") in buscados for f in fids):
                return fc
        return None

    def _id_principal(self, ids, fc):
        if fc is not None:
            fid = fc.get("id")
            return str(fid[0] if isinstance(fid, list) else fid)
        return ids[0] if ids else ""

    def _valor_para_etiqueta(self, form_data, tipo, etiqueta):
        """El mismo camino de decisión que fill_form_fields, partiendo de la etiqueta."""
        if tipo == "radio":
            return None   # lo resuelve _radio_pedido_ios con las opciones de la pantalla
        ids = ids_para_etiqueta(etiqueta)
        fc = self._config_para_ids(ids) if ids else None
        fid = self._id_principal(ids, fc) or etiqueta
        valor = self._resolve_field_value(form_data, fc) if fc is not None else ""
        valor = "" if valor is None else str(valor).strip()

        if self._pide_omitir_campo(fid, valor):
            return OMITIR

        if tipo == "select":
            # Celda vacía o placeholder -> una opción real al azar (lo resuelve el motor nativo).
            return None if (not valor or self._is_placeholder_text(valor)) else valor

        # Texto: un ID dinámico sólo si la celda vino vacía.
        if not valor:
            ids_din = self._cargar_ids_dinamicos()
            for cand in [fid, *ids]:
                if self._usar_valor_fijo(cand, valor, ids_din):
                    valores = self._resolve_dynamic_id_values(ids_din[cand])
                    if valores:
                        valor = str(valores[0])
                        break
        # Brasil: el documento sale del Excel, normalizado (nunca generado).
        if str(self.config.get("pais", "")).lower() in ("brasil", "brazil", "br"):
            tipo_doc = self._tipo_documento_brasil(*(list(ids) + [fid]))
            if tipo_doc:
                valor = self._normalizar_documento_brasil(tipo_doc, valor)
        # Perú: el documento se ajusta al tipo elegido antes en el form.
        if fid in ("ci", "document") and self._tipo_doc_ios and \
                str(self.config.get("pais", "")).lower() in ("peru", "pe"):
            valor = self._sanitize_peru_document(self._tipo_doc_ios, valor)

        if valor:
            return valor
        if fc is not None:
            return None  # campo mapeado sin dato: queda vacío, como en los navegadores
        # Campo que el mapping no conoce: obligatorio -> valor plausible; opcional -> 50%,
        # decidido una vez por campo y por fila (README, "Campos que no están en el Excel").
        if "*" not in etiqueta:
            clave = f"ios::{etiqueta}"
            if clave not in self._decision_opcionales:
                self._decision_opcionales[clave] = random.random() < 0.5
            if not self._decision_opcionales[clave]:
                return None
        from osocio.utils.data_generator import valor_plausible_por_nombre
        return valor_plausible_por_nombre(etiqueta, etiqueta) or "Prueba"

    def _registrar_campo_ios(self, reg):
        """Lo que quedó en cada campo: tracking del Excel y foto para 'Datos vs Excel'."""
        quedo = reg.get("quedo") or ""
        if quedo.startswith(("(", "no ", "sin ", "ERROR")) or reg.get("tipo") == "checkbox":
            return
        # Radio / lista múltiple: con el name del form, si el traductor lo reconoció.
        nombre = self._nombres_ios.get((reg.get("tipo"), reg.get("etiqueta", "")))
        if nombre:
            self._record_field_value(nombre, quedo)
            return
        if reg.get("tipo") == "radio":
            return
        ids = ids_para_etiqueta(reg.get("etiqueta", ""))
        fc = self._config_para_ids(ids) if ids else None
        fid = self._id_principal(ids, fc) or reg.get("etiqueta", "")
        self._record_field_value(fid, quedo)
        # Región / ciudad / concesionario que pidió el Excel: manda, pero cuenta como usado
        # hoy (los al azar ya los anotó _elegir_opcion_ios).
        pedido = str(reg.get("pedido") or "")
        if reg.get("tipo") == "select" and pedido and not pedido.startswith("(") and pedido != "-":
            from osocio.utils.rotacion_dealers import _norm, tipo_de_campo
            if _norm(pedido) in _norm(quedo) or _norm(quedo) in _norm(pedido):
                self._rotacion_anotar(tipo_de_campo(reg.get("etiqueta", "")), reg.get("etiqueta", ""),
                                      quedo, al_azar=False)
        if "document-type" in ids:
            self._tipo_doc_ios = quedo
        entrada = {"id": fid, "realId": fid, "name": fid, "value": quedo, "text": quedo}
        for clave in {fid, *ids}:
            self._snapshot_ios[clave] = entrada

    def _snapshot_form_state(self):
        """'Datos vs Excel' compara contra lo que MUESTRA la pantalla del iPhone."""
        if not self._nativo:
            return super()._snapshot_form_state()
        return dict(getattr(self, "_snapshot_ios", {}) or {})

    # ── envío ────────────────────────────────────────────────────────────────
    def submit_and_verify_form(self, current_ss_number, expected_form_url):
        if bool(self.config.get("solo_verificar_visual", False)
                or self.config.get("no_enviar_lead", False)
                or self.config.get("chequeo_campos", False)):
            return "PASS (Verificación visual)", "Verificación visual"
        nativo = self._nativo
        if nativo is None:
            return super().submit_and_verify_form(current_ss_number, expected_form_url)
        try:
            self._sync_tracked_with_dom_before_submit(getattr(self, "_current_form_data", {}))
        except Exception as e:
            print(f"  ⚠ Datos vs Excel: {e}")
        if not nativo.tocar_enviar():
            return "Error al enviar: no se encontró el botón Enviar", None
        self._marcar_checkboxes_que_pide_el_form(nativo)
        estado, texto = nativo.esperar_ty()
        if estado == "ty":
            print(f"✓ TY: {texto}")
            nativo.mostrar_ty()
            self._ty_cta = "-"
            return "Lead enviado correctamente", self.screenshot_manager.fname("form", "typage",
                                                                                current_ss_number)
        if estado == "event_id":
            print(f"  ⚠ Error event_id: {texto}")
            return "ERROR_EVENT_ID: Error de servidor al envío - formulario recargado automáticamente", None
        return "Formulario sigue visible (TY Page no detectada)", None

    def _marcar_checkboxes_que_pide_el_form(self, nativo) -> bool:
        """Checkboxes obligatorios por JS (la lógica de la app, #27, en el iPhone).

        En Chrome se detectan con el "enviar vacío" del principio. En el iPhone se hace al
        revés, para no arriesgar un envío vacío en un form sin obligatorios: se llena todo,
        se toca Enviar y, si el form frena con un aviso debajo de un checkbox o de su grupo,
        se marcan esos (todas las opciones del grupo, como en Chrome) y se reenvía una vez.
        Si nada frena el envío, es que no eran obligatorios. Lo que pide el Excel para un
        checkbox (SI/NO o la lista de su grupo) se respeta también acá (FormNativoIOS.marcar).
        """
        time.sleep(2)
        try:
            rechazados = nativo.checkboxes_rechazados()
        except Exception as e:
            print(f"  ⚠ No se pudo revisar si el form pidió checkboxes: {e}")
            return False
        if not rechazados:
            return False
        print("  ↺ El form pide marcar: " + ", ".join(
            f"{c.get('grupo') + ' > ' if c.get('grupo') else ''}{c['etiqueta'] or '?'}" for c in rechazados))
        for c in rechazados:
            print(f"    · checkbox {c['etiqueta'] or '?':28} -> {nativo.marcar(c)}")
        # Hay forms (Chile) que dejan "términos" deshabilitado hasta que todo lo demás es
        # válido: el toque del llenado no entró. Ahora que el resto está, se vuelve a marcar.
        try:
            for c in nativo.checkboxes_pendientes():
                print(f"    · checkbox {c['etiqueta'] or '?':28} -> {nativo.marcar(c)}")
        except Exception as e:
            print(f"  ⚠ No se pudieron revisar los checkboxes de términos: {e}")
        return nativo.tocar_enviar()

    def _esperar_recarga_tras_event_id(self, use_iframe, expected_form_url, timeout=12):
        if self._nativo is None:
            return super()._esperar_recarga_tras_event_id(use_iframe, expected_form_url, timeout)
        return self._nativo.esperar_recarga(timeout)
