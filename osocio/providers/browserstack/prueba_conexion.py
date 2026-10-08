"""
prueba_conexion.py
==================
Fase 1 de la migración a BrowserStack: ¿anda el iPhone?

Abre una landing REAL en producción y recorre los pasos que la app necesita, reportando
cuál anda y cuál no. Sólo envía el form con --enviar (lead REAL, email @mrm.com).

Paso 5: los forms viven en un iframe de otro dominio. Mac y Android entran con
switch_to.frame(); Safari en iPhone lo bloquea (probado en iOS 16 y 17). Como el form
tiene que llenarse SIEMPRE dentro de la landing, en el iPhone se prueba la capa nativa
(NATIVE_APP): detecta cada campo por su etiqueta visible y lo toca y tipea como lo haría
una persona, con los datos del generador de Excels del proyecto (ver ios_nativo.py).

Uso (desde la raíz del proyecto):
    python -m osocio.providers.browserstack.prueba_conexion
    python -m osocio.providers.browserstack.prueba_conexion --enviar
    python -m osocio.providers.browserstack.prueba_conexion --plataforma android
    python -m osocio.providers.browserstack.prueba_conexion --url https://... --form-url secure-developments
"""

import argparse
import random
import sys
import time

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait

from osocio.providers.browserstack.bs_driver import (
    create_bs_driver, fetch_bs_session, load_credentials, mark_bs_status, plataformas,
    puede_entrar_a_iframe,
)

URL_DEFAULT = "https://www.chevrolet.com.co/contacta-un-concesionario"
# El form que esa landing inserta (equivale a la columna B "Formulario" del Excel).
FORM_URL_DEFAULT = "https://secure-developments.com/commonwealth/colombia/gm_front/form/raq"
_GM_IFRAME_KEYWORDS = ("gm_forms", "gm_admin", "gm_front")
_VALOR_PRUEBA = "PruebaBS"


def _log(msg=""):
    # Consola Windows cp1252: que un caracter raro nunca corte la prueba.
    try:
        print(msg, flush=True)
    except UnicodeEncodeError:
        print(str(msg).encode("ascii", "replace").decode("ascii"), flush=True)


def _esperar_carga(driver, timeout=15):
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script("return document.readyState") == "complete")
        return True
    except Exception:
        return False


def _buscar_iframe(driver, form_url: str):
    """Misma prioridad que la app: primero el que matchea la URL esperada, si no uno GM."""
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    srcs = []
    for el in iframes:
        try:
            srcs.append((el, el.get_attribute("src") or ""))
        except Exception:
            continue
    if form_url:
        for el, src in srcs:
            if form_url in src:
                return el, src, len(srcs)
    for el, src in srcs:
        if any(k in src for k in _GM_IFRAME_KEYWORDS):
            return el, src, len(srcs)
    return None, "", len(srcs)


def _buscar_iframe_scrolleando(driver, form_url: str):
    el, src, total = _buscar_iframe(driver, form_url)
    if el:
        return el, src, total
    try:
        alto = driver.execute_script("return document.body.scrollHeight") or 3000
        paso = int((driver.execute_script("return window.innerHeight") or 800) * 0.8)
        y = 0
        while y < alto and not el:
            y += paso
            driver.execute_script(f"window.scrollTo(0, {y});")
            time.sleep(0.5)
            el, src, total = _buscar_iframe(driver, form_url)
    except Exception:
        pass
    return el, src, total


def _primer_input_de_texto(driver):
    for el in driver.find_elements(By.CSS_SELECTOR, "input[type=text], input[type=email], input:not([type])"):
        try:
            if el.is_displayed() and el.is_enabled():
                return el
        except Exception:
            continue
    return None


def _probar_escritura(driver, paso, nombre: str):
    """Escribe un valor de prueba en el primer input de texto, verifica y lo borra."""
    campo = _primer_input_de_texto(driver)
    if not campo:
        paso(nombre, False, "no hay input de texto visible")
        return
    ident = campo.get_attribute("id") or campo.get_attribute("name") or "?"
    try:
        campo.click()
        campo.clear()
        campo.send_keys(_VALOR_PRUEBA)
        valor = campo.get_attribute("value") or ""
        paso(nombre, valor == _VALOR_PRUEBA, f"'{ident}' quedó con '{valor}'")
    except Exception as e:
        paso(nombre, False, f"'{ident}': {str(e).splitlines()[0][:150]}")
    finally:
        try:
            campo.clear()
        except Exception:
            pass


# ── iOS: llenado por la capa nativa, con la lógica de datos de la app ─────────
def _pais_de(url: str, form_url: str) -> str:
    """País del form: por el segmento del form (/commonwealth/colombia/) o el dominio (.com.co)."""
    from osocio.utils.data_generator import PAIS_ABREV
    import urllib.parse
    ruta = (form_url or "").lower()
    for pais in PAIS_ABREV:
        if f"/{pais.lower()}/" in ruta or (pais == "Brasil" and "/brazil/" in ruta):
            return pais
    host = urllib.parse.urlparse(url).hostname or ""
    for pais, abrev in PAIS_ABREV.items():
        if host.endswith("." + abrev):
            return pais
    return ""


def _resolvedor_de_valores(pais: str, url: str, form_url: str):
    """Arma un lead con el generador de Excels y devuelve valor_para(tipo, etiqueta).

    Es la misma cadena que usa la app para decidir qué va en cada campo
    (_infer_data_key + _get_inferred_value), pero con la etiqueta visible del campo en vez
    del id, que en el iPhone no se ve.
    """
    from osocio.utils.data_generator import generar_fila_datos, valor_plausible_por_nombre
    from osocio.providers.lambdatest_mac.lt_excel_reader import LeadRow
    from osocio.core.campos_inferencia import (
        _infer_data_key, _get_inferred_value, _normalize_for_infer,
    )

    fila = generar_fila_datos(pais, device="iphone")
    lead = LeadRow(index=2, public_url=url, secure_url=form_url, data=fila)
    columnas = {_normalize_for_infer(k): v for k, v in fila.items()}
    sorteo_opcionales = {}

    def valor_para(tipo: str, etiqueta: str):
        if tipo == "texto":
            # 1) Columna del Excel con el mismo nombre que la etiqueta (CPF, CEP, Celular...).
            exacto = columnas.get(_normalize_for_infer(etiqueta))
            if exacto:
                return exacto
            # 2) Inferencia de la app. "Sobrenome" contiene "nome": el apellido se
            #    resuelve antes, si no se tomaría como nombre (la app busca por id y no
            #    tiene este problema; acá se busca por etiqueta).
            n = _normalize_for_infer(etiqueta)
            if any(k in n for k in ("sobrenome", "apellido", "lastname", "surname")):
                clave = "lastname"
            else:
                clave = _infer_data_key(etiqueta, etiqueta, etiqueta)
            valor = _get_inferred_value(clave, lead) if clave else ""
            if valor:
                return valor
            # 3) Campo que el lead no conoce: misma regla que la app (README, "Campos que
            #    no están en el Excel"): obligatorio -> valor plausible siempre; opcional ->
            #    50%, decidido una vez por campo. En el iPhone "obligatorio" es el * del label.
            if "*" not in etiqueta:
                if n not in sorteo_opcionales:
                    sorteo_opcionales[n] = random.random() < 0.5
                if not sorteo_opcionales[n]:
                    return None
            return valor_plausible_por_nombre(etiqueta, etiqueta) or "Prueba"
        # Select: la columna del Excel con ese nombre, si trae dato; si no, None = la app
        # elige una opción real al azar (Modelo, Ciudad, Concesionario... vienen vacías).
        valor = columnas.get(_normalize_for_infer(etiqueta)) or ""
        return valor or None

    return fila, valor_para


def _flujo_app(driver, plataforma: str, url: str, form_url: str, pais: str, paso):
    """Android / Mac: acá sí se puede entrar al iframe, así que se usa EXACTAMENTE el
    flujo de la app en LambdaTest (lt_runner._run_single_lead, el mismo que llaman
    run_lt_batch y run_lt_android_batch por cada lead). Sólo cambia el driver."""
    from osocio.utils.data_generator import generar_fila_datos
    from osocio.providers.lambdatest_mac.lt_excel_reader import LeadRow
    from osocio.providers.lambdatest_mac.lt_runner import (
        _get_field_mapping_for_pais, _load_field_dependencies, _load_ids_dinamicos, _run_single_lead,
    )

    fila = generar_fila_datos(pais, device=plataforma)
    lead = LeadRow(index=2, public_url=url, secure_url=form_url, data=fila)
    _log(f"    lead ({pais or 'país desconocido'}): " + ", ".join(
        f"{k}={v}" for k, v in fila.items() if v and k != "Dispositivo"))
    field_mapping = _get_field_mapping_for_pais(pais)
    result = _run_single_lead(
        driver, pais, lead, field_mapping, _load_field_dependencies(), _load_ids_dinamicos(pais),
        is_mobile=False,
        is_android=(plataforma == "android"),
        brasil_doc_type="cpf",
        log=lambda m="": _log(f"    {m}"),
    )
    paso("5. Form encontrado dentro de la landing", bool(result.get("iframe_found")),
         result.get("iframe_url_found", "")[:110])
    paso("6. Enviado", bool(result.get("submitted")), result.get("result_text", "")[:150])
    paso("7. Thank you page", bool(result.get("ty_confirmed")), result.get("motivo", "") or "")
    _log(f"    datos vs Excel: {result.get('datos_vs_excel', '-')}")


def _reintenta_event_id(form_url: str) -> bool:
    """El error de event_id sólo pasa en los forms gm_forms (no en gm_front/gm_frontend):
    sólo ahí se reintenta el envío."""
    return "gm_forms" in (form_url or "").lower()


def _llenar_iphone(driver, iframe, paso, enviar: bool, pais: str, url: str, form_url: str):
    from osocio.providers.browserstack.ios_nativo import FormNativoIOS

    fila, valor_para = _resolvedor_de_valores(pais, url, form_url)
    _log(f"    lead ({pais or 'país desconocido'}): " + ", ".join(
        f"{k}={v}" for k, v in fila.items() if v and k != "Dispositivo"))

    form = FormNativoIOS(driver, iframe, log=_log)
    registro = form.llenar(valor_para)
    if not registro:
        paso("6. Detectar y llenar el form dentro de la landing", False, "no se detectó ningún control")
        return

    def ok(r):
        if r["tipo"] == "texto":
            return r["quedo"] == r["pedido"] or (
                "".join(ch for ch in r["quedo"] if ch.isdigit()) == r["pedido"])
        if r["tipo"] == "select":
            return bool(r["quedo"]) and not r["quedo"].startswith(("no ", "sin ", "ERROR"))
        return r["pedido"] != "marcar" or r["quedo"] in ("marcado", "ya estaba marcado")

    malos = [r for r in registro if not ok(r)]
    resumen = ", ".join(f"{r['etiqueta'] or r['tipo']}" for r in registro)
    paso("6. Detectar y llenar el form dentro de la landing", not malos,
         f"{len(registro) - len(malos)}/{len(registro)} OK | {resumen}"
         + (f" | MAL: {', '.join(r['etiqueta'] or r['tipo'] for r in malos)}" if malos else ""))

    if not enviar:
        _log("    (sin --enviar: no se envía)")
        time.sleep(8)
        return
    paso("7. Enviar", form.tocar_enviar())
    estado, texto = form.esperar_ty()
    if estado == "event_id" and not _reintenta_event_id(form_url):
        paso("8. Thank you page", False, f"Error event_id (sin reintento en este tipo de form): '{texto[:80]}'")
        return
    if estado == "event_id":
        # Igual que la app: un solo reintento. El form se recarga solo, se espera a que
        # quede vacío, se vuelve a llenar todo y se reenvía.
        _log(f"    ↺ Error event_id (intento 1): '{texto[:80]}'. Esperando recarga del formulario...")
        form.esperar_recarga()
        form.llenar(valor_para)
        paso("7b. Reenviar tras error event_id", form.tocar_enviar())
        estado, texto = form.esperar_ty()
        if estado == "event_id":
            paso("8. Thank you page", False,
                 "Error event_id: formulario no pudo enviarse por falta de event_id (2 intentos fallidos)")
            return
    # El texto post envío queda en el log: sirve para saber qué mostró cada form.
    paso("8. Thank you page", estado == "ty", texto[:120] if texto else "el form no fue reemplazado por una TY")
    form.mostrar_ty()


class _Terminado(Exception):
    """El flujo de la app ya hizo todo: saltar los pasos manuales e ir al cierre."""


def correr(plataforma: str, url: str, form_url: str = "", os_version: str = "",
           enviar: bool = False) -> int:
    pasos = []  # (nombre, ok, detalle)

    def paso(nombre, ok, detalle=""):
        pasos.append((nombre, ok, detalle))
        _log(f"  {'OK ' if ok else 'FALLA'} | {nombre}" + (f" — {detalle}" if detalle else ""))

    username, access_key = load_credentials()
    _log(f"Credenciales OK: {username}")
    _log(f"Plataforma: {plataforma} | URL: {url}{' | ENVIA el form' if enviar else ''}\n")


    driver = None
    session_id = ""
    try:
        try:
            driver = create_bs_driver(username, access_key, platform=plataforma,
                                      build_name="Prueba de conexion",
                                      session_name=f"Prueba {plataforma} {os_version}".strip(),
                                      os_version=os_version)
            session_id = driver.session_id
            caps = driver.capabilities or {}
            equipo = caps.get("deviceName") or caps.get("platformName") or caps.get("platform") or "?"
            equipo += f" ({caps.get('platformName', '?')} {caps.get('platformVersion', '')})".rstrip()
            paso("1. Sesión creada", True,
                 f"id {session_id} | {equipo} | {caps.get('browserName', '')} {caps.get('browserVersion', '')}")
        except Exception as e:
            paso("1. Sesión creada", False, str(e).splitlines()[0][:200])
            return 1

        if enviar and puede_entrar_a_iframe(plataforma):
            # Android / Mac con envío: el flujo completo de la app (navega, llena, envía).
            try:
                _flujo_app(driver, plataforma, url, form_url, _pais_de(url, form_url), paso)
            except Exception as e:
                paso("Flujo de la app", False, str(e).splitlines()[0][:200])
            raise _Terminado

        driver.get(url)
        cargo = _esperar_carga(driver)
        paso("2. Landing cargada", cargo, (driver.title or "")[:80])

        # Mismos pasos previos que la app: scroll dinámico y popup de cookies.
        try:
            from osocio.utils.scroll_dinamico import pre_scroll
            pre_scroll(driver, step_wait=0.1, end_wait=0.2)
            paso("3. Scroll dinámico", True)
        except Exception as e:
            paso("3. Scroll dinámico", False, str(e)[:150])
        try:
            from osocio.providers.lambdatest_mac.lt_runner import _handle_cookie_popups
            cerrado = _handle_cookie_popups(driver, log=lambda *_a, **_k: None)
            paso("3b. Popup de cookies", True, "cerrado" if cerrado else "no apareció")
        except Exception as e:
            paso("3b. Popup de cookies", False, str(e)[:150])

        iframe, src, total = _buscar_iframe_scrolleando(driver, form_url)
        if iframe:
            como = "el form esperado" if form_url and form_url in src else "otro iframe GM (fallback)"
            detalle = f"{como} | {src[:120]}"
        else:
            detalle = f"{total} iframes en la página, ninguno GM"
        paso("4. Iframe del form encontrado", bool(iframe), detalle)

        if iframe and not puede_entrar_a_iframe(plataforma):
            # iOS Safari bloquea switch_to.frame() a un iframe de otro dominio, y el form
            # tiene que llenarse SÍ O SÍ dentro de la landing: se llena por la capa nativa.
            try:
                _llenar_iphone(driver, iframe, paso, enviar, _pais_de(url, form_url), url, form_url)
            except Exception as e:
                paso("Llenado iPhone", False, str(e).splitlines()[0][:200])
        elif iframe:
            try:
                driver.execute_script("arguments[0].scrollIntoView({block:'start'});", iframe)
                time.sleep(1)
                driver.switch_to.frame(iframe)
                _esperar_carga(driver, timeout=10)
                dentro = driver.execute_script("return location.href") or ""
                n_campos = len(driver.find_elements(By.CSS_SELECTOR, "input, select, textarea"))
                paso("5. Entrar al iframe (switch_to.frame)", n_campos > 0,
                     f"{n_campos} campos | {dentro[:100]}")
            except Exception as e:
                paso("5. Entrar al iframe (switch_to.frame)", False, str(e).splitlines()[0][:200])
            else:
                _probar_escritura(driver, paso, "6. Escribir en un campo (send_keys)")
                try:
                    driver.switch_to.default_content()
                    paso("7. Volver a la landing", True)
                except Exception as e:
                    paso("7. Volver a la landing", False, str(e)[:150])
    except _Terminado:
        pass
    finally:
        if driver:
            todo_ok = bool(pasos) and all(ok for _n, ok, _d in pasos)
            fallados = ", ".join(n for n, ok, _d in pasos if not ok)
            mark_bs_status(driver, passed=todo_ok, reason=f"Fallaron: {fallados}" if fallados else "Prueba OK")
            try:
                driver.quit()
            except Exception:
                pass

    if session_id:
        _log("\nEsperando que BrowserStack procese el video...")
        time.sleep(5)
        info = fetch_bs_session(session_id, username, access_key)
        paso("9. Link del video disponible", bool(info["video_url"]), info["device"] or "")
        _log(f"  Video: {info['video_url'] or '-'}")

    _log("\n" + "=" * 60)
    _log("RESUMEN")
    for nombre, ok, _d in pasos:
        _log(f"  {'OK   ' if ok else 'FALLA'}  {nombre}")
    _log("=" * 60)
    return 0 if all(ok for _n, ok, _d in pasos) else 1


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Prueba de conexión a BrowserStack")
    parser.add_argument("--plataforma", choices=plataformas(), default="iphone")
    parser.add_argument("--url", default=URL_DEFAULT, help="Landing pública a abrir")
    parser.add_argument("--form-url", default=FORM_URL_DEFAULT,
                        help="URL del form que inserta la landing (columna B del Excel)")
    parser.add_argument("--os-version", default="",
                        help="Probar otra versión de sistema (ej. 17, 18) sin tocar el catálogo")
    parser.add_argument("--enviar", action="store_true",
                        help="Enviar el form (lead REAL en producción). Sin esto no se envía.")
    args = parser.parse_args(argv)
    return correr(args.plataforma, args.url, args.form_url, args.os_version, args.enviar)


if __name__ == "__main__":
    sys.exit(main())
