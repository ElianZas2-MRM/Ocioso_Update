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


def _enviar_con_la_app(plataforma: str, url: str, form_url: str) -> int:
    """Con --enviar: el MISMO camino que la app. Arma un Excel de 1 lead como "Generar
    Excels" y lo corre con runner.run (motor de la app; iPhone por la capa nativa)."""
    import os
    import pandas as pd
    from openpyxl import load_workbook
    from osocio.paths import TEMPORALES_DIR
    from osocio.providers.browserstack import runner
    from osocio.providers.browserstack.dispositivos import por_cualquier_nombre
    from osocio.utils.data_generator import generar_fila_datos
    from osocio.utils.fixed_field_mapping_store import build_excel_columns_for_country

    pais = _pais_de(url, form_url)
    if not pais:
        _log("No se pudo deducir el país de la URL: --enviar necesita una URL de un mercado.")
        return 1
    columnas = build_excel_columns_for_country(pais)
    datos = generar_fila_datos(pais, device=por_cualquier_nombre(plataforma).token_email)
    fila = [url if c == "URL" else form_url if c == "Formulario" else datos.get(c, "") for c in columnas]
    os.makedirs(TEMPORALES_DIR, exist_ok=True)
    excel = os.path.join(TEMPORALES_DIR, f"_prueba_{pais}_{plataforma}.xlsx")
    pd.DataFrame([fila], columns=columnas).astype(str).to_excel(excel, index=False)
    _log(f"Lead de prueba ({pais}): " + ", ".join(f"{k}={v}" for k, v in zip(columnas, fila)
                                                   if v and k not in ("URL", "Formulario")))

    res = runner.run(pais, plataforma, excel_path=excel, build_name="Prueba de conexion", log_fn=_log)
    _log("\n" + "=" * 60)
    _log(f"RESUMEN: {res['ok']}/{res['total']} OK" + (f" | error: {res['error']}" if res.get("error") else ""))
    if res.get("results_excel") and os.path.exists(res["results_excel"]):
        ws = load_workbook(res["results_excel"]).active
        for h, v in zip([c.value for c in ws[1]], [c.value for c in ws[2]]):
            if h in ("Resultado", "Motivo", "Form coincide", "Datos vs Excel", "TY Page", "Video") and v:
                _log(f"  {h}: {str(v)[:120]}")
        _log(f"  Excel: {res['results_excel']}")
    _log("=" * 60)
    return 0 if res.get("ok") and not res.get("failed") else 1


def _diagnostico_iphone(driver, iframe, paso):
    """Sin envío: qué campos del form ve el iPhone por la capa nativa (no llena nada)."""
    from osocio.providers.browserstack.ios_nativo import FormNativoIOS
    form = FormNativoIOS(driver, iframe, log=_log)
    form.ir_al_form()
    form._nativo()
    try:
        controles = form.escanear()["controles"]
    finally:
        form._webview()
    paso("5. Campos del form vistos en el iPhone (capa nativa)", bool(controles),
         ", ".join(f"{c['tipo']}:{c['etiqueta'] or '?'}" for c in controles))


def correr(plataforma: str, url: str, form_url: str = "", os_version: str = "",
           enviar: bool = False) -> int:
    pasos = []  # (nombre, ok, detalle)

    def paso(nombre, ok, detalle=""):
        pasos.append((nombre, ok, detalle))
        _log(f"  {'OK ' if ok else 'FALLA'} | {nombre}" + (f" — {detalle}" if detalle else ""))

    username, access_key = load_credentials()
    _log(f"Credenciales OK: {username}")
    _log(f"Plataforma: {plataforma} | URL: {url}{' | ENVIA el form' if enviar else ''}\n")
    if enviar:
        return _enviar_con_la_app(plataforma, url, form_url)


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
        iframe, src, total = _buscar_iframe_scrolleando(driver, form_url)
        if iframe:
            como = "el form esperado" if form_url and form_url in src else "otro iframe GM (fallback)"
            detalle = f"{como} | {src[:120]}"
        else:
            detalle = f"{total} iframes en la página, ninguno GM"
        paso("4. Iframe del form encontrado", bool(iframe), detalle)

        if iframe and not puede_entrar_a_iframe(plataforma):
            # iOS Safari bloquea switch_to.frame() a un iframe de otro dominio: el form se
            # maneja por la capa nativa, dentro de la landing.
            try:
                _diagnostico_iphone(driver, iframe, paso)
            except Exception as e:
                paso("Capa nativa del iPhone", False, str(e).splitlines()[0][:200])
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
