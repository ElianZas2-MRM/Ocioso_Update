"""
dealer_comparator_runner.py — Lógica de negocio de la pestaña "Comparador Dealers".

Lee un Excel de dealers esperados (fila de encabezado y columnas configurables porque
el Excel que mandan varía de país a país), navega región→ciudad→dealer en un formulario
real via Selenium (reusando BrowserManager), compara contra el Excel filtrado y arma un
reporte PASS/FAIL/EXTRA/MISSING con export a Excel (openpyxl) y capturas opcionales
(ScreenshotManager, con banner de URL) empaquetadas en un ZIP.
"""
import os
import re
import time
import unicodedata
import zipfile
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import column_index_from_string

from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait

from core.browser_manager import BrowserManager
from core.country_configs import COUNTRY_CONFIGS
from core.screenshot_manager import ScreenshotManager

try:
    from utils.paths import RESULTS_DIR
except Exception:
    RESULTS_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "resultados")


DEFAULT_SELECT_IDS = {"region": "region", "city": "city", "dealer": "dealer"}

_PASS_FILL = PatternFill(fill_type="solid", fgColor="00C6EFCE")
_FAIL_FILL = PatternFill(fill_type="solid", fgColor="00FFC7CE")
_EXTRA_FILL = PatternFill(fill_type="solid", fgColor="00FFEB9C")
_MISSING_FILL = PatternFill(fill_type="solid", fgColor="00D9D2E9")
_DUPLICATE_FILL = PatternFill(fill_type="solid", fgColor="00FFCC99")
_HEADER_FILL = PatternFill(fill_type="solid", fgColor="007D4E9F")
_HEADER_FONT = Font(color="00FFFFFF", bold=True)


class StopRequested(Exception):
    """Señal interna para cortar la corrida cuando el usuario aprieta Detener."""


# ──────────────────────────────────────────────────────────────────────────────
# Config por país
# ──────────────────────────────────────────────────────────────────────────────
def get_country_level_defaults(country_name):
    """A partir de core/country_configs.py, infiere qué niveles (región/ciudad/dealer)
    tiene mapeados el país, para usar como default editable en la config persistida."""
    config = COUNTRY_CONFIGS.get(country_name, {})
    ids = {fm.get("id") for fm in config.get("field_mapping", []) if isinstance(fm.get("id"), str)}
    return {
        "has_region": "region" in ids,
        "has_city": "city" in ids,
        "has_dealer": "dealer" in ids,
    }


# ──────────────────────────────────────────────────────────────────────────────
# Normalización de texto (acentos / mayúsculas), igual criterio que el resto de la app
# ──────────────────────────────────────────────────────────────────────────────
def normalize_text(value):
    text = str(value or "").strip().upper()
    text = unicodedata.normalize("NFD", text)
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", text)


_PLACEHOLDER_KEYWORDS = (
    "SELECCION", "SELECIONE", "SELECIONAR", "ESCOLHA", "ESCOLHER",
    "ELIJA", "ELEGIR", "OPCION", "OPCIONES", "SELECT", "CHOOSE", "REQUIRED",
)


def _is_placeholder(text):
    normalized = normalize_text(text)
    if not normalized:
        return True
    return any(k in normalized for k in _PLACEHOLDER_KEYWORDS)


# ──────────────────────────────────────────────────────────────────────────────
# Lectura de Excel con fila de encabezado configurable
# ──────────────────────────────────────────────────────────────────────────────
def read_excel_rows(file_path, header_row=1, sheet_name=None):
    """Lee un Excel con la fila de encabezado en `header_row` (1-indexado, como lo ve el usuario en Excel).

    Devuelve (headers, rows). Cada row es un dict con:
      - clave = texto del encabezado (si existe y no está vacío)
      - clave = "_c0", "_c1", ... (índice de columna, para resolver por letra A/B/K/...)
      - "__row__" = número de fila real en el Excel (1-indexado)
    """
    wb = load_workbook(file_path, data_only=True, read_only=True)
    try:
        sheet = wb[sheet_name] if sheet_name else wb.worksheets[0]
        all_rows = list(sheet.iter_rows(values_only=True))
    finally:
        wb.close()

    if header_row < 1 or header_row > len(all_rows):
        raise ValueError(f"La fila de encabezado ({header_row}) está fuera de rango del Excel.")

    header_cells = all_rows[header_row - 1]
    headers = [str(h).strip() if h is not None else "" for h in header_cells]

    rows = []
    for i in range(header_row, len(all_rows)):
        raw = all_rows[i]
        if raw is None or all(v is None or str(v).strip() == "" for v in raw):
            continue
        row = {"__row__": i + 1}
        for idx, header in enumerate(headers):
            val = raw[idx] if idx < len(raw) else None
            val = "" if val is None else str(val).strip()
            row[f"_c{idx}"] = val
            if header:
                row[header] = val
        rows.append(row)

    return headers, rows


def resolve_column(headers, input_name):
    """Resuelve un nombre de columna ingresado por el usuario a la clave real del row dict.
    Soporta: letra de columna estilo Excel (A, B, K...), nombre exacto, o coincidencia parcial
    (sin distinguir mayúsculas/acentos), igual criterio que el bookmarklet original."""
    input_name = (input_name or "").strip()
    if not input_name:
        return None

    if re.fullmatch(r"[A-Za-z]{1,3}", input_name) and input_name not in headers:
        try:
            idx = column_index_from_string(input_name.upper()) - 1
            if idx >= 0:
                return f"_c{idx}"
        except Exception:
            pass

    if input_name in headers:
        return input_name

    normalized_input = normalize_text(input_name)
    for h in headers:
        if h and normalize_text(h) == normalized_input:
            return h
    for h in headers:
        if h and (normalized_input in normalize_text(h) or normalize_text(h) in normalized_input):
            return h
    return None


def filter_rows(rows, filter_key, filter_value, mode="include"):
    """Filtra filas por columna+valor. mode='include' deja solo las que matchean el valor,
    'exclude' saca las que matchean (para el caso de "chequear todo MENOS X")."""
    if not filter_key or filter_value in (None, ""):
        return list(rows)
    target = normalize_text(filter_value)

    def _matches(row):
        return normalize_text(row.get(filter_key, "")) == target

    if mode == "exclude":
        return [r for r in rows if not _matches(r)]
    return [r for r in rows if _matches(r)]


# ──────────────────────────────────────────────────────────────────────────────
# Navegación de selects dependientes (región → ciudad → dealer) en el form real
# ──────────────────────────────────────────────────────────────────────────────
def _get_select(driver, element_id, timeout=10):
    try:
        element = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.ID, element_id))
        )
        if (element.tag_name or "").lower() != "select":
            return None
        return element
    except Exception:
        return None


def _valid_options(select_element):
    try:
        return [o for o in select_element.find_elements(By.TAG_NAME, "option") if o.get_attribute("value")]
    except Exception:
        return []


def _find_option_by_text(select_element, text):
    """Busca la <option> cuyo texto matchea (exacto primero, luego parcial normalizado)."""
    target = normalize_text(text)
    if not target:
        return None
    options = _valid_options(select_element)
    for opt in options:
        if normalize_text(opt.text) == target:
            return opt
    for opt in options:
        opt_norm = normalize_text(opt.text)
        if opt_norm and (target in opt_norm or opt_norm in target):
            return opt
    return None


def _select_option(driver, select_element, option_element):
    Select(select_element).select_by_value(option_element.get_attribute("value"))
    driver.execute_script(
        "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));"
        "arguments[0].dispatchEvent(new Event('input', {bubbles:true}));",
        select_element,
    )


FAST_TIMEOUT = 0.6  # primer intento: rápido, alcanza en la gran mayoría de los forms
MAX_TIMEOUT = 1.2    # segundo intento (solo si el primero no alcanzó): tope máximo
                      # (ej. el select de ciudad se puebla más lento en departamentos con
                      # muchas opciones, como Montevideo)


def _wait_options_loaded_once(driver, element_id, timeout, poll=0.1):
    deadline = time.time() + timeout
    element = None
    while time.time() < deadline:
        element = _get_select(driver, element_id, timeout=0.3)
        if element is not None and any(not _is_placeholder(o.text) for o in _valid_options(element)):
            return element
        time.sleep(poll)
    return element


def _wait_options_loaded(driver, element_id):
    """Espera a que un <select> hijo tenga opciones válidas (no placeholder) tras disparar el padre.
    Dos intentos: rápido (FAST_TIMEOUT) y, solo si no alcanzó, uno más largo hasta MAX_TIMEOUT total."""
    element = _wait_options_loaded_once(driver, element_id, FAST_TIMEOUT)
    if element is not None and any(not _is_placeholder(o.text) for o in _valid_options(element)):
        return element
    remaining = max(0.0, MAX_TIMEOUT - FAST_TIMEOUT)
    if remaining <= 0:
        return element
    return _wait_options_loaded_once(driver, element_id, remaining)


def _read_form_field_value(driver, field_id):
    """Lee el valor actual de un campo del form por id: texto de la opción elegida si es
    <select>, o el atributo value si es un input/textarea."""
    try:
        el = driver.find_element(By.ID, field_id)
    except Exception:
        return None
    try:
        tag = (el.tag_name or "").lower()
    except Exception:
        return None
    if tag == "select":
        try:
            return (Select(el).first_selected_option.text or "").strip()
        except Exception:
            return None
    try:
        return (el.get_attribute("value") or "").strip()
    except Exception:
        return None


def list_model_options(driver, field_id):
    """Devuelve los textos de todas las opciones válidas (no placeholder) del <select> de
    modelos, tal como están hoy en el form real. Usado para el modo 'Todos los modelos'."""
    el = _get_select(driver, field_id)
    if el is None:
        return []
    return [o.text.strip() for o in _valid_options(el) if not _is_placeholder(o.text)]


def _select_model(driver, field_id, model_text):
    el = _get_select(driver, field_id)
    if el is None:
        return False
    opt = _find_option_by_text(el, model_text)
    if opt is None:
        return False
    _select_option(driver, el, opt)
    return True


def _check_stop(stop_flag):
    if stop_flag is not None and stop_flag.is_set():
        raise StopRequested()


def _wait_document_ready(driver, timeout=30):
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )


def _find_form_iframe(driver, expected_form_url):
    expected_form_url = (expected_form_url or "").strip()
    if not expected_form_url:
        return None
    try:
        for iframe in driver.find_elements(By.TAG_NAME, "iframe"):
            src = iframe.get_attribute("src") or ""
            if expected_form_url in src:
                return iframe
    except Exception:
        pass
    return None


def handle_cookie_popups(driver):
    """Cierra popups de cookies/legales (mismo criterio que BaseFormFiller.handle_cookie_popups,
    incluye el caso puntual de GM/Chevrolet vía <gb-legal-notification>)."""
    try:
        popups = driver.find_elements(By.TAG_NAME, "gb-legal-notification")
        if popups:
            close_buttons = driver.find_elements(By.CSS_SELECTOR, ".close-btn.js-close-icon.silent-consent")
            for btn in close_buttons:
                if btn.is_displayed():
                    driver.execute_script("arguments[0].click();", btn)
                    time.sleep(1)
                    return True
    except Exception:
        pass

    selectors = [
        "button[onclick*='cookie']",
        "button[class*='cookie-accept']",
        "button[id*='cookie-accept']",
        ".cookie-accept",
        "#cookie-accept",
        ".js-close-icon",
        ".silent-consent",
        ".close-btn",
    ]
    for selector in selectors:
        try:
            for element in driver.find_elements(By.CSS_SELECTOR, selector):
                if element.is_displayed():
                    driver.execute_script("arguments[0].click();", element)
                    time.sleep(1)
                    return True
        except Exception:
            continue
    return False


def open_target(driver, url_mode, landing_url="", form_url="", page_ready_timeout=30):
    """Abre el form real, igual criterio que 'Envío de Leads':
    - url_mode='landing_form': abre la landing, cierra cookies y, si el form está embebido en un
      iframe, cambia de contexto a ese iframe. Si no hay iframe reconocible, navega directo al form_url.
    - url_mode='solo_forms': abre form_url directamente (sin landing).
    En ambos casos cierra popups de cookies antes y después de entrar al form.
    Devuelve True si terminó sobre el contexto del form (o su iframe)."""
    if url_mode == "solo_forms" or not landing_url:
        driver.get(form_url)
        _wait_document_ready(driver, page_ready_timeout)
        handle_cookie_popups(driver)
        return True

    driver.get(landing_url)
    _wait_document_ready(driver, page_ready_timeout)
    handle_cookie_popups(driver)

    iframe = _find_form_iframe(driver, form_url)
    if iframe is not None:
        driver.switch_to.frame(iframe)
        handle_cookie_popups(driver)
        return True

    if form_url:
        driver.get(form_url)
        _wait_document_ready(driver, page_ready_timeout)
        handle_cookie_popups(driver)
    return True


# ──────────────────────────────────────────────────────────────────────────────
# Comparación principal: recorre las filas esperadas del Excel contra el form real
# ──────────────────────────────────────────────────────────────────────────────
def compare_dealers(
    driver,
    rows,
    column_map,
    level_ids=None,
    has_region=True,
    has_city=True,
    chk_bac=True,
    extra_validations=None,
    field_checks=None,
    model_field_id=None,
    models=None,
    log_cb=None,
    progress_cb=None,
    stop_flag=None,
    screenshot_cb=None,
):
    """
    rows: filas del Excel ya filtradas (dicts, ver read_excel_rows/filter_rows).
    column_map: {"region": <col_key>, "city": <col_key>, "dealer": <col_key>, "bac": <col_key>}
    level_ids: ids de los <select> del form, default DEFAULT_SELECT_IDS (region/city/dealer).
    extra_validations: lista de {"attr": "lat", "label": "Latitud", "column": <col_key>}
                       — compara data-<attr> del <option> de dealer contra la columna del Excel.
    field_checks: lista de {"column": <col_key>, "field_id": "id-del-campo"} — compara el valor
                  actual de cualquier campo del form (input o select, por id) contra una columna
                  del Excel. Se corre por fila, independientemente de si el dealer fue encontrado.
    model_field_id + models: si se pasan ambos (models no vacía), antes de cada tanda de filas se
                  selecciona ese modelo en el <select> de modelos (por id) y se repite la comparación
                  completa para cada modelo de la lista — útil cuando el listado de dealers cambia
                  según el modelo elegido (ej. modelos eléctricos con dealers limitados).
    Los tiempos de espera de los selects dependientes son fijos e internos (FAST_TIMEOUT/MAX_TIMEOUT,
    con reintento automático) — no son configurables desde afuera.

    Devuelve una lista de dicts de resultado: status, region, city, dealer, modelo, bac_excel, bac_form,
    fails (lista de strings), fila (número de fila en el Excel original).
    """
    level_ids = level_ids or DEFAULT_SELECT_IDS
    extra_validations = extra_validations or []
    field_checks = field_checks or []
    log = log_cb or (lambda *_a, **_k: None)

    region_key = column_map.get("region")
    city_key = column_map.get("city")
    dealer_key = column_map.get("dealer")
    bac_key = column_map.get("bac")

    results = []
    models_to_run = list(models) if (model_field_id and models) else [None]
    total = len(rows) * len(models_to_run)
    counter = 0

    for model_text in models_to_run:
        if model_text:
            _check_stop(stop_flag)
            log(f"\n=== MODELO: {model_text} ===", "info")
            if not _select_model(driver, model_field_id, model_text):
                log(f"  Modelo '{model_text}' no encontrado en el form, se omite", "warn")
                counter += len(rows)
                continue
            time.sleep(0.3)

        for row in rows:
            counter += 1
            _check_stop(stop_flag)
            if progress_cb:
                label = row.get(dealer_key) or row.get(city_key) or f"Fila {row.get('__row__')}"
                progress_cb(counter, total, f"{model_text} · {label}" if model_text else label)

            region_text = row.get(region_key, "") if has_region else ""
            city_text = row.get(city_key, "") if has_city else ""
            dealer_text = row.get(dealer_key, "")
            bac_excel = row.get(bac_key, "") if chk_bac and bac_key else ""

            fails = []
            region_found = not has_region
            city_found = not has_city
            dealer_found = False
            bac_ok = None
            dealer_option = None
            extra_results = {}

            try:
                if has_region and region_text:
                    region_el = _get_select(driver, level_ids["region"])
                    if region_el is None:
                        fails.append(f"Select de región '{level_ids['region']}' no encontrado en el form")
                    else:
                        opt = _find_option_by_text(region_el, region_text)
                        region_found = opt is not None
                        if region_found:
                            _select_option(driver, region_el, opt)
                            log(f"  Región: {region_text}", "ok")
                        else:
                            fails.append(f"Región '{region_text}' no encontrada")
                            log(f"  Región no encontrada: {region_text}", "warn")

                if has_city and city_text and (region_found or not has_region):
                    city_el = _wait_options_loaded(driver, level_ids["city"])
                    if city_el is None:
                        fails.append(f"Select de ciudad '{level_ids['city']}' no disponible")
                    else:
                        opt = _find_option_by_text(city_el, city_text)
                        city_found = opt is not None
                        if city_found:
                            _select_option(driver, city_el, opt)
                            log(f"  Ciudad: {city_text}", "ok")
                        else:
                            fails.append(f"Ciudad '{city_text}' no encontrada")
                            log(f"  Ciudad no encontrada: {city_text}", "warn")

                ready_for_dealer = (region_found or not has_region) and (city_found or not has_city)
                if dealer_text and ready_for_dealer:
                    dealer_el = _wait_options_loaded(driver, level_ids["dealer"])
                    if dealer_el is None:
                        fails.append(f"Select de dealer '{level_ids['dealer']}' no disponible")
                    else:
                        dealer_option = _find_option_by_text(dealer_el, dealer_text)
                        dealer_found = dealer_option is not None
                        if dealer_found:
                            _select_option(driver, dealer_el, dealer_option)
                            log(f"  Dealer: {dealer_text}", "ok")
                            time.sleep(0.3)
                        else:
                            fails.append(f"Dealer '{dealer_text}' no encontrado")
                            log(f"  Dealer no encontrado: {dealer_text}", "warn")

                if dealer_found and dealer_option is not None:
                    if chk_bac and bac_key and bac_excel:
                        bac_form = (dealer_option.get_attribute("data-bac") or "").strip()
                        bac_ok = normalize_text(bac_excel) == normalize_text(bac_form)
                        if not bac_ok:
                            fails.append(f"BAC no coincide (excel='{bac_excel}' form='{bac_form}')")
                    for extra in extra_validations:
                        excel_val = row.get(extra["column"], "")
                        if not excel_val:
                            continue
                        form_val = (dealer_option.get_attribute(f"data-{extra['attr']}") or "").strip()
                        ok = normalize_text(excel_val) == normalize_text(form_val)
                        extra_results[extra["label"]] = {"ok": ok, "excel": excel_val, "form": form_val}
                        if not ok:
                            fails.append(f"{extra['label']} no coincide (excel='{excel_val}' form='{form_val}')")

                for check in field_checks:
                    excel_val = row.get(check["column"], "")
                    if not excel_val:
                        continue
                    form_val = _read_form_field_value(driver, check["field_id"]) or ""
                    ok = normalize_text(excel_val) == normalize_text(form_val)
                    extra_results[check["field_id"]] = {"ok": ok, "excel": excel_val, "form": form_val}
                    if not ok:
                        fails.append(
                            f"Campo '{check['field_id']}' no coincide (excel='{excel_val}' form='{form_val}')"
                        )

                status = "PASS" if not fails else "FAIL"

            except StopRequested:
                raise
            except Exception as e:  # noqa: BLE001
                status = "FAIL"
                fails.append(f"Error inesperado: {e}")
                log(f"  Error: {e}", "err")

            result = {
                "status": status,
                "modelo": model_text or "",
                "region": region_text,
                "city": city_text,
                "dealer": dealer_text,
                "bac_excel": bac_excel,
                "bac_ok": bac_ok,
                "extra_results": extra_results,
                "fails": fails,
                "fila": row.get("__row__"),
            }

            if screenshot_cb is not None:
                try:
                    screenshot_cb(result)
                except StopRequested:
                    raise
                except Exception as e:  # noqa: BLE001
                    log(f"  Error tomando captura: {e}", "err")

            results.append(result)

    return results


def find_extra_dealers(
    driver,
    rows,
    column_map,
    level_ids=None,
    has_region=True,
    has_city=True,
    log_cb=None,
    progress_cb=None,
    stop_flag=None,
):
    """Recorre las combinaciones región/ciudad presentes en el Excel (filtrado o no, según
    corresponda) y reporta, mirando el <select> real del form:
    - EXTRA: dealers que aparecen en el form pero no están en la lista esperada del Excel.
    - DUPLICADO: dealers que aparecen más de una vez como <option> en el mismo <select> del
      form (no importa si el Excel tiene duplicados — lo que importa es el form)."""
    level_ids = level_ids or DEFAULT_SELECT_IDS
    log = log_cb or (lambda *_a, **_k: None)

    region_key = column_map.get("region")
    city_key = column_map.get("city")
    dealer_key = column_map.get("dealer")

    expected_by_combo = {}
    for row in rows:
        combo = (row.get(region_key, "") if has_region else "", row.get(city_key, "") if has_city else "")
        expected_by_combo.setdefault(combo, set()).add(normalize_text(row.get(dealer_key, "")))

    combos = list(expected_by_combo.keys()) or [("", "")]
    extra_results = []
    total = len(combos)

    for idx, (region_text, city_text) in enumerate(combos, start=1):
        _check_stop(stop_flag)
        if progress_cb:
            progress_cb(idx, total, f"{region_text} / {city_text}".strip(" /"))

        try:
            if has_region and region_text:
                region_el = _get_select(driver, level_ids["region"])
                opt = _find_option_by_text(region_el, region_text) if region_el is not None else None
                if opt is None:
                    continue
                _select_option(driver, region_el, opt)

            if has_city and city_text:
                city_el = _wait_options_loaded(driver, level_ids["city"])
                opt = _find_option_by_text(city_el, city_text) if city_el is not None else None
                if opt is None:
                    continue
                _select_option(driver, city_el, opt)

            dealer_el = _wait_options_loaded(driver, level_ids["dealer"])
            if dealer_el is None:
                continue
            expected = expected_by_combo[(region_text, city_text)]
            seen_in_form = {}
            for opt in _valid_options(dealer_el):
                if _is_placeholder(opt.text):
                    continue
                norm = normalize_text(opt.text)
                seen_in_form.setdefault(norm, []).append(opt.text.strip())
                if norm not in expected:
                    log(f"  EXTRA: {opt.text} ({region_text}/{city_text})", "warn")
                    extra_results.append({
                        "status": "EXTRA",
                        "region": region_text,
                        "city": city_text,
                        "dealer": opt.text.strip(),
                    })
            for norm, texts in seen_in_form.items():
                if len(texts) > 1:
                    log(f"  DUPLICADO en el form: {texts[0]} x{len(texts)} ({region_text}/{city_text})", "warn")
                    extra_results.append({
                        "status": "DUPLICADO",
                        "region": region_text,
                        "city": city_text,
                        "dealer": texts[0],
                        "fails": [f"Aparece {len(texts)} veces en el <select> del form"],
                    })
        except StopRequested:
            raise
        except Exception as e:  # noqa: BLE001
            log(f"  Error buscando extras en {region_text}/{city_text}: {e}", "err")

    return extra_results


# ──────────────────────────────────────────────────────────────────────────────
# Captura de pantalla con banner de URL + ZIP
# ──────────────────────────────────────────────────────────────────────────────
def capture_result_screenshot(driver, screenshot_dir, filename, form_url):
    """Toma una captura de página completa con la URL del form pegada en un banner
    superior (reusa ScreenshotManager, que ya soporta esto)."""
    os.makedirs(screenshot_dir, exist_ok=True)
    manager = ScreenshotManager(driver, screenshot_dir)
    manager.url_form_esperado = form_url
    manager.take_full_page_screenshot(filename)
    manager._add_url_banner(os.path.join(screenshot_dir, filename))
    return os.path.join(screenshot_dir, filename)


def zip_screenshots(screenshot_paths, output_zip_path):
    os.makedirs(os.path.dirname(output_zip_path), exist_ok=True)
    with zipfile.ZipFile(output_zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in screenshot_paths:
            if os.path.exists(path):
                zf.write(path, arcname=os.path.basename(path))
    return output_zip_path


# ──────────────────────────────────────────────────────────────────────────────
# Export a Excel (PASS verde / FAIL rojo / EXTRA amarillo / MISSING violeta)
# ──────────────────────────────────────────────────────────────────────────────
def export_results_excel(results, output_path=None, pais=""):
    if output_path is None:
        results_dir = os.path.join(RESULTS_DIR, "dealer_comparator")
        os.makedirs(results_dir, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(results_dir, f"dealer_comparator_{pais}_{stamp}.xlsx")

    wb = Workbook()
    ws = wb.active
    ws.title = "resultados"

    headers = ["Estado", "URL Form", "Modelo", "Región", "Ciudad", "Dealer", "BAC Excel", "BAC OK",
               "Detalle", "Fila Excel"]
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT

    fill_by_status = {
        "PASS": _PASS_FILL, "FAIL": _FAIL_FILL, "EXTRA": _EXTRA_FILL,
        "MISSING": _MISSING_FILL, "DUPLICADO": _DUPLICATE_FILL,
    }

    for r in results:
        ws.append([
            r.get("status", ""),
            r.get("url_form", ""),
            r.get("modelo", ""),
            r.get("region", ""),
            r.get("city", ""),
            r.get("dealer", ""),
            r.get("bac_excel", ""),
            {True: "OK", False: "MISMATCH", None: ""}.get(r.get("bac_ok")),
            " | ".join(r.get("fails", [])) if r.get("fails") else "",
            r.get("fila", ""),
        ])
        row_idx = ws.max_row
        fill = fill_by_status.get(r.get("status"))
        if fill:
            for cell in ws[row_idx]:
                cell.fill = fill

    for i, width in enumerate((10, 40, 16, 18, 18, 28, 14, 10, 60, 10), start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = width

    total = len(results)
    counts = {"PASS": 0, "FAIL": 0, "EXTRA": 0, "MISSING": 0, "DUPLICADO": 0}
    for r in results:
        counts[r.get("status", "FAIL")] = counts.get(r.get("status", "FAIL"), 0) + 1

    summary_ws = wb.create_sheet("resumen")
    summary_ws.append(["País", pais])
    summary_ws.append(["Total", total])
    for status, count in counts.items():
        summary_ws.append([status, count])

    wb.save(output_path)
    return output_path
