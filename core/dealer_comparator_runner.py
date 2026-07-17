"""
dealer_comparator_runner.py — Lógica de negocio de la pestaña "Comparador Dealers".

Lee un Excel de dealers esperados (fila de encabezado y columnas configurables porque
el Excel que mandan varía de país a país), navega región→ciudad→dealer en un formulario
real via Selenium (reusando BrowserManager), compara contra el Excel filtrado y arma un
reporte PASS/FAIL/EXTRA/DUPLICADO/NOTA con export a Excel (openpyxl) y capturas opcionales
(ScreenshotManager, con banner de URL) — una carpeta propia por formulario, sin ZIP.
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
from selenium.common.exceptions import StaleElementReferenceException

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
_NOTA_FILL = PatternFill(fill_type="solid", fgColor="00BDD7EE")
_OCULTO_FILL = PatternFill(fill_type="solid", fgColor="00FFB84D")  # naranja: relación con filas OCULTAS del Excel
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
# Matcheo tolerante de nombres de dealer (apóstrofes/comillas, sufijos "(1000km)")
# ──────────────────────────────────────────────────────────────────────────────
_TRAILING_PAREN_RE = re.compile(r"\(\s*[^()]*\d[^()]*\)\s*$")  # "(1000KM)" al final
_PUNCT_STRIP_RE = re.compile(r"[\'’‘´`\.\,\-]+")


def _canonical_dealer_name(text):
    """Nombre normalizado y sin apóstrofes/comillas ni sufijo parentético numérico final
    (ej. '(1000km)'), para el matcheo tolerante de nombres de dealer. Dos nombres que
    coinciden solo tras esta limpieza se consideran 'el mismo dealer con diferencias
    menores' — cualquier otra diferencia (contenido de más o de menos) NO matchea."""
    t = normalize_text(text)
    t = _TRAILING_PAREN_RE.sub("", t).strip()
    t = _PUNCT_STRIP_RE.sub("", t)
    return re.sub(r"\s+", " ", t).strip()


def match_dealer_names(excel_name, form_name):
    """Compara el nombre del Excel contra el texto real del form.
    Devuelve (match: bool, disclaimer: str|None):
      - Coincidencia exacta (normalize_text) → (True, None)
      - Coincidencia solo tras sacar apóstrofes/comillas o un sufijo parentético numérico
        final → (True, "disclaimer explicando la diferencia menor")
      - Cualquier otra diferencia (más o menos contenido) → (False, None)
    """
    excel_norm = normalize_text(excel_name)
    form_norm = normalize_text(form_name)
    if not excel_norm or not form_norm:
        return False, None
    if excel_norm == form_norm:
        return True, None
    excel_canon = _canonical_dealer_name(excel_name)
    form_canon = _canonical_dealer_name(form_name)
    if excel_canon and excel_canon == form_canon:
        disclaimer = (
            f"Nombre difiere en detalles menores (mayúsculas/caracteres especiales/sufijo) "
            f"— Excel: '{str(excel_name).strip()}' | Form: '{str(form_name).strip()}'"
        )
        return True, disclaimer
    return False, None


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


def detect_hidden_rows(file_path, header_row=1, sheet_name=None):
    """Devuelve el set de números de fila (1-indexado) de datos que están OCULTAS en el
    Excel (el archivo a veces viene pre-filtrado con filas escondidas). Se usa sólo para
    avisar al usuario — no cambia qué se compara. read_only no expone row_dimensions,
    por eso se abre en modo normal."""
    hidden = set()
    try:
        wb = load_workbook(file_path, data_only=True, read_only=False)
        try:
            sheet = wb[sheet_name] if sheet_name else wb.worksheets[0]
            for row_num, dim in sheet.row_dimensions.items():
                if row_num > header_row and getattr(dim, "hidden", False):
                    hidden.add(row_num)
        finally:
            wb.close()
    except Exception:
        pass
    return hidden


def detect_hidden_columns(file_path, header_row=1, sheet_name=None):
    """Devuelve [{'column': 'K', 'header': 'VENTA 0KM'}, ...] de columnas OCULTAS en el
    Excel — igual criterio que detect_hidden_rows pero para columnas. Solo informativo:
    no cambia qué se compara, pero avisa porque una columna oculta puede pasar
    desapercibida para quien preparó el Excel."""
    hidden = []
    try:
        wb = load_workbook(file_path, data_only=True, read_only=False)
        try:
            sheet = wb[sheet_name] if sheet_name else wb.worksheets[0]
            header_cells = list(sheet.iter_rows(min_row=header_row, max_row=header_row, values_only=True))
            header_vals = header_cells[0] if header_cells else []
            for col_letter, dim in sheet.column_dimensions.items():
                if not getattr(dim, "hidden", False):
                    continue
                try:
                    idx = column_index_from_string(col_letter) - 1
                    header_name = str(header_vals[idx]).strip() if (0 <= idx < len(header_vals) and header_vals[idx]) else ""
                except Exception:
                    header_name = ""
                hidden.append({"column": col_letter, "header": header_name})
        finally:
            wb.close()
    except Exception:
        pass
    return sorted(hidden, key=lambda h: h["column"])


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
    """Busca la <option> cuyo texto matchea: exacto primero, y si no, solo tolerando
    diferencias menores (apóstrofes/comillas, sufijo numérico entre paréntesis) vía
    match_dealer_names — NO hace substring genérico: un nombre con contenido de más o de
    menos (ej. otro apellido agregado) NO debe considerarse el mismo dealer.
    Devuelve (option, disclaimer_o_None) o (None, None) si no hay match."""
    target = normalize_text(text)
    if not target:
        return None, None
    options = _valid_options(select_element)
    for opt in options:
        if normalize_text(opt.text) == target:
            return opt, None
    for opt in options:
        match, disclaimer = match_dealer_names(text, opt.text)
        if match:
            return opt, disclaimer
    return None, None


def _select_option(driver, select_element, option_element):
    Select(select_element).select_by_value(option_element.get_attribute("value"))
    driver.execute_script(
        "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));"
        "arguments[0].dispatchEvent(new Event('input', {bubbles:true}));",
        select_element,
    )


def _select_by_text_robust(driver, select_id, target_text, wait_for_option=False, timeout=None):
    """
    Selecciona la <option> cuyo texto matchea target_text en el <select> id=select_id,
    de forma resistente a stale elements y a selects dependientes que se repueblan
    tarde (city tras region, dealer tras city). En cada intento re-busca el <select>
    y sus <option> FRESCOS — así un re-render del form (que invalida referencias
    viejas) no rompe la selección ni genera un "no encontrado" falso.

    wait_for_option=True: reintenta hasta `timeout` esperando a que la opción aparezca
    (el select dependiente todavía puede estar cargando las opciones de la nueva región).
    Devuelve (found: bool, option_text: str|None, disclaimer: str|None) — disclaimer no-None
    si matcheó tolerando una diferencia menor de nombre (ver match_dealer_names).
    """
    if timeout is None:
        # Selects dependientes (city/dealer) pueden poblar la opción buscada con retraso
        # (ej. la 3ra opción de dealer en algunos forms carga tarde). Presupuesto amplio:
        # el poll corta APENAS aparece, así que sólo se agota en dealers realmente ausentes.
        timeout = DEPENDENT_SELECT_TIMEOUT if wait_for_option else 0.4
    deadline = time.time() + timeout
    while True:
        try:
            el = _get_select(driver, select_id, timeout=0.5)
            if el is not None:
                opt, disclaimer = _find_option_by_text(el, target_text)
                if opt is not None:
                    text = opt.text.strip()
                    try:
                        _select_option(driver, el, opt)
                    except Exception as sel_err:
                        # La opción EXISTE (dealer presente en el form) pero está
                        # deshabilitada / no seleccionable: a los fines de la comparación
                        # (¿está el dealer en el form?) cuenta como encontrado.
                        if "disabled" in str(sel_err).lower():
                            return True, text, disclaimer
                        raise
                    return True, text, disclaimer
        except StaleElementReferenceException:
            pass
        if time.time() >= deadline:
            return False, None, None
        time.sleep(0.1)


def _read_valid_option_texts(driver, select_id, wait_nonempty=True, timeout=None):
    """Devuelve la lista de textos de <option> válidas (no placeholder) del select,
    leídas de forma stale-safe (re-busca el elemento y re-lee si algo se pone stale
    a mitad de la iteración — el form re-renderiza los selects).

    wait_nonempty=True: espera hasta `timeout` a que el select tenga opciones (el select
    dependiente se puebla async tras cambiar el padre); corta apenas aparecen."""
    if timeout is None:
        timeout = DEPENDENT_SELECT_TIMEOUT if wait_nonempty else 0.4
    deadline = time.time() + timeout
    while True:
        try:
            el = _get_select(driver, select_id, timeout=0.5)
            if el is not None:
                texts = []
                for o in el.find_elements(By.TAG_NAME, "option"):
                    if not o.get_attribute("value"):
                        continue
                    t = (o.text or "").strip()
                    if t and not _is_placeholder(t):
                        texts.append(t)
                if texts or not wait_nonempty:
                    return texts
        except StaleElementReferenceException:
            pass
        if time.time() >= deadline:
            return []
        time.sleep(0.1)


def _selected_dealer_option_attr(driver, dealer_id, attr):
    """Lee un atributo (ej. data-bac) de la <option> de dealer actualmente seleccionada,
    re-buscándola fresca (stale-safe)."""
    for _ in range(3):
        try:
            el = _get_select(driver, dealer_id, timeout=0.5)
            if el is None:
                return ""
            return (Select(el).first_selected_option.get_attribute(attr) or "").strip()
        except StaleElementReferenceException:
            time.sleep(0.1)
    return ""


FAST_TIMEOUT = 0.6  # primer intento: rápido, alcanza en la gran mayoría de los forms
MAX_TIMEOUT = 1.2    # segundo intento (solo si el primero no alcanzó): tope máximo
DEPENDENT_SELECT_TIMEOUT = 5.0  # espera máxima a que aparezca la opción buscada en un
                                 # select dependiente (city/dealer); corta apenas la halla
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
    opt, _disclaimer = _find_option_by_text(el, model_text)
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
    expect_absent=False,
    has_dealer=True,
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
    _log_raw = log_cb or (lambda *_a, **_k: None)
    def log(msg, *a, **k):
        # Nunca debe poder tirar una excepción: un log_cb que falla (ej. encoding de
        # consola en un caracter especial) no puede convertir un resultado válido en FAIL.
        try:
            _log_raw(msg, *a, **k)
        except Exception:
            pass

    region_key = column_map.get("region")
    city_key = column_map.get("city")
    dealer_key = column_map.get("dealer") if has_dealer else None
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

        # Estado de navegación: qué región/ciudad están seleccionadas ahora mismo en el
        # form. Re-seleccionar el MISMO valor rompe algunos forms (limpian el select hijo
        # y no lo repueblan porque "no cambió") → dealers presentes daban FAIL falso. Por
        # eso sólo se re-navega región/ciudad cuando cambian respecto a la fila anterior;
        # si son iguales, el select de dealer ya tiene la lista cargada y se elige directo.
        cur_region = None
        cur_city = None

        for row in rows:
            counter += 1
            _check_stop(stop_flag)
            if progress_cb:
                label = (row.get(dealer_key) if dealer_key else None) or row.get(city_key) or row.get(region_key) or f"Fila {row.get('__row__')}"
                progress_cb(counter, total, f"{model_text} · {label}" if model_text else label)

            region_text = row.get(region_key, "") if has_region else ""
            city_text = row.get(city_key, "") if has_city else ""
            dealer_text = row.get(dealer_key, "") if (has_dealer and dealer_key) else ""
            bac_excel = row.get(bac_key, "") if chk_bac and bac_key else ""

            fails = []
            region_found = not has_region
            city_found = not has_city
            dealer_found = False
            bac_ok = None
            extra_results = {}
            name_disclaimer = None
            detalle_info = ""

            try:
                # Orden SIEMPRE: región → ciudad → dealer (si no hay región, arranca en ciudad).
                # Cada nivel usa _select_by_text_robust: re-busca el <select> fresco en cada
                # intento (stale-safe ante re-render del form) y espera a que la opción aparezca
                # en los selects dependientes (city tras region, dealer tras city). Sólo se
                # (re)selecciona un nivel si su valor cambió respecto a la fila anterior.
                if has_region and region_text:
                    if region_text == cur_region:
                        region_found = True  # ya seleccionada de una fila previa
                    elif _get_select(driver, level_ids["region"], timeout=2) is None:
                        fails.append(f"Select de región '{level_ids['region']}' no encontrado en el form")
                        cur_region = None
                        cur_city = None
                    else:
                        region_found, _, _ = _select_by_text_robust(driver, level_ids["region"], region_text)
                        cur_city = None  # cambió la región → el select de ciudad se repuebla
                        if region_found:
                            cur_region = region_text
                            log(f"  Región: {region_text}", "ok")
                        else:
                            cur_region = None
                            fails.append(f"Región '{region_text}' no encontrada")
                            log(f"  Región no encontrada: {region_text}", "warn")

                if has_city and city_text and (region_found or not has_region):
                    if city_text == cur_city:
                        city_found = True  # ya seleccionada (misma región y ciudad que la fila previa)
                    else:
                        city_found, _, _ = _select_by_text_robust(
                            driver, level_ids["city"], city_text, wait_for_option=True)
                        if city_found:
                            cur_city = city_text
                            log(f"  Ciudad: {city_text}", "ok")
                        else:
                            cur_city = None
                            fails.append(f"Ciudad '{city_text}' no encontrada")
                            log(f"  Ciudad no encontrada: {city_text}", "warn")

                ready_for_dealer = (region_found or not has_region) and (city_found or not has_city)
                if has_dealer and dealer_text and ready_for_dealer:
                    dealer_found, _, name_disclaimer = _select_by_text_robust(
                        driver, level_ids["dealer"], dealer_text, wait_for_option=True)
                    if dealer_found:
                        log(f"  Dealer: {dealer_text}"
                            + (f"  ⚠ {name_disclaimer}" if name_disclaimer else ""), "ok")
                        time.sleep(0.2)
                    elif not expect_absent:
                        fails.append(f"Dealer '{dealer_text}' no encontrado")
                        log(f"  Dealer no encontrado: {dealer_text}", "warn")

                if expect_absent:
                    # Modo EXCLUIR: este dealer NO debería estar en el form. Sólo importa
                    # su presencia — se descartan los fails de región/ciudad (si no están,
                    # el dealer tampoco puede estar → correctamente ausente). detalle_info
                    # explica en QUÉ nivel se confirmó la ausencia (misma lógica que usa
                    # la captura de evidencia para elegir qué dropdown desplegar).
                    fails = []
                    if dealer_found:
                        fails.append(f"Dealer '{dealer_text}' ESTÁ en el form pero NO debería (excluido)")
                        log(f"  ✗ '{dealer_text}' presente cuando debería estar ausente", "warn")
                    else:
                        if has_region and region_text and not region_found:
                            detalle_info = (f"Correctamente ausente: la región '{region_text}' no está en el "
                                            f"form, por lo tanto el dealer tampoco puede aparecer")
                        elif has_city and city_text and not city_found:
                            detalle_info = (f"Correctamente ausente: la ciudad '{city_text}' no está en el "
                                            f"form, por lo tanto el dealer tampoco puede aparecer")
                        else:
                            detalle_info = f"Correctamente ausente: '{dealer_text}' no aparece en el dropdown de dealers"
                        log(f"  ✓ '{dealer_text}' correctamente ausente del form", "ok")
                elif dealer_found:
                    if chk_bac and bac_key and bac_excel:
                        bac_form = _selected_dealer_option_attr(driver, level_ids["dealer"], "data-bac")
                        bac_ok = normalize_text(bac_excel) == normalize_text(bac_form)
                        if not bac_ok:
                            fails.append(f"BAC no coincide (excel='{bac_excel}' form='{bac_form}')")
                    for extra in extra_validations:
                        excel_val = row.get(extra["column"], "")
                        if not excel_val:
                            continue
                        form_val = _selected_dealer_option_attr(driver, level_ids["dealer"], f"data-{extra['attr']}")
                        ok = normalize_text(excel_val) == normalize_text(form_val)
                        extra_results[extra["label"]] = {"ok": ok, "excel": excel_val, "form": form_val}
                        if not ok:
                            fails.append(f"{extra['label']} no coincide (excel='{excel_val}' form='{form_val}')")

                if not expect_absent:
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
                "name_disclaimer": name_disclaimer if (dealer_found and not expect_absent) else None,
                "detalle_info": detalle_info,
                # En qué nivel cortó (o no) la navegación — lo usa la captura de evidencia
                # de exclusión para desplegar el dropdown correcto (región/ciudad/dealer).
                "region_found": bool(region_found),
                "city_found": bool(city_found),
                "dealer_found": bool(dealer_found),
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
    has_dealer=True,
    hidden_rows_data=None,
):
    """
    Recorre las combinaciones del nivel superior presentes en el Excel y reporta:
    - EXTRA: opciones que aparecen en el form pero no están en la lista esperada del Excel.
    - DUPLICADO: opciones que aparecen más de una vez como <option> en el form.
    - OCULTO: opciones del form que SÓLO están declaradas en filas OCULTAS del Excel
      (hidden_rows_data) — están mal porque la fila está oculta; van en naranja.
    Este chequeo se realiza de forma jerárquica (Región -> Ciudad -> Dealer) según los niveles activos,
    buscando en conjunto y reportando la combinación exacta de forma contextual.
    """
    level_ids = level_ids or DEFAULT_SELECT_IDS
    _log_raw = log_cb or (lambda *_a, **_k: None)
    def log(msg, *a, **k):
        # Nunca debe poder tirar una excepción: un log_cb que falla (ej. encoding de
        # consola en un caracter especial) no puede convertir un resultado válido en FAIL.
        try:
            _log_raw(msg, *a, **k)
        except Exception:
            pass

    region_key = column_map.get("region")
    city_key = column_map.get("city")
    dealer_key = column_map.get("dealer")

    def _build_expected_maps(source_rows):
        """Indexa (regiones, ciudades_por_región, dealers_por_región_ciudad, canon_dealers,
        orig_region, orig_city) a partir de un set de filas del Excel."""
        regions = set()
        cities_by_region = {}
        dealers_by_rc = {}
        dealer_canon_by_rc = {}
        orig_reg = {}
        orig_city = {}
        for row in source_rows or []:
            reg_val = str(row.get(region_key, "")).strip() if (has_region and region_key) else ""
            city_val = str(row.get(city_key, "")).strip() if (has_city and city_key) else ""
            dealer_val = str(row.get(dealer_key, "")).strip() if (has_dealer and dealer_key) else ""

            norm_reg = normalize_text(reg_val)
            norm_city = normalize_text(city_val)

            if has_region and reg_val:
                regions.add(norm_reg)
                orig_reg[norm_reg] = reg_val
            if has_city and city_val:
                cities_by_region.setdefault(norm_reg, set()).add(norm_city)
                orig_city[norm_city] = city_val
            if has_dealer and dealer_val:
                dealers_by_rc.setdefault((norm_reg, norm_city), set()).add(normalize_text(dealer_val))
                dealer_canon_by_rc.setdefault((norm_reg, norm_city), {})[
                    _canonical_dealer_name(dealer_val)
                ] = dealer_val
        return regions, cities_by_region, dealers_by_rc, dealer_canon_by_rc, orig_reg, orig_city

    # Filas visibles (lo esperado real) vs filas OCULTAS del Excel (lo que NO debería contar
    # pero, si aparece en el form, hay que señalarlo en naranja como OCULTO)
    (expected_regions, expected_cities_by_region, expected_dealers_by_region_city,
     expected_dealer_canon_by_region_city, orig_region_by_norm, orig_city_by_norm) = _build_expected_maps(rows)
    (hidden_regions, hidden_cities_by_region, hidden_dealers_by_region_city,
     _hidden_canon, hidden_orig_reg, hidden_orig_city) = _build_expected_maps(hidden_rows_data)
    # Para poder navegar también las combinaciones que sólo existen en filas ocultas
    orig_region_by_norm = {**hidden_orig_reg, **orig_region_by_norm}
    orig_city_by_norm = {**hidden_orig_city, **orig_city_by_norm}

    def _classify_dealer_options(option_texts, key):
        """Separa las opciones vistas en el form en (duplicadas, extras, notas, ocultos)
        contra lo esperado para (norm_region, norm_city). Tolera diferencias menores de
        nombre (ver match_dealer_names) → 'notas'. Lo declarado sólo en filas OCULTAS del
        Excel → 'ocultos' (mal: la fila está oculta), no 'extras'."""
        expected_norm = expected_dealers_by_region_city.get(key, set())
        expected_canon = expected_dealer_canon_by_region_city.get(key, {})
        hidden_norm = hidden_dealers_by_region_city.get(key, set())
        seen_in_form = {}
        extras, notas, ocultos = [], [], []
        for text in option_texts:
            norm = normalize_text(text)
            seen_in_form.setdefault(norm, []).append(text)
            if norm in expected_norm:
                continue
            if norm in hidden_norm:
                ocultos.append(text)
                continue
            canon = _canonical_dealer_name(text)
            if canon in expected_canon:
                notas.append((text, expected_canon[canon]))
            else:
                extras.append(text)
        duplicated = [(texts[0], len(texts)) for texts in seen_in_form.values() if len(texts) > 1]
        return duplicated, extras, notas, ocultos

    extra_results = []

    # --------------------------------------------------------------------------
    # FASE A: Chequeo de Región (sólo si has_region es True)
    # --------------------------------------------------------------------------
    regions_to_check = [""]
    if has_region:
        log("Chequeando EXTRAS y DUPLICADOS en Regiones...", "info")
        option_texts = _read_valid_option_texts(driver, level_ids["region"])
        seen_in_form = {}
        for text in option_texts:
            norm = normalize_text(text)
            seen_in_form.setdefault(norm, []).append(text)
            if norm in expected_regions:
                pass
            elif norm in hidden_regions:
                log(f"  OCULTO REGION: {text} (declarada solo en fila oculta del Excel)", "warn")
                extra_results.append({
                    "status": "OCULTO",
                    "region": text,
                    "city": "",
                    "dealer": "",
                    "fails": ["Región presente en el form pero declarada SOLO en fila(s) OCULTA(s) del Excel"],
                })
            else:
                log(f"  EXTRA REGION: {text}", "warn")
                extra_results.append({
                    "status": "EXTRA",
                    "region": text,
                    "city": "",
                    "dealer": "",
                    "fails": ["Región extra (no declarada en el Excel)"],
                })
        for norm, texts in seen_in_form.items():
            if len(texts) > 1:
                log(f"  DUPLICADA REGION: {texts[0]} x{len(texts)}", "warn")
                extra_results.append({
                    "status": "DUPLICADO",
                    "region": texts[0],
                    "city": "",
                    "dealer": "",
                    "fails": [f"Región aparece {len(texts)} veces en el dropdown"],
                })

        # Regiones a recorrer: las declaradas (visibles u ocultas) que existen en el form —
        # las ocultas también se navegan para poder clasificar sus ciudades/dealers.
        regions_to_check = [norm for norm in (expected_regions | hidden_regions) if norm in seen_in_form]

    # --------------------------------------------------------------------------
    # FASE B: Chequeo de Ciudad (sólo si has_city es True)
    # --------------------------------------------------------------------------
    cities_to_check = []  # lista de (norm_reg, norm_city)
    
    if has_city:
        if has_region:
            log("Chequeando EXTRAS y DUPLICADOS en Ciudades por Región...", "info")
            for norm_reg in regions_to_check:
                _check_stop(stop_flag)
                reg_orig_text = orig_region_by_norm[norm_reg]
                if progress_cb:
                    progress_cb(1, 1, f"Región: {reg_orig_text}")

                # Seleccionar la región para ver sus ciudades
                ok, _, _ = _select_by_text_robust(driver, level_ids["region"], reg_orig_text)
                if not ok:
                    continue
                time.sleep(0.6)  # Pausa estable para repoblar ciudades

                option_texts = _read_valid_option_texts(driver, level_ids["city"])
                expected_cities = expected_cities_by_region.get(norm_reg, set())
                hidden_cities = hidden_cities_by_region.get(norm_reg, set())
                seen_in_form = {}
                for text in option_texts:
                    norm = normalize_text(text)
                    seen_in_form.setdefault(norm, []).append(text)
                    if norm in expected_cities:
                        pass
                    elif norm in hidden_cities:
                        log(f"  OCULTO CIUDAD: {text} en {reg_orig_text} (solo en fila oculta del Excel)", "warn")
                        extra_results.append({
                            "status": "OCULTO",
                            "region": reg_orig_text,
                            "city": text,
                            "dealer": "",
                            "fails": [f"Ciudad presente en el form para '{reg_orig_text}' pero declarada SOLO en fila(s) OCULTA(s) del Excel"],
                        })
                    else:
                        log(f"  EXTRA CIUDAD: {text} (en Región: {reg_orig_text})", "warn")
                        extra_results.append({
                            "status": "EXTRA",
                            "region": reg_orig_text,
                            "city": text,
                            "dealer": "",
                            "fails": [f"Ciudad extra para la región '{reg_orig_text}'"],
                        })
                for norm, texts in seen_in_form.items():
                    if len(texts) > 1:
                        log(f"  DUPLICADA CIUDAD: {texts[0]} x{len(texts)} (en Región: {reg_orig_text})", "warn")
                        extra_results.append({
                            "status": "DUPLICADO",
                            "region": reg_orig_text,
                            "city": texts[0],
                            "dealer": "",
                            "fails": [f"Ciudad aparece {len(texts)} veces en el dropdown de la región '{reg_orig_text}'"],
                        })

                for norm_city in (expected_cities | hidden_cities):
                    if norm_city in seen_in_form:
                        cities_to_check.append((norm_reg, norm_city))
        else:
            # Sin Región, las ciudades son globales
            log("Chequeando EXTRAS y DUPLICADOS en Ciudades...", "info")
            option_texts = _read_valid_option_texts(driver, level_ids["city"])
            expected_cities = expected_cities_by_region.get("", set())
            hidden_cities = hidden_cities_by_region.get("", set())
            seen_in_form = {}
            for text in option_texts:
                norm = normalize_text(text)
                seen_in_form.setdefault(norm, []).append(text)
                if norm in expected_cities:
                    pass
                elif norm in hidden_cities:
                    log(f"  OCULTO CIUDAD: {text} (solo en fila oculta del Excel)", "warn")
                    extra_results.append({
                        "status": "OCULTO",
                        "region": "",
                        "city": text,
                        "dealer": "",
                        "fails": ["Ciudad presente en el form pero declarada SOLO en fila(s) OCULTA(s) del Excel"],
                    })
                else:
                    log(f"  EXTRA CIUDAD: {text}", "warn")
                    extra_results.append({
                        "status": "EXTRA",
                        "region": "",
                        "city": text,
                        "dealer": "",
                        "fails": ["Ciudad extra (no declarada en el Excel)"],
                    })
            for norm, texts in seen_in_form.items():
                if len(texts) > 1:
                    log(f"  DUPLICADA CIUDAD: {texts[0]} x{len(texts)}", "warn")
                    extra_results.append({
                        "status": "DUPLICADO",
                        "region": "",
                        "city": texts[0],
                        "dealer": "",
                        "fails": [f"Ciudad aparece {len(texts)} veces en el dropdown"],
                    })
            for norm_city in (expected_cities | hidden_cities):
                if norm_city in seen_in_form:
                    cities_to_check.append(("", norm_city))

    # --------------------------------------------------------------------------
    # FASE C: Chequeo de Dealer (sólo si has_dealer es True)
    # --------------------------------------------------------------------------
    def _append_dealer_classification(option_texts, key, reg_orig_text, city_orig_text):
        duplicated, extras, notas, ocultos = _classify_dealer_options(option_texts, key)
        for text in extras:
            log(f"  EXTRA DEALER: {text} ({reg_orig_text} / {city_orig_text})", "warn")
            extra_results.append({
                "status": "EXTRA", "region": reg_orig_text, "city": city_orig_text, "dealer": text,
                "fails": [f"Dealer extra para '{reg_orig_text} / {city_orig_text}'".strip(" /")],
            })
        for text in ocultos:
            log(f"  OCULTO DEALER: {text} ({reg_orig_text} / {city_orig_text}) — solo en fila oculta del Excel", "warn")
            extra_results.append({
                "status": "OCULTO", "region": reg_orig_text, "city": city_orig_text, "dealer": text,
                "fails": [f"Dealer presente en el form para '{reg_orig_text} / {city_orig_text}' pero declarado SOLO en fila(s) OCULTA(s) del Excel".strip(" /")],
            })
        for text, count in duplicated:
            log(f"  DUPLICADO DEALER: {text} x{count} ({reg_orig_text} / {city_orig_text})", "warn")
            extra_results.append({
                "status": "DUPLICADO", "region": reg_orig_text, "city": city_orig_text, "dealer": text,
                "fails": [f"Dealer aparece {count} veces en el dropdown de '{reg_orig_text} / {city_orig_text}'".strip(" /")],
            })
        for text, excel_orig in notas:
            log(f"  NOTA DEALER: '{text}' matchea '{excel_orig}' con diferencias menores "
                f"({reg_orig_text} / {city_orig_text})", "info")
            extra_results.append({
                "status": "NOTA", "region": reg_orig_text, "city": city_orig_text, "dealer": excel_orig,
                "fails": [], "name_disclaimer": match_dealer_names(excel_orig, text)[1],
            })

    if has_dealer:
        log("Chequeando EXTRAS y DUPLICADOS en Dealers...", "info")
        if has_region and has_city:
            for norm_reg, norm_city in cities_to_check:
                _check_stop(stop_flag)
                reg_orig_text = orig_region_by_norm[norm_reg]
                city_orig_text = orig_city_by_norm[norm_city]
                if progress_cb:
                    progress_cb(1, 1, f"{reg_orig_text} / {city_orig_text}")

                ok, _, _ = _select_by_text_robust(driver, level_ids["region"], reg_orig_text)
                if not ok:
                    continue
                ok, _, _ = _select_by_text_robust(driver, level_ids["city"], city_orig_text, wait_for_option=True)
                if not ok:
                    continue
                time.sleep(0.6)  # Pausa estable para repoblar dealers

                option_texts = _read_valid_option_texts(driver, level_ids["dealer"])
                _append_dealer_classification(option_texts, (norm_reg, norm_city), reg_orig_text, city_orig_text)
        elif not has_region and has_city:
            for _, norm_city in cities_to_check:
                _check_stop(stop_flag)
                city_orig_text = orig_city_by_norm[norm_city]
                if progress_cb:
                    progress_cb(1, 1, f"Ciudad: {city_orig_text}")

                ok, _, _ = _select_by_text_robust(driver, level_ids["city"], city_orig_text)
                if not ok:
                    continue
                time.sleep(0.6)

                option_texts = _read_valid_option_texts(driver, level_ids["dealer"])
                _append_dealer_classification(option_texts, ("", norm_city), "", city_orig_text)
        elif has_region and not has_city:
            for norm_reg in regions_to_check:
                _check_stop(stop_flag)
                reg_orig_text = orig_region_by_norm[norm_reg]
                if progress_cb:
                    progress_cb(1, 1, f"Región: {reg_orig_text}")

                ok, _, _ = _select_by_text_robust(driver, level_ids["region"], reg_orig_text)
                if not ok:
                    continue
                time.sleep(0.6)

                option_texts = _read_valid_option_texts(driver, level_ids["dealer"])
                _append_dealer_classification(option_texts, (norm_reg, ""), reg_orig_text, "")
        else:
            # Sin región ni ciudad, dealers globales
            option_texts = _read_valid_option_texts(driver, level_ids["dealer"])
            _append_dealer_classification(option_texts, ("", ""), "", "")

    return extra_results


# ──────────────────────────────────────────────────────────────────────────────
# Captura de pantalla con banner de URL + ZIP
# ──────────────────────────────────────────────────────────────────────────────
def _ensure_window_covers_content(driver):
    """Agranda la ventana SOLO si hay overflow horizontal real (el contenido es más ancho
    que el viewport y quedaría recortado en la captura). Compara contra el viewport interior
    (clientWidth) y no contra el ancho de la ventana: un layout fluido siempre llena el 100%
    del viewport y NO debe disparar el agrandado (antes esto hacía crecer la ventana +60px
    en cada captura, en loop). Tope 2200px."""
    try:
        doc_w, client_w = driver.execute_script(
            "return [Math.max(document.body.scrollWidth || 0, document.documentElement.scrollWidth || 0),"
            " document.documentElement.clientWidth || 0];"
        )
        if not doc_w or not client_w or doc_w <= client_w + 5:
            return  # sin overflow horizontal: no tocar la ventana
        size = driver.get_window_size()
        extra = int(doc_w) - int(client_w)
        new_w = min(size.get("width", 1366) + extra + 20, 2200)
        if new_w > size.get("width", 0):
            driver.set_window_size(new_w, size.get("height", 1080))
    except Exception:
        pass


def banner_lines_for_result(result, has_region=True, has_city=True, has_dealer=True):
    """Arma las líneas extra del banner para una captura del Comparador: qué combinación
    se revisó (región/ciudad/dealer según niveles activos) y el estado PASS/FAIL."""
    partes = []
    if has_region and result.get("region"):
        partes.append(f"Región: {result['region']}")
    if has_city and result.get("city"):
        partes.append(f"Ciudad: {result['city']}")
    if has_dealer and result.get("dealer"):
        partes.append(f"Dealer: {result['dealer']}")
    if result.get("modelo"):
        partes.append(f"Modelo: {result['modelo']}")
    lines = []
    if partes:
        lines.append((f"Revisado:        {' · '.join(partes)}", (220, 220, 100)))
    status = result.get("status", "")
    color = (80, 200, 80) if status == "PASS" else (220, 80, 80)
    if result.get("fails"):
        detalle = " — " + " | ".join(result.get("fails", []))
    elif result.get("detalle_info"):
        detalle = " — " + result["detalle_info"]
    else:
        detalle = ""
    lines.append((f"Estado:          {status}{detalle}"[:220], color))
    return lines


def capture_result_screenshot(driver, screenshot_dir, filename, form_url="", landing_url="",
                              extra_lines=None):
    """Toma una captura de página completa con las URLs del form y landing pegadas en un banner
    superior (reusa ScreenshotManager). extra_lines: líneas extra del banner, ej. las de
    banner_lines_for_result (revisado + estado PASS/FAIL)."""
    os.makedirs(screenshot_dir, exist_ok=True)
    _ensure_window_covers_content(driver)
    manager = ScreenshotManager(driver, screenshot_dir)
    # Pasar ambas URLs para que el banner muestre las dos (si aplica)
    if landing_url:
        manager.url_landing = landing_url
    manager.url_form_esperado = form_url
    manager.extra_lines = list(extra_lines or [])
    manager.take_full_page_screenshot(filename)
    manager._add_url_banner(os.path.join(screenshot_dir, filename))
    return os.path.join(screenshot_dir, filename)


# ──────────────────────────────────────────────────────────────────────────────
# Evidencia de dropdown desplegado DENTRO del form (para verificar exclusión).
# El popup nativo del <select> es UI del sistema operativo y Selenium no puede
# fotografiarlo, pero sí se puede convertir temporalmente el <select> en una lista
# abierta (atributo size = cantidad de opciones): el dropdown queda desplegado en su
# lugar real dentro del formulario mostrando TODAS las opciones, se captura, y se
# restaura. Si el buscado aparece, se resalta en amarillo.
# ──────────────────────────────────────────────────────────────────────────────
_EXPAND_SELECT_JS = """
var id = arguments[0], searched = (arguments[1] || '').trim().toLowerCase();
var el = document.getElementById(id);
if (!el || !el.options) return -1;
el.setAttribute('data-ocioso-size-prev', el.getAttribute('size') || '');
el.size = Math.max(2, el.options.length);
el.style.height = 'auto';
el.style.maxHeight = 'none';
el.style.overflow = 'visible';
el.style.zIndex = '999999';
el.style.position = 'relative';
// Que ningún contenedor recorte la lista desplegada
var p = el.parentElement, hops = 0;
while (p && p !== document.body && hops < 6) {
    if (getComputedStyle(p).overflow !== 'visible') p.style.overflow = 'visible';
    p.style.maxHeight = 'none';
    p = p.parentElement; hops++;
}
for (var i = 0; i < el.options.length; i++) {
    var o = el.options[i];
    if (searched && (o.text || '').trim().toLowerCase() === searched) {
        o.style.background = '#ffd873';
        o.style.color = '#000000';
        o.style.fontWeight = 'bold';
        o.setAttribute('data-ocioso-marked', '1');
    }
}
try { el.scrollIntoView({block: 'center'}); } catch (e) {}
return el.options.length;
"""

_RESTORE_SELECT_JS = """
var id = arguments[0];
var el = document.getElementById(id);
if (!el) return;
var prev = el.getAttribute('data-ocioso-size-prev');
if (prev) el.setAttribute('size', prev); else el.removeAttribute('size');
el.removeAttribute('data-ocioso-size-prev');
el.style.height = ''; el.style.maxHeight = ''; el.style.overflow = '';
el.style.zIndex = ''; el.style.position = '';
for (var i = 0; i < el.options.length; i++) {
    var o = el.options[i];
    if (o.getAttribute('data-ocioso-marked')) {
        o.style.background = ''; o.style.color = ''; o.style.fontWeight = '';
        o.removeAttribute('data-ocioso-marked');
    }
}
"""


def evidence_level_for_result(result, has_region=True, has_city=True, has_dealer=True):
    """Decide en QUÉ dropdown desplegar la evidencia de exclusión: el primer nivel de la
    cadena región→ciudad→dealer que NO se encontró en el form (si la región no está, la
    ciudad y el dealer tampoco pueden estar — la evidencia es el dropdown de regiones).
    Si todo se encontró (dealer presente = FAIL de exclusión), la evidencia es el dropdown
    de dealer con el buscado resaltado. Sólo considera los niveles activos.
    Devuelve (nivel_label, select_key, texto_buscado, found_bool)."""
    if has_region and result.get("region") and not result.get("region_found"):
        return "Región", "region", result.get("region", ""), False
    if has_city and result.get("city") and not result.get("city_found"):
        return "Ciudad", "city", result.get("city", ""), False
    if has_dealer:
        return "Dealer", "dealer", result.get("dealer", ""), bool(result.get("dealer_found"))
    if has_city:
        return "Ciudad", "city", result.get("city", ""), bool(result.get("city_found"))
    return "Región", "region", result.get("region", ""), bool(result.get("region_found"))


_RESIZE_IFRAME_JS = """
var iframe = arguments[0], newHeight = arguments[1];
iframe.setAttribute('data-ocioso-h-prev', iframe.style.height || '');
iframe.style.height = newHeight + 'px';
iframe.style.maxHeight = 'none';
iframe.setAttribute('scrolling', 'no');
"""

_RESTORE_IFRAME_JS = """
var iframe = arguments[0];
iframe.style.height = iframe.getAttribute('data-ocioso-h-prev') || '';
iframe.style.maxHeight = '';
iframe.removeAttribute('data-ocioso-h-prev');
iframe.removeAttribute('scrolling');
"""


def capture_dropdown_evidence(driver, select_id, screenshot_dir, filename, title="", searched_text="",
                               found=None, form_url="", landing_url="", extra_lines=None):
    """Captura el formulario con el <select> indicado DESPLEGADO en su lugar: convierte
    temporalmente el dropdown en lista abierta (atributo size) para que se vean TODAS sus
    opciones dentro del form real — el popup nativo del <select> no es fotografiable. Si el
    texto buscado aparece entre las opciones queda resaltado. Como el form suele vivir en un
    iframe de altura fija que recortaría la lista desplegada, también se agranda el iframe
    desde la página padre a la altura del contenido expandido, se captura la página completa
    desde el documento principal, y se restaura todo (select, iframe y contexto)."""
    os.makedirs(screenshot_dir, exist_ok=True)
    expanded = False
    iframe_resized = None
    try:
        try:
            n = driver.execute_script(_EXPAND_SELECT_JS, select_id, searched_text or "")
            expanded = isinstance(n, (int, float)) and n >= 0
        except Exception:
            expanded = False
        time.sleep(0.3)  # dejar que el layout se acomode con la lista abierta

        # Altura real del documento del form con la lista abierta (contexto actual = form)
        try:
            inner_height = driver.execute_script(
                "return Math.max(document.documentElement.scrollHeight || 0,"
                " document.body ? document.body.scrollHeight : 0);"
            ) or 0
            driver.execute_script("window.scrollTo(0, 0);")
        except Exception:
            inner_height = 0

        # Si el form está embebido en un iframe, agrandarlo desde la página padre para que
        # la lista desplegada no quede recortada por la altura fija del iframe.
        try:
            driver.switch_to.default_content()
            iframe = _find_form_iframe(driver, form_url)
        except Exception:
            iframe = None
        if iframe is not None and inner_height:
            try:
                driver.execute_script(_RESIZE_IFRAME_JS, iframe, int(inner_height) + 30)
                iframe_resized = iframe
                time.sleep(0.3)
            except Exception:
                iframe_resized = None
        elif iframe is None:
            # No hay iframe reconocible: el form está en la página principal, seguimos ahí.
            pass

        _ensure_window_covers_content(driver)
        manager = ScreenshotManager(driver, screenshot_dir)
        if landing_url:
            manager.url_landing = landing_url
        manager.url_form_esperado = form_url
        manager.extra_lines = list(extra_lines or [])
        manager.take_full_page_screenshot(filename)
        manager._add_url_banner(os.path.join(screenshot_dir, filename))
    finally:
        # Restaurar iframe (estamos en el documento principal desde el resize)
        if iframe_resized is not None:
            try:
                driver.execute_script(_RESTORE_IFRAME_JS, iframe_resized)
            except Exception:
                pass
        # Volver al contexto del form para restaurar el select y que la próxima fila navegue
        try:
            iframe = _find_form_iframe(driver, form_url)
            if iframe is not None:
                driver.switch_to.frame(iframe)
        except Exception:
            pass
        if expanded:
            try:
                driver.execute_script(_RESTORE_SELECT_JS, select_id)
            except Exception:
                pass
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
def export_results_excel(results, output_path=None, pais="", hidden_rows=None, hidden_columns=None):
    """hidden_rows / hidden_columns: filas y columnas OCULTAS del Excel de dealers de ORIGEN
    (detect_hidden_rows / detect_hidden_columns). Se listan en el Resumen, y todo resultado
    relacionado con filas ocultas se pinta de NARANJA (estado OCULTO o flag hidden_in_excel)."""
    if output_path is None:
        results_dir = os.path.join(RESULTS_DIR, "dealer_comparator")
        os.makedirs(results_dir, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(results_dir, f"dealer_comparator_{pais}_{stamp}.xlsx")

    from openpyxl.styles import Alignment, Border, Side

    thin_border = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )
    header_border = Border(
        left=Side(style="thin", color="FFFFFF"),
        right=Side(style="thin", color="FFFFFF"),
        top=Side(style="thin", color="FFFFFF"),
        bottom=Side(style="medium", color="7D4E9F"),
    )

    wb = Workbook()
    ws = wb.active
    ws.title = "Resultados"

    headers = ["Estado", "URL Landing", "URL Form", "Modelo", "Región", "Ciudad", "Dealer",
               "BAC Excel", "BAC OK", "Detalle", "Nota Nombre", "Fila Excel"]
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.border = header_border
        cell.alignment = Alignment(horizontal="center", vertical="center")

    fill_by_status = {
        "PASS": _PASS_FILL, "FAIL": _FAIL_FILL, "EXTRA": _EXTRA_FILL,
        "MISSING": _MISSING_FILL, "DUPLICADO": _DUPLICATE_FILL, "NOTA": _NOTA_FILL,
        "OCULTO": _OCULTO_FILL,
    }

    for r in results:
        ws.append([
            r.get("status", ""),
            r.get("url_landing", ""),
            r.get("url_form", ""),
            r.get("modelo", ""),
            r.get("region", ""),
            r.get("city", ""),
            r.get("dealer", ""),
            r.get("bac_excel", ""),
            {True: "OK", False: "MISMATCH", None: ""}.get(r.get("bac_ok")),
            " | ".join(r.get("fails", [])) if r.get("fails") else (r.get("detalle_info") or ""),
            r.get("name_disclaimer", "") or "",
            r.get("fila", ""),
        ])
        row_idx = ws.max_row
        fill = fill_by_status.get(r.get("status"))
        # Todo lo relacionado con filas OCULTAS del Excel va en naranja, pise lo que pise
        if r.get("hidden_in_excel"):
            fill = _OCULTO_FILL
        for cell in ws[row_idx]:
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if fill:
                cell.fill = fill
        # Nota de nombre presente en fila PASS: resaltar solo esa celda para no
        # confundir con FAIL, ya que el dealer sí se encontró.
        if r.get("name_disclaimer") and r.get("status") == "PASS" and not r.get("hidden_in_excel"):
            ws.cell(row=row_idx, column=11).fill = _NOTA_FILL

    # Auto-fit column widths (basado en contenido, con mínimos y máximos razonables)
    col_widths = {1: 10, 2: 45, 3: 45, 4: 16, 5: 22, 6: 22, 7: 30, 8: 14, 9: 10, 10: 50, 11: 55, 12: 10}
    for col_idx, width in col_widths.items():
        col_letter = ws.cell(row=1, column=col_idx).column_letter
        ws.column_dimensions[col_letter].width = width

    # Fijar encabezado (freeze panes)
    ws.freeze_panes = "A2"

    # Autofiltro
    ws.auto_filter.ref = ws.dimensions

    # ── Hoja Resumen ──
    total = len(results)
    counts = {"PASS": 0, "FAIL": 0, "EXTRA": 0, "MISSING": 0, "DUPLICADO": 0, "NOTA": 0, "OCULTO": 0}
    for r in results:
        counts[r.get("status", "FAIL")] = counts.get(r.get("status", "FAIL"), 0) + 1
    disclaimer_count = sum(1 for r in results if r.get("name_disclaimer"))

    summary_ws = wb.create_sheet("Resumen")
    summary_data = [
        ["País", pais],
        ["Fecha", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        ["", ""],
        ["Total chequeados", total],
        ["", ""],
        ["🟢 PASS", counts.get("PASS", 0)],
        ["🔴 FAIL", counts.get("FAIL", 0)],
        ["🟡 EXTRA (en el form, no declarado en el Excel)", counts.get("EXTRA", 0)],
        ["🔵 DUPLICADO (repetido en el dropdown del form)", counts.get("DUPLICADO", 0)],
        ["🟠 OCULTO (en el form, pero declarado solo en filas OCULTAS del Excel)", counts.get("OCULTO", 0)],
        ["🟣 MISSING", counts.get("MISSING", 0)],
        ["🔷 NOTA (dealer del form no está en el Excel, solo difiere en nombre)", counts.get("NOTA", 0)],
        ["⚠ PASS con nombre distinto en detalles menores", disclaimer_count],
    ]

    # Agregar URLs únicas usadas
    urls_form = sorted(set(r.get("url_form", "") for r in results if r.get("url_form")))
    urls_landing = sorted(set(r.get("url_landing", "") for r in results if r.get("url_landing")))
    if urls_form or urls_landing:
        summary_data.append(["", ""])
        summary_data.append(["URLs procesadas", ""])
        for url in urls_landing:
            summary_data.append(["  Landing", url])
        for url in urls_form:
            summary_data.append(["  Form", url])

    # Avisos de calidad de datos del Excel de dealers (no afectan la comparación en sí)
    if hidden_rows:
        summary_data.append(["", ""])
        summary_data.append(["⚠ Filas OCULTAS en el Excel de dealers", len(hidden_rows)])
        summary_data.append(["  Número de fila(s)", ", ".join(str(r) for r in sorted(hidden_rows))])
    if hidden_columns:
        summary_data.append(["", ""])
        summary_data.append(["⚠ Columnas OCULTAS en el Excel de dealers", len(hidden_columns)])
        for hc in hidden_columns:
            label = f"  Columna {hc['column']}" + (f" ({hc['header']})" if hc.get("header") else "")
            summary_data.append([label, ""])

    for row_data in summary_data:
        summary_ws.append(row_data)

    # Formato básico de la hoja resumen
    summary_ws.column_dimensions["A"].width = 22
    summary_ws.column_dimensions["B"].width = 60
    for row in summary_ws.iter_rows(min_row=1, max_row=summary_ws.max_row, max_col=2):
        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center")
    # Encabezados en negrita
    for cell in [summary_ws.cell(1, 1), summary_ws.cell(1, 2),
                 summary_ws.cell(4, 1), summary_ws.cell(4, 2)]:
        cell.font = Font(bold=True)

    wb.save(output_path)
    return output_path
