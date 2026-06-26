import logging
import re
import time
import unicodedata
from datetime import datetime

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException

from core.browser_manager import BrowserManager
from core.field_dependencies import FIELD_DEPENDENCIES

from selenium.webdriver.support.ui import Select

from .error_message_validator import generar_inputs_test, read_error_near_element, validar_error_ui
from .text_field_validator import validate_text_field


LOGGER = logging.getLogger(__name__)

DEPENDENCY_DROPDOWN_TIMEOUT = 8.0
DEPENDENCY_DROPDOWN_POLL_INTERVAL = 0.2


def _sanitize_url(value):
    return str(value or "").strip()


def _wait_document_ready(driver, timeout):
    WebDriverWait(driver, timeout).until(
        lambda current_driver: current_driver.execute_script("return document.readyState") == "complete"
    )


def _pre_scroll_for_dynamic_content(driver):
    try:
        total_height = driver.execute_script("return document.body.parentNode.scrollHeight")
        viewport_height = driver.execute_script("return window.innerHeight")
        current_position = 0
        scroll_step = max(1, int(viewport_height * 0.8))

        while current_position < total_height:
            driver.execute_script("window.scrollTo(0, arguments[0]);", current_position)
            time.sleep(0.25)
            current_position += scroll_step
            if current_position > total_height * 3:
                break

        driver.execute_script("window.scrollTo(0, document.body.parentNode.scrollHeight);")
        time.sleep(0.4)
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(0.4)
    except Exception:
        pass


def _find_form_iframe(driver, expected_form_url):
    expected_form_url = _sanitize_url(expected_form_url)
    if not expected_form_url:
        return None

    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    for iframe in iframes:
        try:
            src = iframe.get_attribute("src") or ""
            if expected_form_url in src:
                return iframe
        except Exception:
            continue
    return None


def _is_text_input_candidate(element):
    try:
        tag_name = (element.tag_name or "").lower()
        if tag_name == "textarea":
            return True
        if tag_name != "input":
            return False

        input_type = (element.get_attribute("type") or "text").strip().lower()
        return input_type in {"text", "email", "tel", "search", "url", "password", "number"}
    except Exception:
        return False


def _get_visible_element_by_id(driver, element_id):
    try:
        element = driver.find_element(By.ID, element_id)
        if element.is_displayed() and _is_text_input_candidate(element):
            return element
    except Exception:
        return None
    return None


def _collect_visible_fields(driver, normalized_fields):
    visible = []
    for field_config in normalized_fields:
        if field_config.get("dropdown"):
            continue
        element_id = field_config["element_id"]
        element = _get_visible_element_by_id(driver, element_id)
        if element is None:
            continue
        visible.append((field_config, element))
    return visible


def _get_element_by_id_any_tag(driver, element_id):
    """Busca cualquier elemento visible por ID, sin restriccion de tag."""
    try:
        element = driver.find_element(By.ID, element_id)
        if element.is_displayed():
            return element
    except Exception:
        pass
    return None


def _collect_visible_dropdowns(driver, normalized_fields):
    """Devuelve los campos marcados como dropdown que tienen un elemento visible en la pagina."""
    visible = []
    for field_config in normalized_fields:
        element = _get_element_by_id_any_tag(driver, field_config["element_id"])
        if element is None:
            continue
        # Robustez: si la regla viene sin dropdown=True pero el elemento es <select>,
        # se procesa igual como dropdown para no bloquear el avance del step.
        tag_name = (element.tag_name or "").lower()
        if not field_config.get("dropdown") and tag_name != "select":
            continue
        visible.append((field_config, element))
    return visible


def _collect_fallback_dom_dropdowns(driver, normalized_fields):
    """Detecta <select> visibles no presentes en reglas y los devuelve como dropdowns fallback."""
    fallback = []
    configured_ids = {fc["element_id"] for fc in normalized_fields}
    try:
        dom_selects = driver.find_elements(By.CSS_SELECTOR, "select[id]")
    except Exception:
        return fallback

    for element in dom_selects:
        try:
            if not element.is_displayed():
                continue
            element_id = (element.get_attribute("id") or "").strip()
            if not element_id or element_id in configured_ids:
                continue
            fallback.append(
                (
                    {
                        "field_name": element_id,
                        "descripcion": f"Dropdown DOM no mapeado ({element_id})",
                        "element_id": element_id,
                        "regex_full": "",
                        "regex_char": "",
                        "test_text": "",
                        "teclado_mobile": False,
                        "rules": {},
                        "error_messages": {},
                        "error_message_patterns": [],
                        "error_config": {},
                        "error_priority": ["required"],
                        "dropdown": True,
                        "dropdown_error_message": "",
                        "dependencies": [],
                    },
                    element,
                )
            )
        except Exception:
            continue
    return fallback


def _sort_dropdowns_by_dependency(dropdown_list, dependencies):
    """Ordena (field_config, element) respetando dependencias: el padre siempre antes que el hijo."""
    # dependencies formato: {parent_id: child_id}
    # Construimos mapa inverso: child_id -> parent_id
    child_to_parent = {child: parent for parent, child in dependencies.items()}

    id_to_item = {fc["element_id"]: (fc, el) for fc, el in dropdown_list}
    ordered = []
    visited = set()

    def _visit(eid):
        if eid in visited:
            return
        visited.add(eid)
        parent_id = child_to_parent.get(eid)
        if parent_id and parent_id in id_to_item:
            _visit(parent_id)
        if eid in id_to_item:
            ordered.append(id_to_item[eid])

    for fc, _el in dropdown_list:
        _visit(fc["element_id"])

    return ordered


def _is_placeholder_text(option_text):
    normalized = _normalize_text(option_text)
    if not normalized:
        return True

    placeholder_keywords = (
        "seleccione",
        "selecciona",
        "seleccionar",
        "seleccion",
        "selecione",
        "selecionar",
        "selecao",
        "escolha",
        "escolher",
        "elija",
        "elegir",
        "opcion",
        "opciones",
        "select",
        "choose",
        "please",
        "porfavor",
        "obrigatorio",
        "required",
    )
    return any(keyword in normalized for keyword in placeholder_keywords)


def _is_option_disabled(option_element):
    try:
        disabled_attr = option_element.get_attribute("disabled")
        if disabled_attr is None:
            return False
        normalized = str(disabled_attr).strip().lower()
        return normalized in ("", "true", "disabled")
    except Exception:
        return False


def _get_valid_select_options(select_element):
    valid_options = []
    try:
        select = Select(select_element)
    except Exception:
        return valid_options

    for opt in select.options:
        opt_text = (opt.text or "").strip()
        if not opt_text:
            continue
        if _is_option_disabled(opt):
            continue
        if _is_placeholder_text(opt_text):
            continue
        valid_options.append(opt)
    return valid_options


def _select_first_available_option(driver, element):
    """Intenta seleccionar la primera opcion util en dropdown nativo o custom."""
    try:
        tag = (element.tag_name or "").lower()

        if tag == "select":
            valid_options = _get_valid_select_options(element)

            for option in valid_options:
                try:
                    value = (option.get_attribute("value") or "").strip()
                    text = (option.text or "").strip()

                    select = Select(element)
                    if value:
                        select.select_by_value(value)
                    else:
                        select.select_by_visible_text(text)

                    driver.execute_script(
                        "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));"
                        "arguments[0].dispatchEvent(new Event('input', { bubbles: true }));",
                        element,
                    )
                    return True
                except Exception:
                    continue
            return False

        # Fallback para dropdowns custom (div/input con lista de opciones)
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
        time.sleep(0.2)
        element.click()
        time.sleep(0.3)

        option_selectors = [
            "[role='option']",
            ".dropdown-item",
            ".select2-results__option",
            ".vs__dropdown-option",
            ".ng-option",
            "li",
        ]
        for selector in option_selectors:
            try:
                options = driver.find_elements(By.CSS_SELECTOR, selector)
                for option in options:
                    try:
                        if not option.is_displayed() or not option.is_enabled():
                            continue
                        option_text = (option.text or option.get_attribute("innerText") or "").strip()
                        if _is_placeholder_text(option_text):
                            continue
                        if not option_text:
                            continue
                        driver.execute_script("arguments[0].click();", option)
                        return True
                    except Exception:
                        continue
            except Exception:
                continue

        try:
            element.send_keys(Keys.ARROW_DOWN)
            time.sleep(0.1)
            element.send_keys(Keys.ENTER)
            return True
        except Exception:
            return False
    except Exception:
        return False


def _wait_child_dropdown_options(driver, child_id, timeout=4.0, poll_interval=0.2):
    """Espera que el dropdown hijo tenga al menos una opcion util (no placeholder)."""
    deadline = time.time() + max(0.2, float(timeout))
    while time.time() < deadline:
        try:
            child_element = _get_element_by_id_any_tag(driver, child_id)
            if child_element is None:
                time.sleep(max(0.05, float(poll_interval)))
                continue

            if (child_element.tag_name or "").lower() != "select":
                return True

            valid = _get_valid_select_options(child_element)
            if valid:
                return True
        except Exception:
            # Ignorar errores transitorios durante la espera.
            pass

        time.sleep(max(0.05, float(poll_interval)))

    LOGGER.debug("Timeout esperando opciones en dropdown hijo '%s'", child_id)
    return False


def _has_valid_dropdown_selection(element):
    """Indica si un <select> tiene una opción útil seleccionada (no placeholder)."""
    try:
        if (element.tag_name or "").lower() != "select":
            return False
        selected_option = Select(element).first_selected_option
        if not selected_option:
            return False
        text = (selected_option.text or "").strip()
        if not text:
            return False
        if _is_option_disabled(selected_option):
            return False
        return not _is_placeholder_text(text)
    except Exception:
        return False


def _get_parent_chain(field_id, dependencies):
    """Devuelve la cadena de padres de un campo (de raíz a padre directo)."""
    child_to_parent = {child: parent for parent, child in dependencies.items()}
    chain = []
    seen = set()
    current = field_id

    while current in child_to_parent:
        parent = child_to_parent[current]
        if parent in seen:
            break
        seen.add(parent)
        chain.append(parent)
        current = parent

    chain.reverse()
    return chain


def _ensure_dropdown_parent_context(driver, field_id, dependencies):
    """Selecciona padres dependientes para validar correctamente hijos como required."""
    parent_chain = _get_parent_chain(field_id, dependencies)
    if not parent_chain:
        return True

    for parent_id in parent_chain:
        parent_element = _get_element_by_id_any_tag(driver, parent_id)
        if parent_element is None:
            return False

        if _has_valid_dropdown_selection(parent_element):
            continue

        if not _select_first_available_option(driver, parent_element):
            return False

        child_id = dependencies.get(parent_id)
        if child_id:
            _wait_child_dropdown_options(
                driver,
                child_id,
                timeout=DEPENDENCY_DROPDOWN_TIMEOUT,
                poll_interval=DEPENDENCY_DROPDOWN_POLL_INTERVAL,
            )

    return True


def _normalize_text(value):
    """Normaliza texto: minúsculas, sin acentos, sin guiones ni underscore."""
    text = str(value or "").strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return text


def _normalize_dependency_token(value):
    """Normaliza tokens para comparar valores de dependencias de forma robusta."""
    normalized = _normalize_text(value)
    return "".join(ch for ch in normalized if ch.isalnum())


_DEPENDENCY_VALUE_ALIASES = {
    "rut": ["ruc"],
    "ruc": ["rut"],
}


def _build_dependency_candidate_tokens(expected_value):
    """Genera tokens alternativos permitidos para resolver dependencias por valor."""
    base_token = _normalize_dependency_token(expected_value)
    if not base_token:
        return set()

    candidates = {base_token}
    for alias in _DEPENDENCY_VALUE_ALIASES.get(base_token, []):
        alias_token = _normalize_dependency_token(alias)
        if alias_token:
            candidates.add(alias_token)
    return candidates


def _normalize_rule_dependencies(raw_dependencies):
    """Normaliza la lista de dependencias de una regla de validación."""
    if not raw_dependencies:
        return []

    if isinstance(raw_dependencies, dict):
        raw_dependencies = [raw_dependencies]
    if not isinstance(raw_dependencies, list):
        return []

    normalized = []
    for raw_dependency in raw_dependencies:
        if not isinstance(raw_dependency, dict):
            continue

        dependency_id = str(
            raw_dependency.get("element_id")
            or raw_dependency.get("id")
            or raw_dependency.get("field_id")
            or ""
        ).strip()
        dependency_value = str(
            raw_dependency.get("value")
            or raw_dependency.get("option")
            or raw_dependency.get("valor")
            or ""
        ).strip()

        if not dependency_id or not dependency_value:
            continue

        normalized.append(
            {
                "element_id": dependency_id,
                "value": dependency_value,
            }
        )

    return normalized


def _select_option_by_dependency_value(driver, element, expected_value):
    """Selecciona una opción específica en un dropdown por texto o value."""
    expected_tokens = _build_dependency_candidate_tokens(expected_value)
    if not expected_tokens:
        return False

    try:
        tag = (element.tag_name or "").lower()

        if tag == "select":
            try:
                select = Select(element)
            except Exception:
                return False

            exact_match = None
            contains_match = None
            for option in select.options:
                if _is_option_disabled(option):
                    continue

                option_text = (option.text or "").strip()
                option_value = (option.get_attribute("value") or "").strip()
                if _is_placeholder_text(option_text):
                    continue

                text_token = _normalize_dependency_token(option_text)
                value_token = _normalize_dependency_token(option_value)

                if text_token in expected_tokens or value_token in expected_tokens:
                    exact_match = option
                    break

                if any(token in text_token or token in value_token for token in expected_tokens):
                    contains_match = contains_match or option

            selected_option = exact_match or contains_match
            if selected_option is None:
                return False

            selected_value = (selected_option.get_attribute("value") or "").strip()
            selected_text = (selected_option.text or "").strip()
            if selected_value:
                select.select_by_value(selected_value)
            else:
                select.select_by_visible_text(selected_text)

            driver.execute_script(
                "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));"
                "arguments[0].dispatchEvent(new Event('input', { bubbles: true }));",
                element,
            )
            return True

        # Fallback para dropdowns custom
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
        time.sleep(0.2)
        element.click()
        time.sleep(0.3)

        option_selectors = [
            "[role='option']",
            ".dropdown-item",
            ".select2-results__option",
            ".vs__dropdown-option",
            ".ng-option",
            "li",
        ]

        exact_option = None
        contains_option = None
        for selector in option_selectors:
            try:
                options = driver.find_elements(By.CSS_SELECTOR, selector)
            except Exception:
                continue

            for option in options:
                try:
                    if not option.is_displayed() or not option.is_enabled():
                        continue
                    option_text = (option.text or option.get_attribute("innerText") or "").strip()
                    if not option_text:
                        continue
                    if _is_placeholder_text(option_text):
                        continue

                    token = _normalize_dependency_token(option_text)
                    if token in expected_tokens:
                        exact_option = option
                        break
                    if any(expected_token in token for expected_token in expected_tokens) and contains_option is None:
                        contains_option = option
                except Exception:
                    continue

            if exact_option is not None:
                break

        selected_custom_option = exact_option or contains_option
        if selected_custom_option is None:
            return False

        driver.execute_script("arguments[0].click();", selected_custom_option)
        return True
    except Exception:
        return False


def _apply_field_dependencies(driver, field_config, dependencies_map):
    """Aplica dependencias configuradas para una regla antes de validarla."""
    dependencies = list(field_config.get("dependencies") or [])
    if not dependencies:
        return True, ""

    for dependency in dependencies:
        dependency_id = str(dependency.get("element_id") or "").strip()
        dependency_value = str(dependency.get("value") or "").strip()
        if not dependency_id or not dependency_value:
            continue

        context_ready = _ensure_dropdown_parent_context(driver, dependency_id, dependencies_map)
        if not context_ready:
            return False, f"No se pudo preparar contexto para dependencia '{dependency_id}'"

        dependency_element = _get_element_by_id_any_tag(driver, dependency_id)
        if dependency_element is None:
            return False, f"No se encontró el campo dependiente '{dependency_id}'"

        selected_ok = _select_option_by_dependency_value(driver, dependency_element, dependency_value)
        if not selected_ok:
            return (
                False,
                f"No se pudo seleccionar '{dependency_value}' en dependencia '{dependency_id}'",
            )

        child_id = dependencies_map.get(dependency_id)
        if child_id:
            child_ready = _wait_child_dropdown_options(
                driver,
                child_id,
                timeout=DEPENDENCY_DROPDOWN_TIMEOUT,
                poll_interval=DEPENDENCY_DROPDOWN_POLL_INTERVAL,
            )
            if not child_ready:
                return False, f"Timeout esperando opciones de '{child_id}' tras '{dependency_id}'"

    return True, ""


def _expects_error_for_empty_input(field_config):
    """Indica si la regla espera mensaje de error cuando el input está vacío."""
    rules = dict(field_config.get("rules") or {})
    if bool(rules.get("required", False)):
        return True

    patterns = list(field_config.get("error_message_patterns") or [])
    for pattern_rule in patterns:
        if not isinstance(pattern_rule, dict):
            continue
        regex_text = str(pattern_rule.get("regex") or "").strip()
        expected_message = str(pattern_rule.get("mensaje") or pattern_rule.get("message") or "").strip()
        if not regex_text or not expected_message:
            continue
        try:
            compiled = re.compile(regex_text)
        except re.error:
            continue
        if compiled.fullmatch("") or compiled.match(""):
            return True

    return False


def _build_field_config_with_trigger(field_config, trigger_name):
    """Clona configuración del campo y fuerza trigger de validación UI."""
    updated = dict(field_config)
    error_config = dict(field_config.get("error_config") or {})
    error_config["trigger"] = str(trigger_name or "").strip().lower() or "blur"
    updated["error_config"] = error_config
    return updated


# ── Palabras clave para clasificar botones CTA ───────────────────────────────
# «Siguiente» — botones de avance de step. Se normaliza el texto antes de comparar,
# por eso se usan formas sin acento (próximo → proximo, avançar → avancar).
_NEXT_KEYWORDS = (
    "siguiente",   # ES
    "seguinte",    # PT
    "continuar",   # ES/PT
    "continuacao", # PT
    "prosseguir",  # PT-BR
    "avancar",     # PT-BR (avançar → avancar tras strip de acentos)
    "seguir",      # ES
    "proximo",     # PT (próximo → proximo)
    "next",        # EN
    "continue",    # EN
)
# «Enviar/Submit» — botones de envío final. NUNCA deben usarse para avanzar steps.
_SUBMIT_KEYWORDS = (
    "enviar",      # ES/PT
    "envio",       # ES/PT
    "submit",      # EN
    "send",        # EN
    "envoy",       # FR informal
    "inviare",     # IT
    "soumettre",   # FR
    "verzenden",   # NL
    "submitbutton",  # clase CSS común
    "btn-visid-submit",  # Bolivia GM
)


def _is_next_button_element(element):
    try:
        if not element.is_displayed():
            return False

        text_parts = [
            element.text or "",
            element.get_attribute("data-dtm") or "",
            element.get_attribute("class") or "",
            element.get_attribute("aria-label") or "",
            element.get_attribute("title") or "",
        ]
        normalized = " ".join(_normalize_text(part) for part in text_parts if part)
        if any(kw in normalized for kw in _SUBMIT_KEYWORDS):
            return False
        return any(kw in normalized for kw in _NEXT_KEYWORDS)
    except Exception:
        return False


def _find_next_button(driver):
    selectors = [
        ".button.next.pulsate.stat-button-link",
        "button[class*='next']",
        "button[data-dtm*='next']",
        ".next-button",
        "button[type='button']",  # ultimo recurso — filtrado por _is_next_button_element
    ]

    for selector in selectors:
        try:
            buttons = driver.find_elements(By.CSS_SELECTOR, selector)
            for button in buttons:
                if _is_next_button_element(button):
                    return button
        except Exception:
            continue
    return None


def _click_next_button(driver):
    try:
        button = _find_next_button(driver)
        if not button:
            return False

        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", button)
        time.sleep(0.5)  # igual que Tab 1 (_click_next_button en base_form_filler)
        driver.execute_script("arguments[0].click();", button)
        return True
    except Exception:
        try:
            button = _find_next_button(driver)
            if button:
                button.click()
                return True
        except Exception:
            pass
    return False


_SUBMIT_KEYWORDS = (
    "enviar",      # ES/PT
    "envio",       # ES/PT
    "submit",      # EN
    "send",        # EN
    "envoy",       # FR informal
    "inviare",     # IT
    "soumettre",   # FR
    "verzenden",   # NL
    "submitbutton",  # clase CSS común
    "btn-visid-submit",  # Bolivia GM
)


def _find_cta_button(driver):
    """Busca el botón CTA del step: primero 'siguiente', luego 'enviar/submit'.

    Retorna (element, is_submit) donde is_submit=True indica que el botón es de envío
    final y NO debe ser clickeado para avanzar al siguiente step.
    """
    # Intentar primero con botón siguiente (no es submit)
    next_btn = _find_next_button(driver)
    if next_btn:
        return next_btn, False

    # Buscar botón de envío/submit como fallback
    submit_selectors = [
        "button.submit-button.stat-button-link",   # GM principal
        "button.btn-visid-submit.stat-button-link", # Bolivia GM
        "button[type='submit']",
        "input[type='submit']",
        ".button.submit",
        "button",
    ]
    for selector in submit_selectors:
        try:
            buttons = driver.find_elements(By.CSS_SELECTOR, selector)
            for button in buttons:
                try:
                    if not button.is_displayed():
                        continue
                    text_parts = [
                        button.text or "",
                        button.get_attribute("class") or "",
                        button.get_attribute("aria-label") or "",
                        button.get_attribute("value") or "",
                        button.get_attribute("data-dtm") or "",
                    ]
                    normalized = " ".join(p.strip().lower() for p in text_parts if p)
                    if any(kw in normalized for kw in _SUBMIT_KEYWORDS):
                        return button, True
                except Exception:
                    continue
        except Exception:
            continue
    return None, False


def _visible_ids_signature(driver, normalized_fields):
    text_ids = []
    dropdown_ids = []
    for field_config in normalized_fields:
        element_id = field_config["element_id"]
        if field_config.get("dropdown"):
            if _get_element_by_id_any_tag(driver, element_id) is not None:
                dropdown_ids.append(element_id)
            continue

        if _get_visible_element_by_id(driver, element_id) is not None:
            text_ids.append(element_id)

    # Incluir dropdowns visibles del DOM aunque no estén en reglas,
    # para detectar cambios de step cuando aparecen campos no mapeados.
    try:
        mapped = set(dropdown_ids)
        for el in driver.find_elements(By.CSS_SELECTOR, "select[id]"):
            if not el.is_displayed():
                continue
            el_id = (el.get_attribute("id") or "").strip()
            if el_id and el_id not in mapped:
                dropdown_ids.append(el_id)
    except Exception:
        pass

    return (tuple(sorted(text_ids)), tuple(sorted(dropdown_ids)))


def _wait_for_step_change(driver, normalized_fields, before_signature, timeout=3.0, poll_interval=0.2):
    """Espera cambio de firma visible del step para evitar cortes por carga tardia."""
    deadline = time.time() + max(0.2, float(timeout))
    while time.time() < deadline:
        after_signature = _visible_ids_signature(driver, normalized_fields)
        if after_signature != before_signature:
            return True, after_signature
        time.sleep(max(0.05, float(poll_interval)))

    return False, _visible_ids_signature(driver, normalized_fields)


def _check_inputmode(element):
    """Devuelve el valor real del atributo inputmode del elemento, o string vacío si no tiene."""
    try:
        value = element.get_attribute("inputmode") or ""
        return value.strip().lower()
    except Exception:
        return ""


def _validate_teclado_mobile(element, expects_numeric):
    """
    Valida que el inputmode coincida con la expectativa.
    - expects_numeric=True  → debe tener inputmode="numeric"
    - expects_numeric=False → NO debe tener inputmode="numeric"
    Devuelve (ok: bool, resultado: str, inputmode_real: str)
    """
    real_inputmode = _check_inputmode(element)
    tiene_numeric = real_inputmode == "numeric"

    if expects_numeric and not tiene_numeric:
        return False, "ERROR", real_inputmode or "(ninguno)"
    if not expects_numeric and tiene_numeric:
        return False, "ERROR", real_inputmode
    return True, "OK", real_inputmode or "(ninguno)"


def _normalize_field_mapping(field_mapping):
    if not isinstance(field_mapping, dict):
        raise ValueError("field_mapping debe ser un diccionario")

    normalized = []
    for field_name, config in field_mapping.items():
        if not isinstance(config, dict):
            raise ValueError(f"La configuracion del campo '{field_name}' debe ser un dict")

        element_id = (config.get("element_id") or config.get("id") or field_name or "").strip()
        regex_full = (config.get("regex_full") or "").strip()
        regex_char = (config.get("regex_char") or "").strip()
        test_text = str(config.get("test_text") or config.get("texto_prueba") or "")
        descripcion = str(config.get("descripcion") or field_name).strip() or str(field_name)
        teclado_mobile = bool(config.get("teclado_mobile", False))
        rules = dict(config.get("rules") or {})
        error_messages = dict(config.get("error_messages") or {})
        error_message_patterns = list(config.get("error_message_patterns") or [])
        error_config = dict(config.get("error_config") or {})
        error_priority = list(config.get("error_priority") or ["required", "invalid_chars", "min_length", "max_length"])

        is_dropdown = bool(config.get("dropdown", False))
        dropdown_error_message = str(config.get("dropdown_error_message") or "").strip()
        dependencies = _normalize_rule_dependencies(config.get("dependencies"))

        if not element_id:
            raise ValueError(f"El campo '{field_name}' no tiene element_id")
        if not is_dropdown and (not regex_full or not regex_char):
            raise ValueError(f"El campo '{field_name}' requiere regex_full y regex_char")

        normalized.append(
            {
                "field_name": str(field_name),
                "descripcion": descripcion,
                "element_id": element_id,
                "regex_full": regex_full,
                "regex_char": regex_char,
                "test_text": test_text,
                "teclado_mobile": teclado_mobile,
                "rules": rules,
                "error_messages": error_messages,
                "error_message_patterns": error_message_patterns,
                "error_config": error_config,
                "error_priority": error_priority,
                "dropdown": is_dropdown,
                "dropdown_error_message": dropdown_error_message,
                "dependencies": dependencies,
            }
        )

    return normalized


def run_field_validations(
    field_mapping,
    url=None,
    landing_url=None,
    expected_form_url=None,
    browser="chrome",
    viewport="fullscreen",
    headless=False,
    page_ready_timeout=30,
    ui_error_timeout=0.3,
    post_load_wait=0.5,
    max_steps=6,
    driver=None,
):
    """Ejecuta la validacion de varios campos de texto por ID en una pagina.

    Si se pasa ``driver``, se reutiliza el navegador existente y NO se cierra
    al finalizar (la responsabilidad del cierre queda en el llamador).
    """
    target_landing_url = _sanitize_url(landing_url or url)
    target_form_url = _sanitize_url(expected_form_url)

    if not target_landing_url:
        raise ValueError("La URL de landing es obligatoria")
    if not target_form_url:
        raise ValueError("La URL del form es obligatoria")

    normalized_fields = _normalize_field_mapping(field_mapping)
    _driver_owned = driver is None  # True si lo creamos aquí, False si es externo
    started_at = datetime.now()
    switched_to_form_iframe = False

    try:
        if _driver_owned:
            driver = BrowserManager.create_browser(
                browser_type=browser,
                viewport=viewport,
                headless=headless,
            )
        driver.get(target_landing_url)
        _wait_document_ready(driver, page_ready_timeout)
        if post_load_wait > 0:
            time.sleep(post_load_wait)

        _pre_scroll_for_dynamic_content(driver)
        iframe = _find_form_iframe(driver, target_form_url)
        if iframe is None:
            raise ValueError(f"No se encontró iframe para el form URL: {target_form_url}")

        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", iframe)
        time.sleep(0.3)
        driver.switch_to.frame(iframe)
        switched_to_form_iframe = True

        try:
            _wait_document_ready(driver, page_ready_timeout)
        except TimeoutException:
            LOGGER.warning("El iframe no reportó readyState=complete dentro del timeout")

        detail_rows = []
        field_summaries = []
        error_rows = []
        unmapped_rows = []
        tracked_unmapped_ids = set()
        configured_ids = {fc["element_id"] for fc in normalized_fields}

        steps_processed = 0
        for step in range(1, max_steps + 1):
            print(f"\n{'='*60}")
            print(f"[STEP {step}] inicio")

            # ── Inventario de campos visibles en este step ────────────────────────────
            visible_fields = _collect_visible_fields(driver, normalized_fields)
            visible_dropdowns = _collect_visible_dropdowns(driver, normalized_fields)
            fallback_dropdowns = _collect_fallback_dom_dropdowns(driver, normalized_fields)
            if fallback_dropdowns:
                visible_dropdowns.extend(fallback_dropdowns)

            print(f"[STEP {step}] campos texto visibles   : {[fc['element_id'] for fc, _ in visible_fields]}")
            print(f"[STEP {step}] dropdowns visibles      : {[fc['element_id'] for fc, _ in visible_dropdowns]}")
            if fallback_dropdowns:
                print(f"[STEP {step}] dropdowns fallback DOM : {[fc['element_id'] for fc, _ in fallback_dropdowns]}")

            # ── Todos los inputs/selects del DOM (incluyendo los no configurados) ─────
            try:
                visible_controls = [
                    el
                    for el in driver.find_elements(By.CSS_SELECTOR, "input[id], select[id], textarea[id]")
                    if el.is_displayed()
                ]
                all_dom_ids = [el.get_attribute("id") or "(sin-id)" for el in visible_controls]
                print(f"[STEP {step}] todos los inputs/selects visibles en DOM: {all_dom_ids}")

                for el in visible_controls:
                    element_id = (el.get_attribute("id") or "").strip()
                    if not element_id:
                        continue
                    if element_id in configured_ids:
                        continue
                    unique_key = (target_form_url, element_id)
                    if unique_key in tracked_unmapped_ids:
                        continue

                    tracked_unmapped_ids.add(unique_key)
                    unmapped_rows.append(
                        {
                            "timestamp": started_at.strftime("%Y-%m-%d %H:%M:%S"),
                            "landing_url": target_landing_url,
                            "form_url": target_form_url,
                            "url": target_landing_url,
                            "browser": browser,
                            "viewport": viewport,
                            "step": step,
                            "element_id": element_id,
                            "tag": (el.tag_name or "").strip().lower(),
                            "name": (el.get_attribute("name") or "").strip(),
                            "input_type": (el.get_attribute("type") or "").strip().lower(),
                            "placeholder": (el.get_attribute("placeholder") or "").strip(),
                            "motivo": "ID visible en formulario sin regla en field_validation_rules.",
                        }
                    )
                    print(f"[STEP {step}][UNMAPPED] id='{element_id}' sin regla en mapping")
            except Exception as _e:
                print(f"[STEP {step}] error inspeccionando DOM: {_e}")

            if not visible_fields and not visible_dropdowns:
                print(f"[STEP {step}] sin campos visibles — intentando avanzar")
            else:
                steps_processed += 1

            # ── 1. Detectar CTA ───────────────────────────────────────────────────────
            cta_button, is_submit_cta = _find_cta_button(driver)
            if cta_button:
                try:
                    cta_text = (cta_button.text or "").strip()
                    cta_class = (cta_button.get_attribute("class") or "").strip()
                    print(f"[STEP {step}] CTA encontrado: texto='{cta_text}' class='{cta_class}' is_submit={is_submit_cta}")
                except Exception:
                    print(f"[STEP {step}] CTA encontrado (sin poder leer atributos) is_submit={is_submit_cta}")
            else:
                print(f"[STEP {step}] CTA NO encontrado (ni siguiente ni enviar)")

            # ── 2. Click CTA inicial para disparar errores base de campos vacíos ─────
            if cta_button and (visible_fields or visible_dropdowns):
                print(f"[STEP {step}] clickeando CTA para disparar errores de validación...")
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", cta_button)
                time.sleep(0.2)
                driver.execute_script("arguments[0].click();", cta_button)
                time.sleep(0.8)
                print(f"[STEP {step}] CTA clickeado OK")
            else:
                print(f"[STEP {step}] salteando click CTA (cta={cta_button is not None}, campos={bool(visible_fields or visible_dropdowns)})")

            # ── 3. Leer y registrar errores de dropdowns en orden de dependencia ─────
            ordered_dropdowns = _sort_dropdowns_by_dependency(visible_dropdowns, FIELD_DEPENDENCIES)
            processed_dd_ids: set = set()
            print(f"[STEP {step}] ── revisando errores de {len(ordered_dropdowns)} dropdown(s) con dependencias")
            for dd_config, _dd_element in ordered_dropdowns:
                dd_id = dd_config["element_id"]
                dd_descripcion = dd_config.get("descripcion") or dd_config["field_name"]
                expected_msg = (dd_config.get("dropdown_error_message") or "").strip()

                context_ready = _ensure_dropdown_parent_context(driver, dd_id, FIELD_DEPENDENCIES)
                if not context_ready:
                    print(f"[STEP {step}][DD ERROR] contexto incompleto para '{dd_id}', se valida igual")

                dd_element = _get_element_by_id_any_tag(driver, dd_id)
                if dd_element is None:
                    print(f"[STEP {step}][DD ERROR] no se encontró elemento visible para '{dd_id}', se omite")
                    continue

                cta_for_dd, _ = _find_cta_button(driver)
                if cta_for_dd:
                    try:
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", cta_for_dd)
                        time.sleep(0.1)
                        driver.execute_script("arguments[0].click();", cta_for_dd)
                        time.sleep(0.5)
                    except Exception:
                        pass

                real_msg = read_error_near_element(driver, dd_element).strip()
                if not real_msg:
                    time.sleep(0.3)
                    real_msg = read_error_near_element(driver, dd_element).strip()

                if expected_msg:
                    dd_resultado = "OK" if real_msg == expected_msg else "ERROR"
                else:
                    dd_resultado = "OK" if real_msg else "ERROR"

                LOGGER.info(
                    "Dropdown step %s '%s' (%s): esperado='%s' real='%s' → %s",
                    step, dd_descripcion, dd_id, expected_msg, real_msg, dd_resultado,
                )
                print(f"[STEP {step}][DD ERROR] id='{dd_id}' esperado='{expected_msg}' real='{real_msg}' resultado={dd_resultado}")
                error_rows.append({
                    "timestamp": started_at.strftime("%Y-%m-%d %H:%M:%S"),
                    "landing_url": target_landing_url,
                    "form_url": target_form_url,
                    "url": target_landing_url,
                    "browser": browser,
                    "viewport": viewport,
                    "step": step,
                    "element_id": dd_id,
                    "descripcion": dd_descripcion,
                    "field_name": dd_config["field_name"],
                    "campo": dd_config["field_name"],
                    "input": "(sin selección)",
                    "regla_principal": "dropdown_required",
                    "regex_disparada": "",
                    "error_esperado": expected_msg,
                    "error_real": real_msg,
                    "resultado": dd_resultado,
                    "detalle": "" if dd_resultado == "OK" else f"Se esperaba: '{expected_msg}' | Se obtuvo: '{real_msg}'",
                    "trigger": "cta_click",
                    "error_selector": "",
                })
                field_summaries.append({
                    "step": step,
                    "landing_url": target_landing_url,
                    "form_url": target_form_url,
                    "element_id": dd_id,
                    "descripcion": dd_descripcion,
                    "field_name": dd_config["field_name"],
                    "regex_ok": None,
                    "final_value": "",
                    "errores": 0,
                    "errores_ui": 1 if dd_resultado == "ERROR" else 0,
                    "tests_ui": 1,
                    "teclado_mobile": "",
                    "teclado_mobile_ok": True,
                })

                # Dejar este dropdown resuelto para habilitar dependencias hijas.
                dd_element_after = _get_element_by_id_any_tag(driver, dd_id)
                if dd_element_after is not None:
                    _select_first_available_option(driver, dd_element_after)
                    child_id = FIELD_DEPENDENCIES.get(dd_id)
                    if child_id:
                        _wait_child_dropdown_options(
                            driver,
                            child_id,
                            timeout=DEPENDENCY_DROPDOWN_TIMEOUT,
                            poll_interval=DEPENDENCY_DROPDOWN_POLL_INTERVAL,
                        )
                processed_dd_ids.add(dd_id)

            # ── 4. Completar dropdowns para desbloquear el step ───────────────────────
            def _get_pending_dropdowns():
                current = _collect_visible_dropdowns(driver, normalized_fields)
                current.extend(_collect_fallback_dom_dropdowns(driver, normalized_fields))
                # Deduplicar por element_id preservando el primero encontrado.
                unique = {}
                for fc, el in current:
                    unique.setdefault(fc["element_id"], (fc, el))
                current = list(unique.values())
                pending = [(fc, el) for fc, el in current if fc["element_id"] not in processed_dd_ids]
                return _sort_dropdowns_by_dependency(pending, FIELD_DEPENDENCIES)

            dd_fill_queue = _get_pending_dropdowns()
            print(f"[STEP {step}] ── completando {len(dd_fill_queue)} dropdown(s): {[fc['element_id'] for fc,_ in dd_fill_queue]}")
            while dd_fill_queue:
                dd_config, dd_el = dd_fill_queue.pop(0)
                dd_id = dd_config["element_id"]
                if dd_id in processed_dd_ids:
                    print(f"[DD QUEUE] '{dd_id}' ya procesado, se omite")
                    continue

                print(f"[DD QUEUE] step {step}: procesando '{dd_id}' ({dd_config.get('descripcion')})")
                try:
                    el_tag = dd_el.tag_name
                    el_displayed = dd_el.is_displayed()
                    el_enabled = dd_el.is_enabled()
                    el_disabled_attr = dd_el.get_attribute("disabled")
                    print(f"[DD QUEUE]   elemento: tag={el_tag} displayed={el_displayed} enabled={el_enabled} disabled_attr={el_disabled_attr}")
                except Exception as _e:
                    print(f"[DD QUEUE]   no se pudo inspeccionar elemento: {_e}")

                selected_ok = _select_first_available_option(driver, dd_el)
                processed_dd_ids.add(dd_id)

                if not selected_ok:
                    print(f"[STEP {step}][DD] FALLO al seleccionar '{dd_id}'")
                    LOGGER.warning(
                        "No se pudo completar dropdown step %s: '%s' (%s)",
                        step,
                        dd_config.get("descripcion"),
                        dd_id,
                    )
                    time.sleep(0.5)
                    dd_fill_queue = _get_pending_dropdowns()
                    print(f"[STEP {step}][DD] reintentando — cola: {[fc['element_id'] for fc,_ in dd_fill_queue]}")
                    continue

                child_id = FIELD_DEPENDENCIES.get(dd_id)
                configured_ids = {fc["element_id"] for fc in normalized_fields}
                if child_id and child_id in configured_ids:
                    print(f"[STEP {step}][DD] '{dd_id}' tiene hijo='{child_id}', esperando carga...")
                    child_loaded = _wait_child_dropdown_options(
                        driver,
                        child_id,
                        timeout=DEPENDENCY_DROPDOWN_TIMEOUT,
                        poll_interval=DEPENDENCY_DROPDOWN_POLL_INTERVAL,
                    )
                    print(f"[STEP {step}][DD] hijo '{child_id}' cargado={child_loaded}")
                    LOGGER.info(
                        "Dependencia dropdown step %s: padre=%s hijo=%s cargado=%s",
                        step, dd_id, child_id, child_loaded,
                    )
                elif child_id:
                    print(f"[STEP {step}][DD] '{dd_id}' tiene hijo='{child_id}' fuera de reglas, omitido")
                else:
                    print(f"[STEP {step}][DD] '{dd_id}' sin hijo en FIELD_DEPENDENCIES")

                print(f"[STEP {step}][DD] '{dd_id}' completado OK")
                LOGGER.info("Dropdown completado step %s: '%s' (%s)", step, dd_config.get("descripcion"), dd_id)
                time.sleep(0.5)
                dd_fill_queue = _get_pending_dropdowns()
                print(f"[STEP {step}][DD] cola restante: {[fc['element_id'] for fc,_ in dd_fill_queue]}")

            # ── 5. Validar campos de texto ────────────────────────────────────────────
            print(f"[STEP {step}] ── validando {len(visible_fields)} campo(s) de texto")
            for field_config, _input_element in visible_fields:
                field_name = field_config["field_name"]
                descripcion = field_config.get("descripcion") or field_name
                element_id = field_config["element_id"]
                expects_numeric = bool(field_config.get("teclado_mobile", False))
                LOGGER.info("Validando step %s: '%s' (%s)", step, descripcion, element_id)

                dependencies_ok, dependencies_error = _apply_field_dependencies(driver, field_config, FIELD_DEPENDENCIES)
                if not dependencies_ok:
                    print(f"[STEP {step}][FIELD] dependencias fallidas para '{element_id}': {dependencies_error}")
                    error_rows.append(
                        {
                            "timestamp": started_at.strftime("%Y-%m-%d %H:%M:%S"),
                            "landing_url": target_landing_url,
                            "form_url": target_form_url,
                            "url": target_landing_url,
                            "browser": browser,
                            "viewport": viewport,
                            "step": step,
                            "element_id": element_id,
                            "descripcion": descripcion,
                            "field_name": field_name,
                            "campo": field_name,
                            "input": "",
                            "regla_principal": "dependencies",
                            "regex_disparada": "",
                            "error_esperado": "Dependencias aplicadas correctamente",
                            "error_real": dependencies_error,
                            "resultado": "ERROR",
                            "detalle": dependencies_error,
                            "trigger": str((field_config.get("error_config") or {}).get("trigger") or "blur"),
                            "error_selector": str((field_config.get("error_config") or {}).get("selector") or ""),
                        }
                    )
                    field_summaries.append(
                        {
                            "step": step,
                            "landing_url": target_landing_url,
                            "form_url": target_form_url,
                            "element_id": element_id,
                            "descripcion": descripcion,
                            "field_name": field_name,
                            "regex_ok": False,
                            "final_value": "",
                            "errores": 1,
                            "errores_ui": 1,
                            "tests_ui": 0,
                            "teclado_mobile": "",
                            "teclado_mobile_ok": False,
                        }
                    )
                    continue

                input_element = _get_visible_element_by_id(driver, element_id)
                if input_element is None:
                    print(f"[STEP {step}][FIELD] no visible tras dependencias: '{element_id}'")
                    error_rows.append(
                        {
                            "timestamp": started_at.strftime("%Y-%m-%d %H:%M:%S"),
                            "landing_url": target_landing_url,
                            "form_url": target_form_url,
                            "url": target_landing_url,
                            "browser": browser,
                            "viewport": viewport,
                            "step": step,
                            "element_id": element_id,
                            "descripcion": descripcion,
                            "field_name": field_name,
                            "campo": field_name,
                            "input": "",
                            "regla_principal": "dependencies",
                            "regex_disparada": "",
                            "error_esperado": "Campo visible tras aplicar dependencias",
                            "error_real": f"Campo '{element_id}' no visible después de aplicar dependencias",
                            "resultado": "ERROR",
                            "detalle": "El campo dejó de estar visible en el contexto de dependencias actual.",
                            "trigger": str((field_config.get("error_config") or {}).get("trigger") or "blur"),
                            "error_selector": str((field_config.get("error_config") or {}).get("selector") or ""),
                        }
                    )
                    field_summaries.append(
                        {
                            "step": step,
                            "landing_url": target_landing_url,
                            "form_url": target_form_url,
                            "element_id": element_id,
                            "descripcion": descripcion,
                            "field_name": field_name,
                            "regex_ok": False,
                            "final_value": "",
                            "errores": 1,
                            "errores_ui": 1,
                            "tests_ui": 0,
                            "teclado_mobile": "",
                            "teclado_mobile_ok": False,
                        }
                    )
                    continue

                print(f"[STEP {step}][FIELD] validando '{element_id}' test_text='{field_config.get('test_text','')}' regex_char='{field_config.get('regex_char','')}' regex_full='{field_config.get('regex_full','')}'")
                teclado_ok, teclado_resultado, teclado_real = _validate_teclado_mobile(
                    input_element, expects_numeric
                )
                teclado_mobile_col = (
                    f"numeric ({teclado_resultado})"
                    if expects_numeric
                    else (f"no-numeric-INESPERADO ({teclado_resultado})" if not teclado_ok else "(no aplica)")
                )
                print(f"[STEP {step}][FIELD] teclado_mobile expects_numeric={expects_numeric} real='{teclado_real}' ok={teclado_ok}")

                ui_inputs = generar_inputs_test(field_config)
                print(f"[STEP {step}][FIELD] {len(ui_inputs)} caso(s) UI a probar: {[c.get('input','') for c in ui_inputs]}")
                field_ui_error_count = 0
                tested_ui_count = 0

                # Para campos con dependencias, primero forzamos el caso vacío con CTA
                # para capturar el mensaje required antes de ejecutar otras validaciones.
                has_dependencies = bool(field_config.get("dependencies"))
                expects_empty_error = _expects_error_for_empty_input(field_config)
                forced_empty_idx = next(
                    (idx for idx, case in enumerate(ui_inputs) if str(case.get("input", "")) == ""),
                    None,
                )
                if has_dependencies and expects_empty_error:
                    if forced_empty_idx is not None:
                        forced_case = ui_inputs.pop(forced_empty_idx)
                    else:
                        forced_case = {"tipo": "required", "input": ""}
                    forced_config = _build_field_config_with_trigger(field_config, "next")
                    print(f"[STEP {step}][FIELD] dependencia activa → forzando caso vacío vía CTA para '{element_id}'")
                    forced_result = validar_error_ui(
                        campo=input_element,
                        config=forced_config,
                        driver=driver,
                        input_test=forced_case.get("input", ""),
                        timeout=max(0.3, float(ui_error_timeout)),
                    )
                    print(
                        f"[STEP {step}][FIELD] UI(required+CTA) resultado={forced_result.get('resultado')} "
                        f"esperado='{forced_result.get('error_esperado','')}' real='{forced_result.get('error_real','')}'"
                    )
                    forced_row = {
                        "timestamp": started_at.strftime("%Y-%m-%d %H:%M:%S"),
                        "landing_url": target_landing_url,
                        "form_url": target_form_url,
                        "url": target_landing_url,
                        "browser": browser,
                        "viewport": viewport,
                        "step": step,
                        "element_id": element_id,
                        "descripcion": descripcion,
                        "field_name": field_name,
                        "campo": field_name,
                        "input": forced_result.get("input", ""),
                        "regla_principal": forced_result.get("regla_principal", ""),
                        "regex_disparada": forced_result.get("regex_disparada", ""),
                        "error_esperado": forced_result.get("error_esperado", ""),
                        "error_real": forced_result.get("error_real", ""),
                        "resultado": forced_result.get("resultado", "ERROR"),
                        "detalle": forced_result.get("detalle", ""),
                        "trigger": forced_result.get("trigger_usado", "next"),
                        "trigger_inicial": forced_result.get("trigger_inicial", "next"),
                        "fallback_cta": bool(forced_result.get("fallback_cta", False)),
                        "error_selector": str((field_config.get("error_config") or {}).get("selector") or ""),
                    }
                    if forced_row["resultado"] == "ERROR":
                        field_ui_error_count += 1
                    error_rows.append(forced_row)
                    tested_ui_count += 1

                validation_result = validate_text_field(
                    input_element=input_element,
                    test_text=field_config["test_text"],
                    regex_full=field_config["regex_full"],
                    regex_char=field_config["regex_char"],
                )

                print(f"[STEP {step}][FIELD] validate_text_field regex_ok={validation_result.get('regex_ok')} final_value='{validation_result.get('final_value')}'")
                for ui_case in ui_inputs:
                    print(f"[STEP {step}][FIELD] UI caso input='{ui_case.get('input','')}' tipo='{ui_case.get('tipo','')}'")
                    ui_result = validar_error_ui(
                        campo=input_element,
                        config=field_config,
                        driver=driver,
                        input_test=ui_case.get("input", ""),
                        timeout=max(0.3, float(ui_error_timeout)),
                    )

                    print(f"[STEP {step}][FIELD] UI resultado={ui_result.get('resultado')} esperado='{ui_result.get('error_esperado','')}' real='{ui_result.get('error_real','')}'")
                    ui_row = {
                        "timestamp": started_at.strftime("%Y-%m-%d %H:%M:%S"),
                        "landing_url": target_landing_url,
                        "form_url": target_form_url,
                        "url": target_landing_url,
                        "browser": browser,
                        "viewport": viewport,
                        "step": step,
                        "element_id": element_id,
                        "descripcion": descripcion,
                        "field_name": field_name,
                        "campo": field_name,
                        "input": ui_result.get("input", ""),
                        "regla_principal": ui_result.get("regla_principal", ""),
                        "regex_disparada": ui_result.get("regex_disparada", ""),
                        "error_esperado": ui_result.get("error_esperado", ""),
                        "error_real": ui_result.get("error_real", ""),
                        "resultado": ui_result.get("resultado", "ERROR"),
                        "detalle": ui_result.get("detalle", ""),
                        "trigger": ui_result.get("trigger_usado", str((field_config.get("error_config") or {}).get("trigger") or "blur")),
                        "trigger_inicial": ui_result.get("trigger_inicial", str((field_config.get("error_config") or {}).get("trigger") or "blur")),
                        "fallback_cta": bool(ui_result.get("fallback_cta", False)),
                        "error_selector": str((field_config.get("error_config") or {}).get("selector") or ""),
                    }
                    if ui_row["resultado"] == "ERROR":
                        field_ui_error_count += 1
                    error_rows.append(ui_row)
                    tested_ui_count += 1

                rows = validation_result["rows"]
                field_summaries.append(
                    {
                        "step": step,
                        "landing_url": target_landing_url,
                        "form_url": target_form_url,
                        "element_id": element_id,
                        "descripcion": descripcion,
                        "field_name": field_name,
                        "regex_ok": validation_result["regex_ok"],
                        "final_value": validation_result["final_value"],
                        "errores": sum(1 for row in rows if row["resultado"] == "ERROR"),
                        "errores_ui": field_ui_error_count,
                        "tests_ui": tested_ui_count,
                        "teclado_mobile": teclado_mobile_col,
                        "teclado_mobile_ok": teclado_ok,
                    }
                )

                for index, row in enumerate(rows, start=1):
                    detail_rows.append(
                        {
                            "timestamp": started_at.strftime("%Y-%m-%d %H:%M:%S"),
                            "landing_url": target_landing_url,
                            "form_url": target_form_url,
                            "url": target_landing_url,
                            "browser": browser,
                            "viewport": viewport,
                            "step": step,
                            "element_id": element_id,
                            "descripcion": descripcion,
                            "field_name": field_name,
                            "char_index": index,
                            "char": row["char"],
                            "esperado": row["esperado"],
                            "real": row["real"],
                            "resultado": row["resultado"],
                            "valor_final": row["valor_final"],
                            "regex_ok": row["regex_ok"],
                            "regex_full": field_config["regex_full"],
                            "regex_char": field_config["regex_char"],
                            "test_text": field_config["test_text"],
                            "teclado_mobile": teclado_mobile_col,
                            "teclado_mobile_ok": teclado_ok,
                        }
                    )

            # ── 6. Re-llenar con test_text válido para desbloquear el botón siguiente ─
            # validar_error_ui deja los campos con valores inválidos (de prueba).
            # Si queda alguno inválido, el form bloquea el avance al siguiente step.
            print(f"[STEP {step}] ── re-llenando {len(visible_fields)} campo(s) con test_text válido")
            for _fc, _el in visible_fields:
                _test_text = (_fc.get("test_text") or "").strip()
                if not _test_text:
                    print(f"[STEP {step}][REFILL] '{_fc['element_id']}' sin test_text, saltando")
                    continue
                try:
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", _el)
                    _el.click()
                    _el.clear()
                    driver.execute_script("arguments[0].value='';", _el)
                    _el.send_keys(_test_text)
                    driver.execute_script(
                        "arguments[0].dispatchEvent(new Event('input',{bubbles:true}));"
                        "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
                        _el,
                    )
                    _el.send_keys(Keys.TAB)
                    # Verificar que el campo quedó con el valor esperado
                    try:
                        val_after = _el.get_attribute("value") or ""
                        print(f"[STEP {step}][REFILL] '{_fc['element_id']}' escrito='{_test_text}' valor_final='{val_after}'")
                    except Exception:
                        print(f"[STEP {step}][REFILL] '{_fc['element_id']}' escrito='{_test_text}' (no se pudo leer valor final)")
                except Exception as _refill_exc:
                    print(f"[STEP {step}][REFILL] ERROR en '{_fc['element_id']}': {_refill_exc}")

            # ── 7. Avanzar o finalizar ────────────────────────────────────────────────
            if is_submit_cta:
                print(f"[STEP {step}] CTA es ENVIAR/SUBMIT → break (próximo form en misma pestaña)")
                break

            before_signature = _visible_ids_signature(driver, normalized_fields)
            print(f"[STEP {step}] firma DOM antes de clic siguiente: {before_signature}")
            next_btn_found = _click_next_button(driver)
            print(f"[STEP {step}] _click_next_button → {next_btn_found}")
            if not next_btn_found:
                print(f"[STEP {step}] no se encontró botón siguiente → break")
                break

            time.sleep(1.5)
            changed, after_signature = _wait_for_step_change(
                driver,
                normalized_fields,
                before_signature,
                timeout=4.5,
                poll_interval=0.75,
            )
            print(f"[STEP {step}] cambio de step detectado={changed} firma_antes={before_signature} firma_después={after_signature}")
            if not changed:
                # Inspeccionar por qué no hubo cambio
                try:
                    err_els = driver.find_elements(By.CSS_SELECTOR, ".error, .invalid-feedback, [class*='error'], [class*='invalid']")
                    errs_vis = [(e.get_attribute("id") or "?", (e.text or "").strip()[:80]) for e in err_els if e.is_displayed() and (e.text or "").strip()]
                    print(f"[STEP {step}] STUCK — mensajes de error/validación visibles: {errs_vis}")
                    stuck_ids = [
                        el.get_attribute("id") or "?"
                        for el in driver.find_elements(By.CSS_SELECTOR, "input[id], select[id], textarea[id]")
                        if el.is_displayed()
                    ]
                    print(f"[STEP {step}] STUCK — campos visibles en DOM ahora: {stuck_ids}")
                except Exception as _stuck_e:
                    print(f"[STEP {step}] STUCK — error en debug extra: {_stuck_e}")
                print(f"[STEP {step}] DOM no cambió tras siguiente → break")
                break
            print(f"[STEP {step}] avanzó al siguiente step correctamente")

        total_errors = sum(1 for row in detail_rows if row["resultado"] == "ERROR")
        total_error_ui = sum(1 for row in error_rows if row["resultado"] == "ERROR")
        total_no_baseline_ui = sum(1 for row in error_rows if row.get("resultado") == "NO_BASELINE")
        total_unmapped_ids = len(unmapped_rows)
        return {
            "started_at": started_at,
            "url": target_landing_url,
            "landing_url": target_landing_url,
            "form_url": target_form_url,
            "browser": browser,
            "viewport": viewport,
            "headless": bool(headless),
            "rows": detail_rows,
            "fields": field_summaries,
            "error_rows": error_rows,
            "unmapped_ids": unmapped_rows,
            "summary": {
                "steps": steps_processed,
                "fields": len(field_summaries),
                "characters": len(detail_rows),
                "errors": total_errors,
                "ok": len(detail_rows) - total_errors,
                "ui_error_tests": len(error_rows),
                "ui_errors": total_error_ui,
                "ui_no_baseline": total_no_baseline_ui,
                "ui_ok": len(error_rows) - total_error_ui - total_no_baseline_ui,
                "ids_no_mapeados": total_unmapped_ids,
                "regex_ok": all(field["regex_ok"] for field in field_summaries) if field_summaries else True,
            },
        }
    finally:
        if driver is not None and switched_to_form_iframe:
            try:
                driver.switch_to.default_content()
            except Exception:
                pass
        if _driver_owned and driver is not None:
            driver.quit()