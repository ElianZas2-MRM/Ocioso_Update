"""
error_message_validator.py — Detecta y verifica mensajes de error en formularios.
Busca elementos de error en el DOM después de ingresar un valor inválido y compara
el texto del error con el esperado según las reglas de validación configuradas.
"""
import re
import time
import unicodedata

from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


DEFAULT_ERROR_PRIORITY = ["required", "invalid_chars", "min_length", "max_length"]


def _normalize_error_message_patterns(config):
    raw_patterns = config.get("error_message_patterns") or []
    if not isinstance(raw_patterns, list):
        return []

    normalized = []
    for raw_rule in raw_patterns:
        if not isinstance(raw_rule, dict):
            continue
        regex_text = str(raw_rule.get("regex") or "").strip()
        message_text = str(raw_rule.get("mensaje") or raw_rule.get("message") or "").strip()
        if not regex_text or not message_text:
            continue
        try:
            compiled = re.compile(regex_text)
        except re.error:
            continue
        normalized.append(
            {
                "regex": regex_text,
                "mensaje": message_text,
                "compiled": compiled,
            }
        )

    return normalized


def _resolve_error_message_from_patterns(input_value, config):
    patterns = _normalize_error_message_patterns(config)
    if not patterns:
        return "", ""

    for pattern_rule in patterns:
        compiled = pattern_rule["compiled"]
        if compiled.fullmatch(input_value) or compiled.match(input_value):
            return pattern_rule["mensaje"], pattern_rule["regex"]

    return "", ""


def _compile_char_regex(pattern):
    if not pattern:
        return None
    try:
        return re.compile(pattern)
    except re.error:
        return None


def _compile_full_regex(pattern):
    if not pattern:
        return None
    try:
        return re.compile(pattern)
    except re.error:
        return None


def _normalize_rules(config):
    rules = dict(config.get("rules") or {})
    return {
        "required": bool(rules.get("required", False)),
        "min_length": int(rules.get("min_length") or 0),
        "max_length": int(rules.get("max_length") or 0),
        "invalid_chars": bool(rules.get("invalid_chars", False)),
    }


def _normalize_priority(config):
    raw_priority = config.get("error_priority") or DEFAULT_ERROR_PRIORITY
    if not isinstance(raw_priority, list):
        return list(DEFAULT_ERROR_PRIORITY)
    cleaned = []
    for item in raw_priority:
        key = str(item or "").strip()
        if key and key not in cleaned:
            cleaned.append(key)
    return cleaned or list(DEFAULT_ERROR_PRIORITY)


def evaluar_reglas(valor, config):
    value = "" if valor is None else str(valor)
    rules = _normalize_rules(config)
    errors = []

    if rules["required"] and not value.strip():
        errors.append("required")

    if rules["min_length"] > 0 and len(value) < rules["min_length"]:
        errors.append("min_length")

    if rules["max_length"] > 0 and len(value) > rules["max_length"]:
        errors.append("max_length")

    if rules["invalid_chars"]:
        char_regex = _compile_char_regex(config.get("regex_char") or "")
        if char_regex is not None:
            has_invalid = any(not bool(char_regex.fullmatch(char)) for char in value)
            if has_invalid:
                errors.append("invalid_chars")

    return errors


def obtener_error_principal(errores, prioridad):
    if not errores:
        return None

    ordered_priority = prioridad or list(DEFAULT_ERROR_PRIORITY)
    for rule_name in ordered_priority:
        if rule_name in errores:
            return rule_name

    return errores[0]


def _normalize_by_char_regex(value, char_regex):
    raw_value = "" if value is None else str(value)
    if char_regex is None:
        return raw_value
    return "".join(char for char in raw_value if char_regex.fullmatch(char))


def _build_candidate_inputs(config):
    sample_text = str(config.get("test_text") or "")
    char_regex = _compile_char_regex(config.get("regex_char") or "")
    full_regex = _compile_full_regex(config.get("regex_full") or "")
    normalized_sample = _normalize_by_char_regex(sample_text, char_regex)

    base_candidates = [
        "",
        sample_text,
        normalized_sample,
        sample_text[:1],
        sample_text[:2],
        sample_text[:5],
        sample_text[:10],
        normalized_sample[:1],
        normalized_sample[:2],
        normalized_sample[:5],
        normalized_sample[:10],
        sample_text + "@",
        normalized_sample + "@",
        "@",
        "1",
        "11",
        "111111",
        "123456",
        "1234567",
        "0123456",
        "9999999999",
        "A",
        "AA",
        "ABC",
        "aaaaaa",
        "      ",
    ]

    candidates = []
    seen = set()
    for candidate in base_candidates:
        candidate_text = "" if candidate is None else str(candidate)
        if candidate_text in seen:
            continue
        seen.add(candidate_text)
        candidates.append(candidate_text)

    if full_regex is None:
        return candidates

    invalid_candidates = [candidate for candidate in candidates if not full_regex.fullmatch(candidate)]
    return invalid_candidates or candidates


def _append_unique_input(inputs, seen_inputs, tipo, input_value, extra=None):
    normalized_value = "" if input_value is None else str(input_value)
    unique_key = (tipo, normalized_value)
    if unique_key in seen_inputs:
        return
    seen_inputs.add(unique_key)
    item = {"tipo": tipo, "input": normalized_value}
    if extra:
        item.update(extra)
    inputs.append(item)


def generar_inputs_test(config):
    rules = _normalize_rules(config)
    inputs = []
    seen_inputs = set()

    if rules["required"]:
        _append_unique_input(inputs, seen_inputs, "required", "")

    if rules["min_length"] > 0:
        short_value = "a" if rules["min_length"] > 1 else ""
        _append_unique_input(inputs, seen_inputs, "min_length", short_value)

    if rules["invalid_chars"]:
        _append_unique_input(inputs, seen_inputs, "invalid_chars", "123@")

    if rules["max_length"] > 0:
        max_len = int(rules["max_length"])
        _append_unique_input(inputs, seen_inputs, "max_length", "a" * (max_len + 1))

    patterns = _normalize_error_message_patterns(config)
    if patterns:
        candidates = _build_candidate_inputs(config)
        for pattern_rule in patterns:
            for candidate in candidates:
                compiled = pattern_rule["compiled"]
                if compiled.fullmatch(candidate) or compiled.match(candidate):
                    _append_unique_input(
                        inputs,
                        seen_inputs,
                        "pattern",
                        candidate,
                        extra={
                            "regex": pattern_rule["regex"],
                            "mensaje": pattern_rule["mensaje"],
                        },
                    )
                    break

    # If no explicit rule-based tests are configured, at least test current sample text.
    sample_text = str(config.get("test_text") or "")
    if sample_text and not any(str(item.get("input", "")) == sample_text for item in inputs):
        _append_unique_input(inputs, seen_inputs, "sample_text", sample_text)

    return inputs


def _clear_input_value(input_element):
    input_element.click()
    input_element.clear()

    current_value = input_element.get_attribute("value") or ""
    if current_value:
        input_element.send_keys(Keys.CONTROL, "a")
        input_element.send_keys(Keys.DELETE)

    current_value = input_element.get_attribute("value") or ""
    if current_value:
        driver = getattr(input_element, "parent", None)
        if driver is not None:
            driver.execute_script("arguments[0].value = '';", input_element)


def _resolve_selector(selector):
    selector_text = str(selector or "").strip()
    if not selector_text:
        return None, None

    if selector_text.startswith("/") or selector_text.startswith("("):
        return By.XPATH, selector_text

    return By.CSS_SELECTOR, selector_text


def _normalize_text(value):
    """Minúsculas y sin acentos para comparación robusta de keywords."""
    text = str(value or "").strip().lower()
    text = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in text if not unicodedata.combining(ch))


# Keywords que identifican el boton de avance (siguiente / next).
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
# Keywords que identifican botones de envio — NUNCA deben clickearse.
_SUBMIT_KEYWORDS = (
    "enviar",      # ES/PT
    "envio",       # ES/PT
    "submit",      # EN
    "send",        # EN
    "envoy",       # FR informal
    "inviare",     # IT
    "soumettre",   # FR
    "verzenden",   # NL
    "submitbutton",  # clase CSS comun
    "btn-visid-submit",  # Bolivia GM
)


def _is_next_button_safe(element):
    """Devuelve True solo si el elemento es un boton de avance y NO es de envio."""
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
        normalized = " ".join(_normalize_text(p) for p in text_parts if p)
        if any(kw in normalized for kw in _SUBMIT_KEYWORDS):
            return False
        return any(kw in normalized for kw in _NEXT_KEYWORDS)
    except Exception:
        return False


def _trigger_next_button(driver):
    """Hace click en el boton de avance (siguiente/next). Nunca toca enviar/submit."""
    selectors = [
        ".button.next.pulsate.stat-button-link",
        "button[class*='next']",
        "button[data-dtm*='next']",
        ".next-button",
        "button[type='button']",
    ]
    for selector in selectors:
        try:
            for element in driver.find_elements(By.CSS_SELECTOR, selector):
                if _is_next_button_safe(element):
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
                    time.sleep(0.2)
                    driver.execute_script("arguments[0].click();", element)
                    return True
        except Exception:
            continue
    return False


def _trigger_submit_button(driver):
    """Hace click en boton enviar/submit visible cuando no hay boton siguiente."""
    submit_selectors = [
        "button.submit-button.stat-button-link",
        "button.btn-visid-submit.stat-button-link",
        "button[type='submit']",
        "input[type='submit']",
        ".button.submit",
        "button",
    ]

    for selector in submit_selectors:
        try:
            for element in driver.find_elements(By.CSS_SELECTOR, selector):
                if not element.is_displayed():
                    continue

                text_parts = [
                    element.text or "",
                    element.get_attribute("class") or "",
                    element.get_attribute("aria-label") or "",
                    element.get_attribute("value") or "",
                    element.get_attribute("data-dtm") or "",
                ]
                normalized = " ".join(_normalize_text(p) for p in text_parts if p)
                if any(keyword in normalized for keyword in _SUBMIT_KEYWORDS):
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
                    time.sleep(0.2)
                    driver.execute_script("arguments[0].click();", element)
                    return True
        except Exception:
            continue

    return False


def _trigger_cta_button(driver):
    """Intenta trigger de validación por CTA: primero siguiente y luego submit."""
    if _trigger_next_button(driver):
        return True
    return _trigger_submit_button(driver)


def _regex_matches_empty(regex_text):
    """Devuelve True cuando el patrón coincide con cadena vacía."""
    pattern = str(regex_text or "").strip()
    if not pattern:
        return False
    try:
        return bool(re.compile(pattern).fullmatch(""))
    except re.error:
        return False


def _trigger_validation(input_element, driver, trigger_mode):
    trigger = str(trigger_mode or "blur").strip().lower()
    if trigger == "blur":
        input_element.send_keys(Keys.TAB)
        return True
    if trigger == "input":
        return True
    if trigger in ("submit", "next", "cta"):
        return _trigger_cta_button(driver)
    return False


def _resolve_input_element(campo, config, driver):
    if hasattr(campo, "send_keys"):
        return campo

    field_id = str(config.get("element_id") or "").strip()
    if not field_id:
        return None

    try:
        return driver.find_element(By.ID, field_id)
    except Exception:
        return None


def read_error_near_element(driver, element):
    """Versión pública de _read_visible_error_text_near_input para uso externo."""
    return _read_visible_error_text_near_input(driver, element)


def _read_visible_error_text_near_input(driver, input_element):
        script = """
        const input = arguments[0];
        if (!input) return "";

        const isVisible = (el) => {
            if (!el) return false;
            const style = window.getComputedStyle(el);
            if (!style) return false;
            if (style.display === 'none' || style.visibility === 'hidden' || style.opacity === '0') return false;
            const rect = el.getBoundingClientRect();
            return rect.width > 0 && rect.height > 0;
        };

        const getText = (el) => (el && el.textContent ? el.textContent.trim() : "");
        const norm = (v) => String(v || '').trim().toLowerCase();

        // Priorizar el mensaje enlazado al campo para evitar capturar errores globales.
        const describedBy = (input.getAttribute('aria-describedby') || '')
            .split(/\s+/)
            .map((id) => id.trim())
            .filter(Boolean);
        for (const errId of describedBy) {
            const errEl = document.getElementById(errId);
            if (!isVisible(errEl)) continue;
            const errText = getText(errEl);
            if (errText) return errText;
        }

        const wrappers = [];
        const nearest = input.closest('.relative, .form-group, [data-field-wrap]');
        if (nearest) wrappers.push(nearest);
        if (input.parentElement) wrappers.push(input.parentElement);
        if (input.parentElement && input.parentElement.parentElement) wrappers.push(input.parentElement.parentElement);

        const specificTargets = [];
        const inputId = (input.id || '').trim();
        const inputName = (input.name || '').trim();
        const fieldTokens = [norm(inputId), norm(inputName)].filter(Boolean);

        const isAssociatedError = (el) => {
            if (!el) return false;
            if (!fieldTokens.length) return false;

            const attrs = [
                el.getAttribute('id'),
                el.getAttribute('for'),
                el.getAttribute('name'),
                el.getAttribute('data-error-for'),
                el.getAttribute('aria-describedby'),
                el.getAttribute('aria-labelledby'),
                el.getAttribute('data-dtm'),
                el.className,
            ];
            const haystack = norm(attrs.filter(Boolean).join(' '));
            if (!haystack) return false;
            return fieldTokens.some((token) => token && haystack.includes(token));
        };
        if (inputId) {
            specificTargets.push(`#${CSS.escape(inputId)}-error`);
            specificTargets.push(`#${CSS.escape(inputId)}_error`);
            specificTargets.push(`[id='${inputId}-error']`);
            specificTargets.push(`[id='${inputId}_error']`);
            specificTargets.push(`[data-error-for='${inputId}']`);
        }
        if (inputName) {
            specificTargets.push(`[id='${inputName}-error']`);
            specificTargets.push(`[id='${inputName}_error']`);
            specificTargets.push(`[data-error-for='${inputName}']`);
        }

        const selectors = [
            ".error.invalid-feedback",
            "[id$='-error'].error",
            "[id$='_error'].error",
            "[id$='-error']",
            "[id$='_error']",
            "[id*='error']",
            ".error",
            "div[x-show*='errors'][class*='text-red'] span",
            "div[x-show*='errors'] span[x-text]",
            "div[x-show*='errors'] span",
            "div[class*='text-red'] span",
            "span.text-red-600",
            ".invalid-feedback",
            "[data-error]",
        ];

        for (const specificSelector of specificTargets) {
            try {
                const nodes = document.querySelectorAll(specificSelector);
                for (const node of nodes) {
                    if (!isVisible(node)) continue;
                    const text = getText(node);
                    if (text) return text;
                }
            } catch (e) {
                // ignore malformed selectors
            }
        }

        // Priorizar errores asociados explícitamente al campo dentro de wrappers cercanos.
        for (const wrap of wrappers) {
            if (!wrap) continue;
            for (const selector of selectors) {
                const elements = wrap.querySelectorAll(selector);
                for (const el of elements) {
                    if (!isVisible(el)) continue;
                    if (!isAssociatedError(el)) continue;
                    const text = getText(el);
                    if (text) return text;
                }
            }
        }

        // Fallback local (sin recorrer el <form> completo para evitar corrimiento de mensajes).
        for (const wrap of wrappers) {
            if (!wrap) continue;
            for (const selector of selectors) {
                const elements = wrap.querySelectorAll(selector);
                for (const el of elements) {
                    if (!isVisible(el)) continue;
                    const text = getText(el);
                    if (text) return text;
                }
            }
        }

        return "";
        """

        try:
                return (driver.execute_script(script, input_element) or "").strip()
        except Exception:
                return ""


def validar_error_ui(campo, config, driver, input_test, timeout=4):
    input_value = "" if input_test is None else str(input_test)
    error_messages = dict(config.get("error_messages") or {})
    prioridad = _normalize_priority(config)
    selector_by, selector_value = _resolve_selector((config.get("error_config") or {}).get("selector"))
    trigger = (config.get("error_config") or {}).get("trigger") or "blur"
    trigger_inicial = str(trigger or "blur").strip().lower() or "blur"
    trigger_usado = trigger_inicial
    fallback_cta = False
    input_element = _resolve_input_element(campo, config, driver)

    if input_element is None:
        return {
            "input": input_value,
            "errores_detectados": [],
            "error_esperado": "",
            "error_real": "",
            "resultado": "ERROR",
            "detalle": "No se encontro el input del campo.",
        }

    errores = evaluar_reglas(input_value, config)
    regla_principal = obtener_error_principal(errores, prioridad)
    matched_message, matched_regex = _resolve_error_message_from_patterns(input_value, config)
    error_esperado = matched_message or (str(error_messages.get(regla_principal) or "") if regla_principal else "")

    try:
        _clear_input_value(input_element)
        if input_value:
            input_element.send_keys(input_value)

        trigger_ok = _trigger_validation(input_element, driver, trigger)
        if not trigger_ok:
            return {
                "input": input_value,
                "errores_detectados": errores,
                "regla_principal": regla_principal,
                "regex_disparada": matched_regex,
                "error_esperado": error_esperado,
                "error_real": "",
                "resultado": "ERROR",
                "detalle": "No se pudo disparar la validacion UI.",
                "trigger_inicial": trigger_inicial,
                "trigger_usado": trigger_usado,
                "fallback_cta": fallback_cta,
            }

        if trigger_inicial in ("submit", "next", "cta"):
            trigger_usado = "cta"

        error_real = ""

        def _read_error_once():
            if selector_by and selector_value:
                try:
                    nodes = driver.find_elements(selector_by, selector_value)
                    for node in nodes:
                        try:
                            if not node.is_displayed():
                                continue
                            text = (node.text or "").strip()
                            if text:
                                return text
                        except Exception:
                            continue
                except Exception:
                    return ""
                return ""
            return _read_visible_error_text_near_input(driver, input_element)

        expected_norm = (error_esperado or "").strip()
        wait_timeout = max(0.3, float(timeout))

        if expected_norm:
            # Si esperamos mensaje, sí conviene esperar hasta timeout completo.
            deadline = time.time() + wait_timeout
            while time.time() < deadline:
                error_real = _read_error_once().strip()
                if error_real:
                    break
                time.sleep(0.1)
        else:
            # Si NO esperamos mensaje, hacer sondeo corto para evitar pausas largas.
            quick_timeout = min(wait_timeout, 0.8)
            deadline = time.time() + quick_timeout
            while time.time() < deadline:
                error_real = _read_error_once().strip()
                if error_real:
                    break
                time.sleep(0.1)

        real_norm = error_real.strip()

        empty_case_requires_submit = (
            not input_value
            and bool(expected_norm)
            and (
                regla_principal == "required"
                or _regex_matches_empty(matched_regex)
            )
        )
        should_retry_with_cta = (
            empty_case_requires_submit
            and str(trigger or "").strip().lower() not in ("submit", "next", "cta")
            and (not real_norm or real_norm != expected_norm)
        )
        if should_retry_with_cta:
            cta_ok = _trigger_cta_button(driver)
            if cta_ok:
                fallback_cta = True
                trigger_usado = "cta_fallback"
                deadline = time.time() + max(0.8, wait_timeout)
                while time.time() < deadline:
                    error_real = _read_error_once().strip()
                    if error_real:
                        break
                    time.sleep(0.1)
                real_norm = error_real.strip()
        detalle = ""
        if not expected_norm and not real_norm:
            resultado = "OK"
        elif not expected_norm and real_norm:
            resultado = "NO_BASELINE"
            detalle = "No hay mensaje esperado configurado para este caso."
        elif expected_norm and not real_norm:
            resultado = "ERROR"
            detalle = "No se detecto mensaje UI esperado."
        elif expected_norm == real_norm:
            resultado = "OK"
        else:
            resultado = "ERROR"
            detalle = "Mensaje UI distinto al esperado."

        return {
            "input": input_value,
            "errores_detectados": errores,
            "regla_principal": regla_principal,
            "regex_disparada": matched_regex,
            "error_esperado": error_esperado,
            "error_real": error_real,
            "resultado": resultado,
            "detalle": detalle,
            "trigger_inicial": trigger_inicial,
            "trigger_usado": trigger_usado,
            "fallback_cta": fallback_cta,
        }
    except Exception as exc:
        return {
            "input": input_value,
            "errores_detectados": errores,
            "regla_principal": regla_principal,
            "regex_disparada": matched_regex,
            "error_esperado": error_esperado,
            "error_real": "",
            "resultado": "ERROR",
            "detalle": str(exc),
            "trigger_inicial": trigger_inicial,
            "trigger_usado": trigger_usado,
            "fallback_cta": fallback_cta,
        }
