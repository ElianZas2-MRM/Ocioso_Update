import logging
import re

from selenium.webdriver.common.keys import Keys


LOGGER = logging.getLogger(__name__)


def _compile_regex(pattern, field_name):
    try:
        return re.compile(pattern)
    except re.error as exc:
        raise ValueError(f"Regex invalido para {field_name}: {pattern}") from exc


def _matches_char(pattern, value):
    return bool(pattern.fullmatch(value))


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


def _set_input_value_fast(input_element, value):
    driver = getattr(input_element, "parent", None)
    normalized_value = "" if value is None else str(value)
    if driver is None:
        return False

    try:
        applied_value = driver.execute_script(
            """
            const element = arguments[0];
            const nextValue = arguments[1] ?? '';
            const prototype = element.tagName === 'TEXTAREA'
                ? window.HTMLTextAreaElement.prototype
                : window.HTMLInputElement.prototype;
            const descriptor = Object.getOwnPropertyDescriptor(prototype, 'value');

            element.focus();
            if (descriptor && descriptor.set) {
                descriptor.set.call(element, nextValue);
            } else {
                element.value = nextValue;
            }

            element.dispatchEvent(new Event('input', { bubbles: true }));
            element.dispatchEvent(new Event('change', { bubbles: true }));
            return element.value || '';
            """,
            input_element,
            normalized_value,
        )
        return str(applied_value or "") == normalized_value
    except Exception:
        return False


def validate_text_field(input_element, test_text, regex_full, regex_char):
    """Valida un input de texto con verificacion final y caracter por caracter."""
    full_pattern = _compile_regex(regex_full, "regex_full")
    char_pattern = _compile_regex(str(regex_char or "").strip(), "regex_char")

    def _normalize_value(value):
        raw_value = str(value or "")
        return "".join(char for char in raw_value if char_pattern.fullmatch(char))

    rows = []

    for char in test_text:
        _clear_input_value(input_element)
        expected = "ACEPTADO" if _matches_char(char_pattern, char) else "RECHAZADO"

        input_element.send_keys(char)
        current_value = input_element.get_attribute("value") or ""
        normalized_current = _normalize_value(current_value)

        if not current_value:
            real = "RECHAZADO"
        elif current_value == char or current_value.endswith(char):
            real = "ACEPTADO"
        elif normalized_current == char:
            real = "ACEPTADO"
        else:
            real = "MODIFICADO"

        if real == "MODIFICADO":
            result = "ERROR"
        elif expected == "ACEPTADO" and real == "ACEPTADO":
            result = "OK"
        elif expected == "RECHAZADO" and real == "RECHAZADO":
            result = "OK"
        else:
            result = "ERROR"

        rows.append(
            {
                "char": char,
                "esperado": expected,
                "real": real,
                "resultado": result,
                "valor_final": current_value,
            }
        )

    _clear_input_value(input_element)
    used_fast_path = _set_input_value_fast(input_element, test_text)
    if not used_fast_path:
        input_element.send_keys(test_text)
    final_value = input_element.get_attribute("value") or ""
    regex_ok = bool(full_pattern.fullmatch(final_value))

    for row in rows:
        row["regex_ok"] = regex_ok

    LOGGER.debug("Validacion completada. valor_final=%s regex_ok=%s", final_value, regex_ok)
    return {
        "final_value": final_value,
        "regex_ok": regex_ok,
        "rows": rows,
    }