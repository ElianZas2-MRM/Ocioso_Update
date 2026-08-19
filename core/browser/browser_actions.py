"""
browser_actions.py
-------------------
BrowserActions: single seam wrapping every raw Selenium call used by
BaseFormFiller. Step 0 of the Playwright-migration refactor — this class is
scaffolded and importable, but nothing in base_form_filler.py calls it yet
(other than storing the instance on self.actions).

No dependency on BaseFormFiller or config: the constructor takes only a
Selenium `driver`.
"""
import time
import unicodedata

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC


class BrowserActions:
    def __init__(self, driver):
        self.driver = driver

    # ------------------------------------------------------------------
    # find
    # ------------------------------------------------------------------
    def find(self, by, value):
        return self.driver.find_element(by, value)

    def find_all(self, by, value):
        return self.driver.find_elements(by, value)

    def exists(self, by, value):
        try:
            return bool(self.driver.find_elements(by, value))
        except Exception:
            return False

    # ------------------------------------------------------------------
    # wait
    # ------------------------------------------------------------------
    def wait_until(self, predicate, timeout=8, poll_frequency=0.2, ignored_exceptions=None):
        return WebDriverWait(self.driver, timeout, poll_frequency, ignored_exceptions).until(predicate)

    def wait_clickable(self, by, value, timeout=8):
        return WebDriverWait(self.driver, timeout).until(EC.element_to_be_clickable((by, value)))

    def wait_present(self, by, value, timeout=8):
        return WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((by, value)))

    # ------------------------------------------------------------------
    # interaction
    # ------------------------------------------------------------------
    def click(self, element, js_fallback=True):
        try:
            element.click()
        except Exception:
            if js_fallback:
                self.driver.execute_script("arguments[0].click();", element)
            else:
                raise

    def click_js(self, element):
        self.driver.execute_script("arguments[0].click();", element)

    def send_keys(self, element, text):
        element.send_keys(text)

    def set_value(self, element, value):
        """Setea el value con el setter nativo del prototipo (React/Angular no registran
        la asignación directa de el.value). Copiado verbatim de
        BaseFormFiller._set_input_value_js."""
        try:
            self.driver.execute_script(
                "var e=arguments[0], v=arguments[1];"
                "var proto = e instanceof window.HTMLTextAreaElement"
                "    ? window.HTMLTextAreaElement.prototype : window.HTMLInputElement.prototype;"
                "var d = Object.getOwnPropertyDescriptor(proto, 'value');"
                "if (d && d.set) { d.set.call(e, v); } else { e.value = v; }"
                "e.dispatchEvent(new Event('input',{bubbles:true}));"
                "e.dispatchEvent(new Event('change',{bubbles:true}));",
                element, value,
            )
        except Exception:
            pass

    def clear_hard(self, element, get_current=None):
        """Vacía un input y confirma que quedó vacío. Copiado de
        BaseFormFiller._hard_clear_input, con `get_current` opcional para
        permitir un lector de valor actual distinto de get_attribute('value')."""
        def _current():
            if get_current:
                try:
                    return get_current(element) or ""
                except Exception:
                    return ""
            try:
                return element.get_attribute("value") or ""
            except Exception:
                return ""

        for intento in range(3):
            if not _current():
                return True
            if intento == 0:
                try:
                    element.clear()
                except Exception:
                    pass
            else:
                self.set_value(element, "")
            time.sleep(0.05)
        return not _current()

    # ------------------------------------------------------------------
    # checkbox / radio
    # ------------------------------------------------------------------
    def set_checkbox(self, element, checked=True):
        """Fija el estado checked de un checkbox y dispara los eventos del form.
        Generalizado (parámetro `checked`) a partir del patrón dispatchEvent de
        BaseFormFiller._uncheck_checkbox."""
        try:
            self.driver.execute_script(
                """
                var el = arguments[0], checked = arguments[1];
                if (!el || el.checked === checked) return;
                el.checked = checked;
                el.dispatchEvent(new Event('click',  {bubbles:true}));
                el.dispatchEvent(new Event('change', {bubbles:true}));
                el.checked = checked;   // por si algún handler lo revirtió
                """,
                element, checked,
            )
        except Exception:
            pass

    def set_radio_checked(self, element):
        """Copiado verbatim de BaseFormFiller._set_radio_checked_via_js."""
        try:
            result = self.driver.execute_script(
                """
                    const rb = arguments[0];
                    if (rb.checked) {
                        return true;
                    }
                    rb.focus && rb.focus();
                    rb.checked = true;
                    rb.dispatchEvent(new Event('input', { bubbles: true }));
                    rb.dispatchEvent(new Event('change', { bubbles: true }));
                    return rb.checked === true;
                """,
                element,
            )
            if result:
                print("ℹ Radio marcado mediante JS directo")
            return bool(result)
        except Exception:
            return False

    def prepare_for_interaction(self, element):
        """Copiado verbatim de BaseFormFiller._prepare_radio_for_interaction."""
        try:
            self.driver.execute_script(
                """
                    const rb = arguments[0];
                    if (rb.hasAttribute('disabled')) {
                        rb.removeAttribute('disabled');
                    }
                    rb.disabled = false;
                    rb.setAttribute('aria-disabled', 'false');
                    if (rb.tabIndex === -1) {
                        rb.tabIndex = 0;
                    }
                """,
                element,
            )
        except Exception:
            pass

    # ------------------------------------------------------------------
    # queries
    # ------------------------------------------------------------------
    def get_attribute(self, element, name):
        return element.get_attribute(name)

    def is_displayed(self, element):
        return element.is_displayed()

    def is_enabled(self, element):
        return element.is_enabled()

    # ------------------------------------------------------------------
    # select
    # ------------------------------------------------------------------
    def select_by_value(self, select_element, value):
        Select(select_element).select_by_value(value)

    def select_by_index(self, select_element, index):
        Select(select_element).select_by_index(index)

    def get_select_options(self, select_element):
        return Select(select_element).options

    @staticmethod
    def _normalize_text(value):
        """Normaliza texto eliminando acentos, espacios y mayúsculas. Copiado
        verbatim de BaseFormFiller._normalize_text."""
        if value is None:
            return ""

        text = str(value).strip().lower()
        text = unicodedata.normalize('NFKD', text)
        text = ''.join(ch for ch in text if not unicodedata.combining(ch))
        text = text.replace('-', ' ').replace('_', ' ')
        text = ''.join(ch for ch in text if ch.isalnum())
        return text

    @staticmethod
    def _is_placeholder_text(option_text):
        """Detecta textos de placeholder en español/portugués (y variantes
        comunes). Copiado verbatim de BaseFormFiller._is_placeholder_text."""
        normalized = BrowserActions._normalize_text(option_text)
        if not normalized:
            return True

        placeholder_keywords = (
            "seleccione", "selecciona", "seleccionar", "seleccion",
            "selecione", "selecionar", "selecao", "selecao",
            "escolha", "escolher", "escolhe",
            "elija", "elegir", "opcao", "opcoes", "opcion", "opciones",
            "porfavor", "favor", "obrigatorio", "required",
            "select", "choose", "please"
        )

        return any(keyword in normalized for keyword in placeholder_keywords)

    def select_by_visible_text_exact(self, select_element, text):
        """Match controlado por texto visible (NO usa Select.select_by_visible_text,
        que ante un fallo exacto hace un fallback difuso propio de Selenium y agarra
        otra option — ej. "1 mes" terminaba eligiendo "2 meses" porque contiene "mes").
        Reproduce el algoritmo de 3 pasos usado en BaseFormFiller (ver bloque
        "SI HAY VALOR EN EXCEL: match controlado" en _seleccionar_opcion_dropdown
        y equivalentes)."""
        select = Select(select_element)
        norm_desired = self._normalize_text(text)
        desired_plain = (text or "").strip()
        opts = list(select.options)
        chosen_idx = None

        # 1) exacto tal cual (texto visible idéntico)
        for i, o in enumerate(opts):
            if not self._is_placeholder_text(o.text) and (o.text or "").strip() == desired_plain:
                chosen_idx = i
                break
        # 2) exacto normalizado (ignora nbsp/acentos/espacios)
        if chosen_idx is None and norm_desired:
            for i, o in enumerate(opts):
                if not self._is_placeholder_text(o.text) and self._normalize_text(o.text) == norm_desired:
                    chosen_idx = i
                    break
        # 3) contiene normalizado (último recurso tolerante)
        if chosen_idx is None and norm_desired:
            for i, o in enumerate(opts):
                if not self._is_placeholder_text(o.text) and norm_desired in self._normalize_text(o.text):
                    chosen_idx = i
                    break

        if chosen_idx is not None:
            select.select_by_index(chosen_idx)
            return True
        return False

    def dispatch_change_input(self, element):
        """Dispara 'change' + 'input' tras una selección programática. Patrón
        usado repetidamente junto a la selección de dropdowns en
        base_form_filler.py."""
        self.driver.execute_script(
            "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));"
            "arguments[0].dispatchEvent(new Event('input', { bubbles: true }));",
            element,
        )

    # ------------------------------------------------------------------
    # iframe / geometry
    # ------------------------------------------------------------------
    def enter_frame(self, frame_element):
        self.driver.switch_to.frame(frame_element)

    def exit_frame(self):
        self.driver.switch_to.default_content()

    def all_iframes(self):
        return self.find_all(By.TAG_NAME, 'iframe')

    def scroll_into_view(self, element, block='center', behavior='instant'):
        """Copiado del patrón scrollIntoView usado en
        BaseFormFiller._scroll_element_into_view, parametrizando block/behavior."""
        self.driver.execute_script(
            "arguments[0].scrollIntoView({block: arguments[1], behavior: arguments[2]});",
            element, block, behavior,
        )

    def in_iframe(self):
        return self.exec_js('return window.self !== window.top;')

    def get_bounding_rect(self, element):
        """Copiado del patrón getBoundingClientRect usado en
        BaseFormFiller._scroll_element_into_view."""
        return self.exec_js(
            "const r = arguments[0].getBoundingClientRect();"
            "return {top: r.top, left: r.left, width: r.width, height: r.height};",
            element,
        )

    def viewport_height(self):
        return self.exec_js('return window.innerHeight;')

    def page_y_offset(self):
        return self.exec_js('return window.pageYOffset;')

    def document_ready(self):
        return self.exec_js('return document.readyState') == 'complete'

    # ------------------------------------------------------------------
    # escape hatch
    # ------------------------------------------------------------------
    def exec_js(self, script, *args):
        return self.driver.execute_script(script, *args)
