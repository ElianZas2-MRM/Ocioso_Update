"""Navegacion y manipulacion del DOM: scroll, iframes y popups de cookies.

Metodos extraidos de BaseFormFiller sin tocarles una linea. Siguen siendo parte de la
misma clase via herencia, asi que pueden llamar a sus hermanos y usar el mismo self.

Este grupo toca solo tres cosas del estado compartido (config, driver, screenshot_manager),
que es lo que lo hace el candidato mas limpio para salir primero.
"""

import time

from selenium.webdriver.common.by import By


class NavegacionDOMMixin:
    """Scroll, iframes y popups. Se usa solo como mixin de BaseFormFiller."""

    def _scroll_element_into_view(self, element):
        """Scrolls the element into view. If focused inside an iframe (especially on Safari/macOS),
        it also scrolls the parent window so that the element is fully visible in the viewport."""
        if not element:
            return
        try:
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'instant'});", element)
            time.sleep(0.1)
        except Exception:
            pass

        try:
            # Comprobar si estamos dentro de un iframe
            in_iframe = self.driver.execute_script("return window.self !== window.top;")
            if in_iframe:
                rect = self.driver.execute_script(
                    "const r = arguments[0].getBoundingClientRect(); return {top: r.top, height: r.height};",
                    element
                )
                element_top_in_iframe = rect["top"]
                element_height = rect["height"]

                # Obtener el iframe actual
                target_iframe = None
                if hasattr(self, "screenshot_manager") and self.screenshot_manager and getattr(self.screenshot_manager, "current_frame", None):
                    target_iframe = self.screenshot_manager.current_frame
                
                # Cambiar temporalmente al contenido principal
                self.driver.switch_to.default_content()

                try:
                    if not target_iframe:
                        iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
                        for iframe in iframes:
                            if iframe.is_displayed():
                                target_iframe = iframe
                                break
                    
                    if target_iframe:
                        # Obtener la posición del iframe en el parent
                        iframe_rect = self.driver.execute_script(
                            "const r = arguments[0].getBoundingClientRect(); return {top: r.top, yOffset: window.pageYOffset};",
                            target_iframe
                        )
                        iframe_top = iframe_rect["top"]
                        parent_y_offset = iframe_rect["yOffset"]

                        viewport_height = self.driver.execute_script("return window.innerHeight;")
                        
                        target_y_on_parent = parent_y_offset + iframe_top + element_top_in_iframe
                        # Centrar el elemento en la pantalla
                        scroll_y = target_y_on_parent - (viewport_height / 2) + (element_height / 2)
                        scroll_y = max(0, scroll_y)

                        self.driver.execute_script(f"window.scrollTo(0, {scroll_y});")
                        time.sleep(0.15)
                finally:
                    # Siempre volver al iframe
                    if target_iframe:
                        self.driver.switch_to.frame(target_iframe)
        except Exception as e:
            print(f"[DEBUG] Error en _scroll_element_into_view: {e}")

    def handle_gm_cookie_popup(self):
        """Maneja específicamente el popup de cookies de General Motors/Chevrolet"""
        try:
            selectors = [
                "gb-legal-notification",
                ".js-close-icon",
                ".silent-consent", 
                ".close-btn",
                "gb-legal-notification .close-btn",
            ]
            
            for selector in selectors:
                try:
                    if "gb-legal-notification" in selector:
                        popups = self.driver.find_elements(By.TAG_NAME, "gb-legal-notification")
                        if popups:
                            close_buttons = self.driver.find_elements(By.CSS_SELECTOR, ".close-btn.js-close-icon.silent-consent")
                            if close_buttons:
                                close_button = close_buttons[0]
                                if close_button.is_displayed():
                                    self.driver.execute_script("arguments[0].click();", close_button)
                                    time.sleep(1)
                                    return True
                    else:
                        elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                        for element in elements:
                            if element.is_displayed():
                                self.driver.execute_script("arguments[0].click();", element)
                                time.sleep(1)
                                return True
                except Exception:
                    continue

            return False
            
        except Exception as e:
            print(f"Error manejando popup GM: {e}")
            return False

    def handle_cookie_popups(self):
        """Maneja popups de cookies y legales que puedan aparecer"""
        if self.handle_gm_cookie_popup():
            return True
        
        cookie_selectors = [
            "button[onclick*='cookie']",
            "button[class*='cookie-accept']",
            "button[id*='cookie-accept']",
            ".cookie-accept",
            "#cookie-accept",
        ]
        
        for selector in cookie_selectors:
            try:
                elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                for element in elements:
                    if element.is_displayed():
                        self.driver.execute_script("arguments[0].click();", element)
                        time.sleep(1)
                        return True
            except Exception:
                continue

        return False

    def reposition_to_form(self, expected_form_url):
        """Vuelve a posicionarse sobre el formulario para asegurar capturas completas"""
        expected_form_url = self._sanitize_url(expected_form_url)
        if not isinstance(expected_form_url, str):
            expected_form_url = str(expected_form_url or "").strip()
        else:
            expected_form_url = expected_form_url.strip()
        try:
            self.driver.switch_to.default_content()
            if not expected_form_url:
                return True

            iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
            # Priorizar SIEMPRE el iframe GM (gm_forms/gm_front/gm_admin) que matchee el esperado.
            target_iframe, _es_gm = self._pick_gm_iframe(iframes, expected_url=expected_form_url)

            if target_iframe is not None:
                self.driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", target_iframe)
                time.sleep(1)
                self.driver.switch_to.frame(target_iframe)
                if self.screenshot_manager:
                    self.screenshot_manager.current_frame = target_iframe
                return _es_gm or (expected_form_url.lower() in self._iframe_src_of(target_iframe).lower())
            else:
                self.driver.switch_to.frame(self.driver.find_elements(By.TAG_NAME, "iframe")[0])
                return False

        except Exception as e:
            print(f"Error al reposicionar: {e}")
            try:
                self.driver.switch_to.frame(self.driver.find_elements(By.TAG_NAME, "iframe")[0])
            except:
                pass
            return False

    def _reload_form_iframe(self, expected_form_url):
        """Recarga SOLO el iframe del formulario (no toda la landing) y deja el driver
        posicionado y con el contexto cambiado adentro del iframe recargado.
        Evita tener que scrollear toda la landing de nuevo en cada reintento.
        Devuelve True si logró recargar y reposicionar sobre un iframe GM."""
        try:
            self.driver.switch_to.default_content()
        except Exception:
            pass
        try:
            iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
            target_iframe, es_gm = self._pick_gm_iframe(iframes, expected_url=expected_form_url)
            if target_iframe is None:
                return False
            # Recargar el documento del iframe. Preferimos re-setear el src (fuerza recarga
            # limpia); si no hay src usable, entramos y hacemos location.reload().
            src = self._iframe_src_of(target_iframe)
            self.driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", target_iframe)
            time.sleep(0.3)
            if src and src.startswith("http"):
                self.driver.execute_script(
                    "arguments[0].src = arguments[0].src;", target_iframe)
            else:
                self.driver.switch_to.frame(target_iframe)
                self.driver.execute_script("location.reload();")
                self.driver.switch_to.default_content()
            print("↻ Recargado solo el iframe del formulario (sin recargar la landing)")
            time.sleep(2)
            # Reposicionar sobre el iframe (puede ser un elemento nuevo tras la recarga).
            ok = self.reposition_to_form(expected_form_url)
            self.wait_for_form_ready_in_iframe()
            return bool(ok or es_gm)
        except Exception as e:
            print(f"No se pudo recargar solo el iframe: {e}")
            try:
                self.driver.switch_to.default_content()
            except Exception:
                pass
            return False

    # Techo de pasadas del pre-scroll. Una pagina con scroll infinito crece cada vez que se
    # baja, asi que sin un limite el bucle no termina nunca.
    _MAX_PASOS_PRE_SCROLL = 60

    def pre_scroll_for_dynamic_content(self):
        """Baja por la pagina disparando eventos de scroll, para que cargue lo diferido.

        Muchos formularios montan sus campos recien cuando entran en pantalla
        (IntersectionObserver, lazy-loading). Si no se baja primero, la mitad del form no
        existe todavia en el DOM cuando se lo intenta llenar.

        La altura se vuelve a medir en cada paso, y esa es la parte que importa: al bajar,
        la pagina CRECE. Antes se medía una sola vez antes de arrancar, asi que en un
        formulario largo con carga diferida el bucle cortaba con la altura inicial y se
        saltaba todo el medio — justo el caso para el que esta funcion existe. El salto
        final al fondo lo disimulaba, porque llegaba abajo pero sin haber pasado por las
        secciones intermedias, que quedaban sin cargar.
        """
        viewport_height = self.driver.execute_script("return window.innerHeight")
        scroll_step = max(viewport_height * 0.8, 800)

        headless = self.config.get('headless', False)
        step_wait = 0.5 if headless else 0.15
        end_wait = 1 if headless else 0.3

        _SCROLL_JS = (
            "window.scrollTo(0, arguments[0]);"
            "window.dispatchEvent(new Event('scroll', {bubbles:true,cancelable:false}));"
            "document.dispatchEvent(new Event('scroll', {bubbles:true}));"
        )
        _ALTO_JS = "return document.body.parentNode.scrollHeight"

        current_position = 0
        for _ in range(self._MAX_PASOS_PRE_SCROLL):
            total_height = self.driver.execute_script(_ALTO_JS)
            if current_position >= total_height:
                break
            self.driver.execute_script(_SCROLL_JS, current_position)
            time.sleep(step_wait)
            current_position += scroll_step

        self.driver.execute_script(_SCROLL_JS, 999999)
        time.sleep(end_wait)

        self.driver.execute_script(_SCROLL_JS, 0)
        time.sleep(end_wait)
