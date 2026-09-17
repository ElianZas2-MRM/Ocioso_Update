"""Listas desplegables (<select>): elegir opcion, dependencias padre-hijo y validacion.

Metodos extraidos de BaseFormFiller sin tocarles una linea. Siguen siendo parte de
la misma clase via herencia, asi que pueden llamar a sus hermanos y usar el mismo self.
"""
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
import random
import re
import time


class DesplegablesMixin:
    """Se usa solo como mixin de BaseFormFiller."""

    def _get_selected_text_for_select(self, select_element):
        """Obtiene el texto seleccionado actual de un select simple."""
        try:
            selected_option = Select(select_element).first_selected_option
            if selected_option:
                return (selected_option.text or "").strip()
        except Exception:
            return ""
        return ""

    def _is_option_disabled(self, option_element):
        """Determina si una option está deshabilitada de forma robusta."""
        try:
            disabled_attr = option_element.get_attribute("disabled")
            if disabled_attr is None:
                return False
            normalized = str(disabled_attr).strip().lower()
            # HTML boolean attrs suelen venir como "true", "disabled" o "".
            # Algunos frameworks pueden devolver "false" y no debe bloquearse.
            return normalized in ("", "true", "disabled")
        except Exception:
            return False

    def _get_valid_select_options(self, select_element):
        """Retorna opciones seleccionables (no placeholder y no disabled)."""
        valid_options = []
        try:
            select = Select(select_element)
        except Exception:
            return valid_options

        for opt in select.options:
            opt_text = (opt.text or "").strip()
            if not opt_text:
                continue
            if self._is_option_disabled(opt):
                continue
            if self._is_placeholder_text(opt_text):
                continue
            valid_options.append(opt)

        return valid_options

    def _registrar_dropdowns_en_placeholder(self):
        """Anota los <select> visibles que siguen en su placeholder ("Seleccionar",
        "Selecione", "Escolha"…). Devuelve la lista de ids anotados en esta pasada.

        Se llama al cerrar cada paso de un form multi-paso: al avanzar, el paso anterior
        puede salir del DOM y entonces el chequeo previo al submit ya no lo ve.
        """
        try:
            selects = self.driver.execute_script("""
                var out = [];
                document.querySelectorAll('select').forEach(function(s){
                    if (!s.getClientRects().length) return;
                    var id = s.id || s.getAttribute('name') || '';
                    if (!id) return;
                    var o = s.selectedIndex >= 0 ? s.options[s.selectedIndex] : null;
                    out.push({id: id, texto: o ? (o.text || '').trim() : ''});
                });
                return out;
            """) or []
        except Exception:
            return []

        nuevos = []
        for s in selects:
            fid, texto = s.get("id", ""), s.get("texto", "")
            if not fid or not self._is_placeholder_text(texto):
                continue
            # Un campo que el Excel pidió omitir no es un dato que faltó. Es el caso
            # del concesionario de Cadillac: viene bloqueado a propósito, no lo puede
            # elegir nadie, y anotarlo marcaba como sucia una corrida que estuvo bien.
            if fid in getattr(self, "_campos_omitidos", ()):  
                continue
            if fid not in self._dropdowns_sin_elegir:
                self._dropdowns_sin_elegir.append(fid)
                nuevos.append(fid)
            # No puede quedar registrado como si fuera un valor elegido.
            for clave in [k for k in self.current_row_field_values
                          if k.split("::", 1)[-1] == fid]:
                self.current_row_field_values.pop(clave, None)
        if nuevos:
            print(f"⚠ Dropdowns sin elegir en el paso {getattr(self, '_current_step', 1)}: "
                  f"{', '.join(nuevos)}")
        return nuevos

    def _select_has_valid_selected_option(self, select_id):
        """Indica si un select ya tiene una opción elegida que no parece placeholder."""
        try:
            select_element = self.safe_find_element(By.ID, select_id)
            if not select_element:
                return False
            select = Select(select_element)
            selected_option = select.first_selected_option
            if not selected_option:
                return False

            selected_text = (selected_option.text or "").strip()
            if not selected_text:
                return False

            return not self._is_placeholder_text(selected_text)
        except Exception:
            return False

    def _select_current_option_matches_desired(self, select_id, desired_value):
        """True si el select ya tiene una opción válida que coincide con el valor pedido."""
        desired = (desired_value or "").strip()
        if not desired:
            return self._select_has_valid_selected_option(select_id)

        try:
            select_element = self.safe_find_element(By.ID, select_id)
            if not select_element:
                return False

            selected_option = Select(select_element).first_selected_option
            selected_text = (selected_option.text or "").strip()
            if not selected_text or self._is_placeholder_text(selected_text):
                return False

            selected_norm = self._normalize_text(selected_text)
            desired_norm = self._normalize_text(desired)
            if not selected_norm or not desired_norm:
                return False

            return selected_norm == desired_norm or selected_norm in desired_norm or desired_norm in selected_norm
        except Exception:
            return False

    def _get_dependency_wait_settings(self):
        """Obtiene timeout/poll/retries para dependencias entre dropdowns."""
        try:
            timeout = float(self.config.get("dependency_dropdown_timeout", 8.0))
        except Exception:
            timeout = 8.0
        try:
            poll_interval = float(self.config.get("dependency_dropdown_poll_interval", 0.2))
        except Exception:
            poll_interval = 0.2
        try:
            retries = int(self.config.get("dependency_selection_retries", 2))
        except Exception:
            retries = 2

        timeout = max(1.0, timeout)
        poll_interval = max(0.05, poll_interval)
        retries = max(1, retries)
        return timeout, poll_interval, retries

    def _wait_for_dependent_dropdown_ready(self, child_select_id, parent_id=None):
        """Espera a que el dropdown dependiente esté visible/habilitado y con opciones válidas."""
        timeout, poll_interval, _ = self._get_dependency_wait_settings()
        deadline = time.time() + timeout

        while time.time() < deadline:
            try:
                child_element = self.safe_find_element(By.ID, child_select_id)
                if not child_element or not child_element.is_displayed():
                    time.sleep(poll_interval)
                    continue

                tag_name = (child_element.tag_name or "").lower()
                if tag_name != "select":
                    return True

                if child_element.get_attribute("disabled") or not child_element.is_enabled():
                    time.sleep(poll_interval)
                    continue

                if self._get_valid_select_options(child_element):
                    return True
            except Exception:
                pass

            time.sleep(poll_interval)

        if parent_id:
            print(
                f" Timeout esperando dependiente '{child_select_id}' después de '{parent_id}' "
                f"({timeout:.1f}s)"
            )
        else:
            print(f" Timeout esperando dependiente '{child_select_id}' ({timeout:.1f}s)")
        return False

    def _select_dependency_child_with_timeout(self, child_select_id, child_value, child_field_name, parent_id):
        """Selecciona dropdown hijo con espera activa y reintentos."""
        _, _, retries = self._get_dependency_wait_settings()

        child_ready = self._wait_for_dependent_dropdown_ready(child_select_id, parent_id=parent_id)
        if not child_ready:
            raise ValueError(f"No se encontraron opciones en el dropdown '{child_field_name}' al depender de '{parent_id}'")

        for attempt in range(1, retries + 1):
            if self._select_has_valid_selected_option(child_select_id):
                print(f" {child_field_name} ya tiene valor válido, se conserva")
                return True

            selected = self.safe_select_option_if_visible(child_select_id, child_value, child_field_name)
            if selected and self._select_has_valid_selected_option(child_select_id):
                return True

            if attempt < retries:
                time.sleep(0.4)

        return False

    def safe_select_option_if_visible(self, select_id, option_text, field_name):
        """Selecciona una opción solo si el dropdown está visible - CON AUTO-SELECCIÓN ALEATORIA SI ESTÁ VACÍO"""
        if option_text is None:
            option_text = ""
        else:
            option_text = str(option_text)
        
        # Determinar si el campo está vacío (también si el valor es un placeholder)
        is_empty = not option_text or option_text.strip() == "" or self._is_placeholder_text(option_text)
            
        try:
            # Verificar primero si el elemento existe
            try:
                select_element = WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.ID, select_id))
                )
            except TimeoutException:
                print(f" {field_name} (ID: {select_id}) no encontrado en el DOM después de 5s")
                return False
            
            # Verificar si está visible
            if not select_element.is_displayed():
                print(f" {field_name} (ID: {select_id}) existe pero NO está visible")
                return False

            # Forzar habilitación si está deshabilitado
            if select_element.get_attribute("disabled") or not select_element.is_enabled():
                try:
                    WebDriverWait(self.driver, 4).until(
                        EC.element_to_be_clickable((By.ID, select_id))
                    )
                    select_element = self.safe_find_element(By.ID, select_id)
                except TimeoutException:
                    try:
                        self.driver.execute_script("arguments[0].removeAttribute('disabled'); arguments[0].disabled = false;", select_element)
                        print(f" {field_name} habilitado forzosamente")
                    except Exception:
                        print(f"{field_name} está deshabilitado, omitiendo...")
                        return False

            # Asegurar que el select tenga opciones disponibles (solo si no está vacío)
            if not is_empty:
                try:
                    def select_has_options(_):
                        try:
                            select_el = self.safe_find_element(By.ID, select_id)
                            return len(Select(select_el).options) > 0
                        except Exception:
                            return False

                    WebDriverWait(self.driver, 4).until(select_has_options)
                except Exception:
                    pass

            select = Select(select_element)
            is_multiple = select_element.get_attribute("multiple") is not None
            
            # Detectar si es un select múltiple (como kits[])
            if is_multiple:
                print(f" {field_name} es multi-select, usando lógica especial")
                return self._select_multiple_options(select_element, option_text, field_name)

            if is_empty:
                try:
                    selected_option = select.first_selected_option
                except Exception:
                    selected_option = None

                if selected_option:
                    selected_text = (selected_option.text or "").strip()
                    if selected_text and not self._is_placeholder_text(selected_text):
                        print(f" {field_name} ya tiene valor válido ('{selected_text}'), se conserva")
                        self._record_field_value(select_id, selected_text)
                        return True

            # SI EL CAMPO ESTÁ VACÍO: Seleccionar opción ALEATORIA
            if is_empty:
                valid_options = []

                # Esperar opciones válidas para dropdowns que cargan dinámicamente.
                try:
                    def has_valid_options(_):
                        current_select = self.safe_find_element(By.ID, select_id)
                        if not current_select:
                            return False
                        options_now = self._get_valid_select_options(current_select)
                        return options_now if options_now else False

                    valid_options = WebDriverWait(self.driver, 8).until(has_valid_options)
                    select_element = self.safe_find_element(By.ID, select_id) or select_element
                    select = Select(select_element)
                except TimeoutException:
                    # Último intento sin wait para reportar correctamente.
                    valid_options = self._get_valid_select_options(select_element)
                except Exception:
                    valid_options = self._get_valid_select_options(select_element)
                
                if valid_options:
                    random_option = random.choice(valid_options)
                    select.select_by_index(select.options.index(random_option))
                    # Disparar evento change para actualizar dropdowns dependientes
                    self.driver.execute_script(
                        "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));"
                        "arguments[0].dispatchEvent(new Event('input', { bubbles: true }));",
                        select_element
                    )
                    print(f"🎲 {field_name} - Auto-seleccionado (aleatorio): {random_option.text}")
                    self._record_field_value(select_id, random_option.text)
                    return True
                else:
                    print(f" {field_name} - No hay opciones válidas para auto-selección")
                    raise ValueError(f"No se encontraron opciones en el dropdown '{field_name}'")

            # SI HAY VALOR EN EXCEL: match controlado (NO usar select_by_visible_text,
            # que ante un fallo exacto hace un fallback difuso propio de Selenium y agarra
            # otra option — ej. "1 mes" terminaba eligiendo "2 meses" porque contiene "mes").
            try:
                norm_desired = self._normalize_text(option_text)
                desired_plain = option_text.strip()
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
                    self.driver.execute_script(
                        "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));"
                        "arguments[0].dispatchEvent(new Event('input', { bubbles: true }));",
                        select_element
                    )
                    selected_text = (opts[chosen_idx].text or option_text).strip()
                    print(f" {field_name} seleccionado: {selected_text}")
                    self._record_field_value(select_id, selected_text)
                    return True

                print(f"❌ [DEBUG-FILL] No se encontró opción para {field_name}: '{option_text}'.")
                options_now = self._get_valid_select_options(select_element)
                try:
                    options = [o.text.strip() for o in select.options if o.text.strip()]
                    print(f"   💡 Opciones disponibles en dropdown '{field_name}': {options}")
                except Exception:
                    pass
                if not options_now:
                    raise ValueError(f"No se encontraron opciones en el dropdown '{field_name}'")
                if not hasattr(self, "_campos_dropdown_no_encontrados"):
                    self._campos_dropdown_no_encontrados = []
                self._campos_dropdown_no_encontrados.append(f"{field_name}: '{option_text}'")
                return False
            except ValueError:
                raise
            except Exception as e:
                print(f" Error resolviendo opción de {field_name}: {e}")
                return False
                
        except Exception as e:
            print(f" Error en {field_name}: {e}")
            return False

    def _select_dropdown_with_fallback(self, select_id, desired_value, field_name, allow_text_fallback=True, fallback_ids=None):
        """Selecciona opciones tolerando diferencias de texto, con auto-selección aleatoria si está vacío"""
        # Determinar si el campo está vacío
        is_empty = not desired_value or str(desired_value).strip() == ""
        
        fallback_ids = fallback_ids or [select_id]

        try:
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, select_id))
            )
        except Exception as e:
            print(f" {field_name} con id='{select_id}' no disponible: {e}")
            if allow_text_fallback and not is_empty:
                return self._fill_text_field_direct(fallback_ids, desired_value, field_name)
            return False

        tag_name = ''
        try:
            tag_name = element.tag_name.lower()
        except Exception:
            tag_name = ''

        if tag_name != 'select':
            print(f" {field_name} con id='{select_id}' no es dropdown, intentando como texto...")
            if allow_text_fallback and not is_empty:
                return self._fill_text_field_direct(fallback_ids, desired_value, field_name)
            return False

        try:
            self._scroll_element_into_view(element)
            time.sleep(0.3)
        except Exception:
            pass

        options = element.find_elements(By.TAG_NAME, "option")
        if not options:
            print(f" {field_name} no tiene opciones disponibles")
            raise ValueError(f"No se encontraron opciones en el dropdown '{field_name}'")

        if is_empty:
            try:
                selected_option = Select(element).first_selected_option
            except Exception:
                selected_option = None

            if selected_option:
                selected_text = (selected_option.text or "").strip()
                if selected_text and not self._is_placeholder_text(selected_text):
                    print(f" {field_name} ya tiene valor válido ('{selected_text}'), se conserva")
                    return True

        # SI EL CAMPO ESTÁ VACÍO: Auto-seleccionar opción ALEATORIA
        if is_empty:
            valid_options = []
            for opt in options:
                opt_text = opt.text.strip()
                is_placeholder = self._is_placeholder_text(opt_text)
                if opt_text and not opt.get_attribute("disabled") and not is_placeholder:
                    valid_options.append(opt)
            
            if valid_options:
                random_option = random.choice(valid_options)
                option_value = random_option.get_attribute("value") or random_option.text
                option_text = random_option.text.strip()
                try:
                    Select(element).select_by_value(option_value)
                    # Disparar evento change para actualizar dropdowns dependientes
                    self.driver.execute_script(
                        "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));"
                        "arguments[0].dispatchEvent(new Event('input', { bubbles: true }));",
                        element
                    )
                    print(f"🎲 {field_name} - Auto-seleccionado (aleatorio): {option_text}")
                    return True
                except:
                    pass
            print(f" {field_name} - No hay opciones válidas para auto-selección")
            raise ValueError(f"No se encontraron opciones en el dropdown '{field_name}'")

        # SI HAY VALOR: Intentar usar el método optimizado primero
        try:
            if self.safe_select_option_if_visible(select_id, desired_value, field_name):
                return True
        except Exception as e:
            if isinstance(e, ValueError) and "No se encontraron opciones" in str(e):
                raise

        # Búsqueda avanzada con normalización
        normalized_target = self._normalize_text(desired_value)
        fallback_option = None

        for opt in options:
            opt_text = opt.text.strip()
            opt_value = opt.get_attribute("value") or ""

            if opt_text and desired_value.strip().lower() == opt_text.lower():
                fallback_option = opt
                break

            if normalized_target:
                if normalized_target == self._normalize_text(opt_text) or normalized_target == self._normalize_text(opt_value):
                    fallback_option = opt
                    break

        if not fallback_option and normalized_target:
            for opt in options:
                opt_text = opt.text.strip()
                opt_value = opt.get_attribute("value") or ""

                if normalized_target in self._normalize_text(opt_text) or normalized_target in self._normalize_text(opt_value):
                    fallback_option = opt
                    break

        if not fallback_option:
            print(f" Sin coincidencias para {field_name} con valor '{desired_value}'")
            if allow_text_fallback:
                return self._fill_text_field_direct(fallback_ids, desired_value, field_name)
            return False

        option_value = fallback_option.get_attribute("value") or fallback_option.text
        option_text = fallback_option.text.strip()

        try:
            Select(element).select_by_value(option_value)
            print(f" {field_name} seleccionado (match flexible): {option_text or option_value}")
            return True
        except Exception:
            try:
                if option_text:
                    Select(element).select_by_visible_text(option_text)
                    print(f" {field_name} seleccionado (por texto): {option_text}")
                    return True
            except Exception:
                pass

        try:
            self.driver.execute_script(
                "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change', {bubbles: true}));",
                element,
                option_value
            )
            print(f" {field_name} asignado vía JS: {option_text or option_value}")
            return True
        except Exception as e:
            print(f" Error completando {field_name} con id='{select_id}': {e}")
            if allow_text_fallback:
                return self._fill_text_field_direct(fallback_ids, desired_value, field_name)
            return False

    def _get_mapped_select_ids(self, field_mapping=None):
        """Obtiene IDs de campos que tienen un valor de datos asignado (data_index o data_key).
        Los campos auto-descubiertos sin asignación NO se incluyen para que _auto_fill los rellene."""
        mapped_ids = set()
        mapping = field_mapping if field_mapping is not None else self.field_mapping or []
        for field_config in mapping:
            has_data = (
                field_config.get("data_index") is not None
                or field_config.get("data_key")
            )
            if not has_data:
                continue
            field_id = field_config.get("id")
            if isinstance(field_id, list):
                for fid in field_id:
                    if fid:
                        mapped_ids.add(fid)
            elif field_id:
                mapped_ids.add(field_id)
        return mapped_ids

    def _select_multiple_options(self, select_element, option_text, field_name):
        """Selecciona múltiples opciones en un select múltiple usando JavaScript"""
        try:
            # Dividir valores si vienen separados por coma, punto y coma, etc.
            values_to_select = [v.strip() for v in re.split(r'[,;|]', option_text) if v.strip()]
            if not values_to_select:
                values_to_select = [option_text.strip()]
            
            print(f" Intentando seleccionar en {field_name}: {values_to_select}")
            
            selected_count = 0
            for value in values_to_select:
                # Usar JavaScript para seleccionar sin desmarcar otras opciones
                result = self.driver.execute_script("""
                    const select = arguments[0];
                    const searchText = arguments[1].toLowerCase();
                    let matched = false;
                    
                    for (let i = 0; i < select.options.length; i++) {
                        const opt = select.options[i];
                        const optText = opt.text.trim().toLowerCase();
                        const optValue = (opt.value || '').toLowerCase();
                        
                        // Buscar coincidencia exacta o parcial
                        if (optText === searchText || optValue === searchText || 
                            optText.includes(searchText) || optValue.includes(searchText)) {
                            opt.selected = true;
                            matched = true;
                            console.log('Seleccionado:', opt.text);
                            break;
                        }
                    }
                    
                    if (matched) {
                        select.dispatchEvent(new Event('input', { bubbles: true }));
                        select.dispatchEvent(new Event('change', { bubbles: true }));
                    }
                    
                    return matched;
                """, select_element, value)
                
                if result:
                    selected_count += 1
                    print(f"   Seleccionado: {value}")
                else:
                    print(f"   No se encontró: {value}")
            
            if selected_count > 0:
                print(f" {field_name}: {selected_count}/{len(values_to_select)} opciones seleccionadas")
                try:
                    selected_labels = [opt.text.strip() for opt in Select(select_element).all_selected_options if opt.text.strip()]
                except Exception:
                    selected_labels = values_to_select
                select_id = select_element.get_attribute("id") or field_name
                self._record_field_value(select_id, selected_labels)
                return True
            else:
                print(f" {field_name}: No se pudo seleccionar ninguna opción")
                return False
                
        except Exception as e:
            print(f" Error seleccionando múltiples opciones en {field_name}: {e}")
            return False
