"""Casillas de verificacion y radio buttons: localizarlos, prepararlos y marcarlos.

Metodos extraidos de BaseFormFiller sin tocarles una linea. Siguen siendo parte de
la misma clase via herencia, asi que pueden llamar a sus hermanos y usar el mismo self.
"""
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.by import By
import random
import time


class CasillasYRadiosMixin:
    """Se usa solo como mixin de BaseFormFiller."""

    def build_checkbox_prefs(self, raw_headers, row):
        """
        Preferencias de checkbox tomadas del Excel: una columna cuyo encabezado es el
        `name` (o el `id`) del checkbox, con valor SI/NO. Ej: columna "test-drive" = "NO".

        Se resuelve después contra el DOM por name/id, así que una columna con SI/NO que
        no corresponda a ningún checkbox se ignora sola.
        """
        prefs = {}
        for idx, header in enumerate(raw_headers or []):
            key = str(header or "").strip().lower()
            if not key or idx >= len(row):
                continue
            value = str(row[idx] or "").strip().lower()
            if not value:
                continue
            if value in self.CHECKBOX_SI:
                prefs[key] = True
            elif value in self.CHECKBOX_NO:
                prefs[key] = False
        return prefs

    def _checkbox_pref_for(self, lower_name, checkbox_id):
        """Preferencia para este checkbox: primero el Excel (por name o id), si no hay
        columna, IDs dinámicos con valor SI/NO. None = sin preferencia."""
        prefs = getattr(self, "checkbox_prefs", None) or {}
        pref = prefs.get(lower_name)
        if pref is None:
            pref = prefs.get((checkbox_id or "").strip().lower())
        if pref is None:
            pref = self._checkbox_pref_dinamica(lower_name, checkbox_id)
        return pref

    def _checkbox_pref_dinamica(self, lower_name, checkbox_id):
        """SI/NO desde ids_dinamicos.json para checkboxes sin columna en el Excel.
        Si el ID tiene varios valores SI/NO cargados, elige uno al azar por fila."""
        try:
            if getattr(self, "_ids_din_cb_map", None) is None:
                _map = {}
                for _id, _vals in (self._cargar_ids_dinamicos() or {}).items():
                    _sino = [str(v).strip().lower() for v in (_vals or [])]
                    _sino = [v for v in _sino if v in self.CHECKBOX_SI or v in self.CHECKBOX_NO]
                    if not _sino:
                        continue
                    _map[_id.strip().lower()] = random.choice(_sino) in self.CHECKBOX_SI
                self._ids_din_cb_map = _map
            for _key in (lower_name, (checkbox_id or "").strip().lower()):
                if _key and _key in self._ids_din_cb_map:
                    return self._ids_din_cb_map[_key]
        except Exception:
            pass
        return None

    def _uncheck_checkbox(self, checkbox_id, name_attr):
        """Destilda un checkbox (el Excel pidió NO) y dispara los eventos del form."""
        try:
            self.driver.execute_script(
                """
                var id = arguments[0], name = arguments[1];
                var cb = (id && document.getElementById(id)) ||
                         (name && document.querySelector('input[type="checkbox"][name="' + name + '"]'));
                if (!cb || !cb.checked) return;
                cb.checked = false;
                cb.dispatchEvent(new Event('click',  {bubbles:true}));
                cb.dispatchEvent(new Event('change', {bubbles:true}));
                cb.checked = false;   // por si algún handler lo volvió a marcar
                """,
                checkbox_id, name_attr,
            )
        except Exception:
            pass

    def _handle_terms_checkboxes(self):
        """Marca radios y checkboxes requeridos respetando el orden de términos"""
        try:
            radios_marked = self._mark_preferred_radios()
        except Exception as e:
            print(f" Error marcando radios: {e}")
            radios_marked = 0

        try:
            checkboxes_marked = self._mark_required_checkboxes()
        except Exception as e:
            print(f" Error marcando checkboxes: {e}")
            checkboxes_marked = 0

        try:
            guide_cb = self._mark_guide_checkbox_widgets()
            checkboxes_marked += guide_cb
        except Exception as e:
            print(f" Error marcando checkboxes Guide (AEM): {e}")

        total = radios_marked + checkboxes_marked
        print(f" Radios marcados: {radios_marked}, Checkboxes marcados: {checkboxes_marked}")
        return total > 0

    def _mark_preferred_radios(self):
        groups = self._collect_radio_groups()
        marked = 0

        for group in groups:
            try:
                if self._ensure_radio_selected(group):
                    marked += 1
            except Exception as e:
                print(f" No se pudo procesar radios '{group['name']}': {e}")

        if groups:
            print(f" Radios marcados: {marked}/{len(groups)}")
        return marked

    def _collect_radio_groups(self):
        try:
            radios = self.driver.find_elements(By.CSS_SELECTOR, "input[type='radio']")
        except Exception:
            return []

        groups = {}
        for idx, radio in enumerate(radios):
            try:
                original_name_attr = (radio.get_attribute("name") or "").strip()
                has_real_name = bool(original_name_attr)
                if not original_name_attr:
                    original_name_attr = f"radio_group_{idx}"

                group_key = original_name_attr.lower()
                group = groups.setdefault(
                    group_key,
                    {
                        "name": group_key,
                        "original_name": original_name_attr,
                        "has_real_name": has_real_name,
                        "options": [],
                        "first_index": idx,
                    },
                )

                group["options"].append(
                    {
                        "value": (radio.get_attribute("value") or "").strip(),
                        "label": (radio.get_attribute("title") or radio.get_attribute("data-dtm") or radio.get_attribute("data-label") or radio.get_attribute("id") or radio.get_attribute("value") or f"opcion_{idx}"),
                        "index": idx,
                        "element": radio,
                    }
                )
            except StaleElementReferenceException:
                continue

        ordered_groups = sorted(groups.values(), key=lambda g: g["first_index"])
        return ordered_groups

    def _ensure_radio_selected(self, group):
        target_option = self._choose_radio_option(group)
        if not target_option:
            return False

        label = f"radio {group['original_name']}={target_option['value']}"

        for attempt in range(3):
            radio = self._locate_radio_candidate(group, target_option)
            if not radio:
                time.sleep(0.2)
                continue

            try:
                self._prepare_radio_for_interaction(radio)

                try:
                    state_info = (
                        radio.is_enabled(),
                        radio.is_displayed(),
                        radio.is_selected(),
                    )
                    print(f" Radio intento {attempt + 1} {label}: enabled={state_info[0]}, visible={state_info[1]}, seleccionado={state_info[2]}")
                except Exception:
                    pass

                if radio.is_selected():
                    print(f"ℹ {label} ya estaba seleccionado")
                    return True

                if self._set_radio_checked_via_js(radio):
                    radio = self._locate_radio_candidate(group, target_option) or radio
                    if radio.is_selected():
                        print(f" {label} seleccionado vía JS")
                        return True

                click_target = self._find_input_click_target(radio)
                if click_target and self._click_element_stable(click_target):
                    time.sleep(0.2)
                    radio = self._locate_radio_candidate(group, target_option) or radio
                    if radio.is_selected():
                        print(f" {label} seleccionado por click")
                        return True
                    else:
                        print(f" {label} no cambió tras click, reintentando...")

            except StaleElementReferenceException:
                time.sleep(0.2)
                continue
            except Exception as e:
                print(f" Error seleccionando {label}: {e}")
                break

        radio = self._locate_radio_candidate(group, target_option)
        return radio.is_selected() if radio else False

    def _choose_radio_option(self, group):
        options = group.get("options", [])
        if not options:
            return None

        for option in options:
            try:
                element = option.get("element")
                if element and element.is_selected():
                    return option
            except StaleElementReferenceException:
                continue

        priority_values = {
            "si": 10,
            "sí": 10,
            "yes": 9,
            "true": 9,
            "1": 9,
            "renovar": 8,
            "suscribir": 6,
            "no": 2,
        }

        name_key = group.get("name", "")
        specific_priority = {
            "have_interest": ["renovar", "suscribir"],
            "have-chevrolet": ["si", "sí", "no"],
            "client": ["si", "sí"],
            # Libro de Reclamaciones: elegir "menor de edad" abre un bloque extra obligatorio
            # (responsable legal) que el lead no tiene cómo completar.
            "cc-younger-status": ["soy mayor de edad"],
        }

        preferred_values = specific_priority.get(name_key, [])

        def option_priority(option):
            value = option.get("value", "").strip().lower()
            if preferred_values and value in preferred_values:
                return 100 - preferred_values.index(value)
            return priority_values.get(value, 1)

        sorted_options = sorted(options, key=lambda opt: (-option_priority(opt), opt.get("index", 0)))
        return sorted_options[0]

    def _locate_radio_candidate(self, group, target_option):
        name_original = group.get("original_name")
        has_real_name = group.get("has_real_name", True)
        value_target = (target_option.get("value") or "").strip()

        selectors = []

        if name_original and has_real_name:
            selectors.append((By.CSS_SELECTOR, f"input[type='radio'][name=\"{name_original}\"]"))

        selectors.append((By.CSS_SELECTOR, "input[type='radio']"))

        for by, selector in selectors:
            try:
                radios = self.driver.find_elements(by, selector)
            except Exception:
                continue

            for element in radios:
                try:
                    name_attr = (element.get_attribute("name") or "").strip()
                    if has_real_name and name_original and name_attr != name_original:
                        continue

                    if value_target:
                        value_attr = (element.get_attribute("value") or "").strip()
                        if value_attr != value_target:
                            continue

                    return element
                except StaleElementReferenceException:
                    continue

        return None

    def _prepare_radio_for_interaction(self, radio):
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
                radio,
            )
        except Exception:
            pass

    def _set_radio_checked_via_js(self, radio):
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
                radio,
            )
            if result:
                print("ℹ Radio marcado mediante JS directo")
            return bool(result)
        except Exception:
            return False

    def _checkbox_in_dom(self, checkbox):
        """True si el checkbox está en el DOM con layout, aunque tenga opacity:0."""
        try:
            return self.driver.execute_script(
                """
                    const el = arguments[0];
                    if (!el || !el.isConnected) return false;
                    const style = getComputedStyle(el);
                    if (style.display === 'none') return false;
                    if (style.visibility === 'hidden') return false;
                    // opacity:0 es ok — checkboxes custom usan esto; se hace click en label padre
                    return el.offsetWidth > 0 || el.offsetHeight > 0 || el.getClientRects().length > 0;
                """,
                checkbox,
            )
        except Exception:
            return False

    @staticmethod
    def _decidir_marca_checkbox(*, is_known, is_required, pref, tiene_identificador):
        """Qué hacer con un checkbox, en orden de prioridad (sin tocar el driver — testeable
        sin navegador). `pref` es la preferencia explícita SI/NO (True/False) leída del Excel
        o de IDs únicos por id o name (alguno de los dos siempre existe en un campo real);
        None = nadie la pidió.

        1) Excel/IDs únicos manda siempre, sea cual sea el estado del checkbox.
        2) Sin preferencia explícita: sólo se marca si es requerido (HTML required/aria-required)
           o es un checkbox de términos/privacidad conocido. Los opcionales se dejan como están
           — antes se marcaba cualquier checkbox visible, lo que tildaba de más consentimientos
           opcionales que nadie pidió.
        """
        if pref is False:
            return "uncheck"
        if pref is True:
            return "mark"
        if not tiene_identificador:
            return "skip"
        return "mark" if (is_known or is_required) else "skip"

    def _mark_required_checkboxes(self):
        known_names = {
            "terms",
            "terms-and-conditions",    # visid standard
            "terms-contact",
            "terms_contact",
            "termscontact",
            "terms-platform",
            "terms_platform",
            "termsplatform",
            "accept-terms",
            "accept_terms",
            "privacy",
            "privacy_policy",
        }

        priority_map = {
            "terms": 3,
            "terms-and-conditions": 3, # visid — misma prioridad que "terms"
            "terms-platform": 2,
            "terms_platform": 2,
            "termsplatform": 2,
            "terms-contact": 1,
            "terms_contact": 1,
            "termscontact": 1,
        }

        candidates = []

        try:
            raw_checkboxes = self.driver.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
        except Exception:
            raw_checkboxes = []

        for index, checkbox in enumerate(raw_checkboxes):
            try:
                name_attr = (checkbox.get_attribute("name") or "").strip()
                lower_name = name_attr.lower()
                checkbox_id = (checkbox.get_attribute("id") or "").strip()
                is_html_required = bool(checkbox.get_attribute("required"))
                is_aria_required = (checkbox.get_attribute("aria-required") or "").strip().lower() == "true"
                data_dtm = (checkbox.get_attribute("data-dtm") or "").strip()
                value_attr = (checkbox.get_attribute("value") or "").strip()

                is_known = lower_name in known_names

                # El Excel manda: una columna con el name/id del checkbox y valor SI/NO
                pref = self._checkbox_pref_for(lower_name, checkbox_id)
                accion = self._decidir_marca_checkbox(
                    is_known=is_known,
                    is_required=is_html_required or is_aria_required,
                    pref=pref,
                    tiene_identificador=bool(lower_name or checkbox_id),
                )
                if accion == "uncheck":
                    self._uncheck_checkbox(checkbox_id, name_attr)
                    print(f"  ⊘ {name_attr or checkbox_id} desmarcado (Excel = NO)")
                    continue
                if accion == "skip":
                    continue

                # accion == "mark": sólo si sigue en el DOM (con layout, opacity:0 incluido),
                # habilitado y no marcado todavía.
                try:
                    if not self._checkbox_in_dom(checkbox) or not checkbox.is_enabled():
                        continue
                    if checkbox.is_selected():
                        continue
                except StaleElementReferenceException:
                    continue

                priority = priority_map.get(lower_name, 0)
                display = name_attr or checkbox_id or data_dtm or "checkbox"

                candidates.append({
                    "name": lower_name,
                    "name_original": name_attr,
                    "id": checkbox_id,
                    "value": value_attr,
                    "data_dtm": data_dtm,
                    "priority": priority,
                    "order": index,
                    "label": display,
                })
            except StaleElementReferenceException:
                continue

        if not candidates:
            print(" No se encontraron checkboxes requeridos")
            return 0

        candidates.sort(key=lambda item: (item["priority"], item["order"]))

        marked = 0
        for candidate in candidates:
            if self._ensure_checkbox_selected(candidate):
                marked += 1
            else:
                print(f" No se pudo marcar {candidate.get('label', 'checkbox')}")

        print(f" Checkboxes marcados: {marked}/{len(candidates)}")
        return marked

    def _click_aem_guide_checkbox_by_input(self, checkbox_input):
        """AEM Guide suele reaccionar al click en .guideCheckBoxItem, no solo al input."""
        try:
            if checkbox_input:
                self._scroll_element_into_view(checkbox_input)
            return bool(
                self.driver.execute_script(
                    """
                    const inp = arguments[0];
                    if (!inp) return false;
                    const item = inp.closest('.guideCheckBoxItem');
                    if (item) {
                        item.click();
                    } else {
                        inp.click();
                    }
                    return inp.checked === true || inp.getAttribute('aria-checked') === 'true';
                    """,
                    checkbox_input,
                )
            )
        except Exception:
            return False

    def _mark_guide_checkbox_widgets(self):
        """Marca checkboxes de términos Adobe AEM Guide (contenedor ___guide-item + input ___N_widget)."""
        cid = self._GUIDE_CHECKBOX_CONTAINER_ID
        inp_pre = self._GUIDE_CHECKBOX_INPUT_ID_PREFIX
        selectors = (
            f'div[id="{cid}"] input[type="checkbox"]',
            f'[id^="{cid}"] input[type="checkbox"]',
            f'input[type="checkbox"][id^="{inp_pre}"][id$="_widget"]',
        )
        marked = 0
        seen = set()
        for css in selectors:
            try:
                elems = self.driver.find_elements(By.CSS_SELECTOR, css)
            except Exception:
                continue
            for cb in elems:
                try:
                    eid = (cb.get_attribute("id") or "").strip()
                    key = eid or str(cb)
                    if key in seen:
                        continue
                    if not cb.is_displayed():
                        continue
                    seen.add(key)
                    self._scroll_element_into_view(cb)
                    time.sleep(0.15)
                    # Sin name_* falsos: _locate_checkbox_candidate filtraría por name != guide_checkbox
                    candidate = {
                        "name": "",
                        "name_original": "",
                        "id": eid,
                        "value": "",
                        "data_dtm": "",
                        "priority": 100,
                        "order": 0,
                        "label": f"guide_checkbox:{eid or 'sin-id'}",
                    }
                    if self._ensure_checkbox_selected(candidate):
                        marked += 1
                        continue
                    el = self.safe_find_element(By.ID, eid) if eid else None
                    if el and self._click_aem_guide_checkbox_by_input(el):
                        time.sleep(0.25)
                        el2 = self.safe_find_element(By.ID, eid) if eid else None
                        if el2 and el2.is_selected():
                            marked += 1
                            print(f" Guide checkbox (click AEM): {eid}")
                except StaleElementReferenceException:
                    continue
                except Exception as e:
                    print(f" Guide checkbox: {e}")
        if marked:
            print(f" Checkboxes Guide (AEM) marcados: {marked}")
        return marked

    def _ensure_checkbox_selected(self, candidate):
        label = candidate.get("label", "checkbox")

        for attempt in range(4):
            checkbox = self._locate_checkbox_candidate(candidate)
            if not checkbox:
                time.sleep(0.2)
                continue

            try:
                self._prepare_checkbox_for_interaction(checkbox)

                try:
                    state_info = (
                        checkbox.is_enabled(),
                        checkbox.is_displayed(),
                        checkbox.is_selected(),
                    )
                    print(f" Intento {attempt + 1} para {label}: enabled={state_info[0]}, visible={state_info[1]}, seleccionado={state_info[2]}")
                except Exception:
                    pass

                if checkbox.is_selected():
                    print(f"ℹ {label} ya estaba marcado")
                    return True

                if self._set_checkbox_checked_via_js(checkbox):
                    checkbox = self._locate_checkbox_candidate(candidate) or checkbox
                    if checkbox.is_selected():
                        print(f" {label} marcado vía JS")
                        return True

                click_target = self._find_input_click_target(checkbox)
                if click_target and self._click_element_stable(click_target):
                    time.sleep(0.2)
                    checkbox = self._locate_checkbox_candidate(candidate) or checkbox
                    if checkbox.is_selected():
                        print(f" {label} marcado por click")
                        return True
                    else:
                        print(f" {label} no cambió tras click, reintentando...")

            except StaleElementReferenceException:
                time.sleep(0.2)
                continue
            except Exception as e:
                print(f" Error marcando {label}: {e}")
                break

        checkbox = self._locate_checkbox_candidate(candidate)
        return checkbox.is_selected() if checkbox else False

    def _prepare_checkbox_for_interaction(self, checkbox):
        try:
            self.driver.execute_script(
                """
                    const cb = arguments[0];
                    if (cb.hasAttribute('disabled')) {
                        cb.removeAttribute('disabled');
                    }
                    cb.disabled = false;
                    cb.setAttribute('aria-disabled', 'false');
                    cb.classList.remove('is-invalid');
                    if (cb.tabIndex === -1) {
                        cb.tabIndex = 0;
                    }
                """,
                checkbox,
            )
        except Exception:
            pass

    def _locate_checkbox_candidate(self, candidate):
        selectors = []
        checkbox_id = candidate.get("id")
        name_original = candidate.get("name_original")
        lower_name = candidate.get("name")
        data_dtm = candidate.get("data_dtm")

        if checkbox_id:
            selectors.append((By.ID, checkbox_id))

        if name_original:
            selectors.append((By.CSS_SELECTOR, f"input[type='checkbox'][name=\"{name_original}\"]"))

        if data_dtm:
            selectors.append((By.CSS_SELECTOR, f"input[type='checkbox'][data-dtm=\"{data_dtm}\"]"))

        selectors.append((By.CSS_SELECTOR, "input[type='checkbox']"))

        for by, selector in selectors:
            try:
                elements = self.driver.find_elements(by, selector)
            except Exception:
                continue

            for element in elements:
                try:
                    if checkbox_id and (element.get_attribute("id") or "").strip() != checkbox_id:
                        continue

                    if name_original and (element.get_attribute("name") or "").strip() != name_original:
                        continue

                    if lower_name and (element.get_attribute("name") or "").strip().lower() != lower_name:
                        continue

                    if data_dtm and (element.get_attribute("data-dtm") or "").strip() != data_dtm:
                        continue

                    value_attr = (element.get_attribute("value") or "").strip()
                    candidate_value = candidate.get("value", "")
                    # Solo validar value si ambos tienen valores no vacíos
                    if candidate_value and value_attr and candidate_value != value_attr:
                        continue

                    return element
                except StaleElementReferenceException:
                    continue

        return None

    def _set_checkbox_checked_via_js(self, checkbox):
        try:
            result = self.driver.execute_script(
                """
                    const cb = arguments[0];
                    if (cb.checked) {
                        return true;
                    }
                    cb.focus && cb.focus();
                    // Disparar click en el label padre primero (frameworks React/Angular)
                    const parentLabel = cb.closest('label');
                    if (parentLabel) {
                        parentLabel.click();
                        if (cb.checked) return true;
                    }
                    // Fallback: click directo sobre el input
                    cb.click();
                    if (cb.checked) return true;
                    // Fallback: forzar via propiedad + eventos
                    cb.checked = true;
                    cb.setAttribute('checked', 'checked');
                    cb.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
                    cb.dispatchEvent(new Event('input', { bubbles: true }));
                    cb.dispatchEvent(new Event('change', { bubbles: true }));
                    return cb.checked === true;
                """,
                checkbox,
            )
            if result:
                print("ℹ Checkbox marcado mediante JS directo")
            return bool(result)
        except Exception:
            return False
