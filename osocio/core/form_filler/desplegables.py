"""Listas desplegables (<select>): elegir opcion, dependencias padre-hijo y validacion.

Metodos extraidos de BaseFormFiller sin tocarles una linea. Siguen siendo parte de
la misma clase via herencia, asi que pueden llamar a sus hermanos y usar el mismo self.
"""
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.remote.webelement import WebElement
import random
import re
import time


class _OpcionConTexto(WebElement):
    """Un <option> con el texto ya leído: .text no hace otro viaje al navegador."""

    @property
    def text(self):
        return self._texto


class DesplegablesMixin:
    """Se usa solo como mixin de BaseFormFiller."""

    def _get_selected_text_for_select(self, select_element):
        """Obtiene el texto seleccionado actual de un select simple.

        Por JS en un solo viaje: Select.first_selected_option le pregunta a CADA option si
        está seleccionada (un comando por option, lento en un device de la nube).
        """
        try:
            return (self.driver.execute_script(
                "var s = arguments[0];"
                "return (s && s.selectedIndex >= 0) ? (s.options[s.selectedIndex].text || '') : '';",
                select_element) or "").strip()
        except Exception:
            return ""

    def _textos_de_opciones(self, select_element):
        """El texto de todas las options del select, en orden, en un solo viaje."""
        try:
            return [str(t or "") for t in (self.driver.execute_script(
                "return Array.prototype.map.call(arguments[0].options, function (o) {"
                "  return (o.text || '').trim(); });", select_element) or [])]
        except Exception:
            return [(o.text or "").strip() for o in Select(select_element).options]

    def _es_navegador_remoto(self):
        return str(self.config.get("browser", "")).startswith("browserstack_")

    def _elegir_opcion_por_indice(self, select, select_element, indice):
        """Selecciona la option N. En la nube por JS (Select.select_by_index lee el
        atributo 'index' de cada option hasta encontrarla: un viaje por option); en los
        navegadores locales, con Selenium como siempre. Los eventos change/input los
        dispara quien llama, igual en los dos casos."""
        if self._es_navegador_remoto():
            self.driver.execute_script(
                "var s = arguments[0], i = arguments[1];"
                "var d = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'selectedIndex');"
                "if (d && d.set) { d.set.call(s, i); } else { s.selectedIndex = i; }",
                select_element, int(indice))
        else:
            select.select_by_index(indice)

    # ── Rotación del día: región / ciudad / concesionario al azar sin repetirse ──────────
    # (osocio/utils/rotacion_dealers.py). Lo que pide el Excel manda; lo elegido en la fila
    # se anota en self._rotacion_fila y se registra como usado al terminar la fila.

    def _rotacion(self):
        rot = getattr(self, "_rotacion_dia", None)
        if rot is None:
            from osocio.utils.rotacion_dealers import RotacionDelDia
            rot = self._rotacion_dia = RotacionDelDia()
        return rot

    def _pais_rotacion(self):
        return str(self.config.get("pais", "") or "")

    def _rotacion_anotar(self, tipo, select_id, texto, al_azar):
        if not tipo or not texto:
            return
        fila = getattr(self, "_rotacion_fila", None)
        if fila is None:
            fila = self._rotacion_fila = {}
        fila[tipo] = {"id": select_id, "texto": str(texto).strip(), "al_azar": al_azar}

    def _elegir_al_azar(self, select_id, field_name, opciones, excluir=()):
        """La opción al azar de un dropdown. Región, ciudad y concesionario rotan: se elige
        entre las menos usadas hoy, así no se repiten en las corridas del día."""
        from osocio.utils.rotacion_dealers import tipo_de_campo
        tipo = tipo_de_campo(select_id, field_name)
        if not tipo or not opciones:
            return random.choice(opciones)
        i = self._rotacion().elegir(self._pais_rotacion(), tipo, [o.text for o in opciones], excluir=excluir)
        elegida = opciones[i]
        self._rotacion_anotar(tipo, select_id, elegida.text, al_azar=True)
        return elegida

    def _reelegir_al_azar(self, select_id, tipo, excluir):
        """Vuelve a elegir al azar (con rotación) en un dropdown ya completado, descartando
        `excluir`. Devuelve el texto elegido, o None si no queda otra opción."""
        from osocio.utils.rotacion_dealers import _norm
        el = self.safe_find_element(By.ID, select_id)
        if el is None:
            return None
        opciones = [o for o in self._get_valid_select_options(el) if _norm(o.text) not in {_norm(e) for e in excluir}]
        if not opciones:
            return None
        i = self._rotacion().elegir(self._pais_rotacion(), tipo, [o.text for o in opciones])
        elegida = opciones[i]
        _idx = getattr(elegida, "_indice", None)
        if _idx is None:
            _idx = Select(el).options.index(elegida)
        self._elegir_opcion_por_indice(Select(el), el, _idx)
        self.driver.execute_script(
            "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));"
            "arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", el)
        self._rotacion_anotar(tipo, select_id, elegida.text, al_azar=True)
        self._record_field_value(select_id, elegida.text)
        return elegida.text

    def _evitar_concesionario_repetido(self):
        """Si el concesionario elegido al azar ya se usó hoy, busca uno sin usar cambiando de
        ciudad y, si hace falta, de región (sólo los que se eligieron al azar: lo del Excel
        no se toca). Si no queda ninguno sin usar, lo deja anotado para el resultado."""
        from osocio.utils.rotacion_dealers import CIUDAD, DEALER, REGION
        fila = getattr(self, "_rotacion_fila", None) or {}
        dealer = fila.get(DEALER)
        if not dealer or not dealer["al_azar"]:
            return
        rot, pais = self._rotacion(), self._pais_rotacion()
        if rot.veces(pais, DEALER, dealer["texto"]) == 0:
            return

        def dealer_sin_usar():
            if not self._wait_for_dependent_dropdown_ready(dealer["id"]):
                return None
            texto = self._reelegir_al_azar(dealer["id"], DEALER, excluir=())
            return texto if texto and rot.veces(pais, DEALER, texto) == 0 else None

        # 1) Otro concesionario de la misma ciudad (por si quedaba uno sin usar).
        if dealer_sin_usar():
            return
        ciudad, region = fila.get(CIUDAD), fila.get(REGION)
        regiones_probadas = {region["texto"]} if region else set()
        for _vuelta in range(4):
            # 2) Otras ciudades de la región actual.
            if ciudad and ciudad["al_azar"]:
                ciudades_probadas = {fila[CIUDAD]["texto"]}
                for _ in range(6):
                    nueva = self._reelegir_al_azar(ciudad["id"], CIUDAD, excluir=ciudades_probadas)
                    if not nueva:
                        break
                    ciudades_probadas.add(nueva)
                    if dealer_sin_usar():
                        print(f"🔁 Concesionario sin usar hoy: {fila[DEALER]['texto']} (ciudad {nueva})")
                        return
            # 3) Otra región.
            if not (region and region["al_azar"] and ciudad and ciudad["al_azar"]):
                break
            nueva_r = self._reelegir_al_azar(region["id"], REGION, excluir=regiones_probadas)
            if not nueva_r:
                break
            regiones_probadas.add(nueva_r)
            if not self._wait_for_dependent_dropdown_ready(ciudad["id"]):
                continue
            self._reelegir_al_azar(ciudad["id"], CIUDAD, excluir=())
            if dealer_sin_usar():
                print(f"🔁 Concesionario sin usar hoy: {fila[DEALER]['texto']} (región {nueva_r})")
                return
        texto = fila[DEALER]["texto"]
        unico = len(self._get_valid_select_options(self.safe_find_element(By.ID, dealer["id"])) or []) <= 1 \
            and not (ciudad and ciudad["al_azar"])
        self._aviso_rotacion = (f"concesionario repetido hoy: {texto} "
                                f"({'es el único que ofrece el form' if unico else 'ya se usaron hoy todos los que ofrece el form'})")
        print(f"⚠️ {self._aviso_rotacion}")

    def _registrar_rotacion_fila(self):
        """Al terminar la fila: anota como usados hoy la región, ciudad y concesionario que
        quedaron (salvo en corridas de prueba sin envío) y limpia lo de la fila."""
        fila = getattr(self, "_rotacion_fila", None) or {}
        sin_envio = bool(self.config.get("chequeo_campos") or self.config.get("no_enviar_lead")
                         or self.config.get("solo_verificar_visual"))
        if fila and not sin_envio:
            try:
                self._rotacion().registrar(self._pais_rotacion(), {t: v["texto"] for t, v in fila.items()})
            except Exception as e:
                print(f" Rotación de concesionarios: no se pudo anotar — {e}")
        self._rotacion_fila = {}

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
        """Retorna opciones seleccionables (no placeholder y no disabled).

        Un solo viaje al navegador: antes se leía el texto y el 'disabled' de cada
        <option> por separado, y en un device de la nube (BrowserStack) eso eran cientos
        de comandos de ~0,35 s por lead. Devuelve los mismos WebElement de <option>, con
        el texto ya leído (leer .text no vuelve a viajar).
        """
        try:
            crudas = self.driver.execute_script(
                "var s = arguments[0], out = [];"
                "if (!s || !s.options) return out;"
                "for (var i = 0; i < s.options.length; i++) {"
                "  var o = s.options[i], t = (o.text || '').trim();"
                "  if (t && !o.disabled) out.push({el: o, texto: t, indice: i});"
                "}"
                "return out;", select_element) or []
        except Exception:
            return []
        # Los <option> que vinieron dentro del objeto JS no sirven en el iPhone de
        # BrowserStack (ver _con_elementos_reales): se toman los mismos por su índice.
        try:
            todas = select_element.find_elements(By.TAG_NAME, "option")
            for item in crudas:
                i = item.get("indice")
                if isinstance(i, int) and 0 <= i < len(todas):
                    item["el"] = todas[i]
        except Exception:
            pass
        valid_options = []
        for item in crudas:
            texto = item.get("texto", "")
            if self._is_placeholder_text(texto):
                continue
            opcion = item.get("el")
            if opcion is None:
                continue
            opcion.__class__ = _OpcionConTexto
            opcion._texto = texto
            opcion._indice = item.get("indice")
            valid_options.append(opcion)
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
            selected_text = self._get_selected_text_for_select(select_element)
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
                selected_text = self._get_selected_text_for_select(select_element)
                if selected_text:
                    if not self._is_placeholder_text(selected_text):
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
                    random_option = self._elegir_al_azar(select_id, field_name, valid_options)
                    _idx = getattr(random_option, "_indice", None)
                    if _idx is None:
                        _idx = select.options.index(random_option)
                    self._elegir_opcion_por_indice(select, select_element, _idx)
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
                textos = self._textos_de_opciones(select_element)
                chosen_idx = None
                # 1) exacto tal cual (texto visible idéntico)
                for i, t in enumerate(textos):
                    if not self._is_placeholder_text(t) and t == desired_plain:
                        chosen_idx = i
                        break
                # 2) exacto normalizado (ignora nbsp/acentos/espacios)
                if chosen_idx is None and norm_desired:
                    for i, t in enumerate(textos):
                        if not self._is_placeholder_text(t) and self._normalize_text(t) == norm_desired:
                            chosen_idx = i
                            break
                # 3) contiene normalizado (último recurso tolerante)
                if chosen_idx is None and norm_desired:
                    for i, t in enumerate(textos):
                        if not self._is_placeholder_text(t) and norm_desired in self._normalize_text(t):
                            chosen_idx = i
                            break

                if chosen_idx is not None:
                    self._elegir_opcion_por_indice(select, select_element, chosen_idx)
                    self.driver.execute_script(
                        "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));"
                        "arguments[0].dispatchEvent(new Event('input', { bubbles: true }));",
                        select_element
                    )
                    selected_text = (textos[chosen_idx] or option_text).strip()
                    print(f" {field_name} seleccionado: {selected_text}")
                    self._record_field_value(select_id, selected_text)
                    # Lo del Excel manda; igual cuenta como usado hoy, así el azar lo evita.
                    from osocio.utils.rotacion_dealers import tipo_de_campo
                    self._rotacion_anotar(tipo_de_campo(select_id, field_name), select_id, selected_text, al_azar=False)
                    return True

                print(f"❌ [DEBUG-FILL] No se encontró opción para {field_name}: '{option_text}'.")
                options_now = self._get_valid_select_options(select_element)
                print(f"   💡 Opciones disponibles en dropdown '{field_name}': {[t for t in textos if t]}")
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
                random_option = self._elegir_al_azar(select_id, field_name, valid_options)
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
