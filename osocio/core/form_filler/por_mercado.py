"""Lo que solo aplica a algunos mercados: documentos de Brasil y Libro de Reclamaciones de Peru.

Metodos extraidos de BaseFormFiller sin tocarles una linea. Siguen siendo parte de
la misma clase via herencia, asi que pueden llamar a sus hermanos y usar el mismo self.
"""
from selenium.webdriver.common.by import By
import random
import time


class ReglasPorMercadoMixin:
    """Se usa solo como mixin de BaseFormFiller."""

    def _has_libro_reclamaciones_fields(self):
        """True si el DOM actual tiene los campos propios del Libro de Reclamaciones."""
        try:
            return all(self._element_exists_by_id(fid) for fid in self._LIBRO_RECLAMACIONES_IDS)
        except Exception:
            return False

    def _is_libro_reclamaciones_form(self, landing_url):
        """True si el form actual es un Libro de Reclamaciones.

        Se mira la URL (landing Y form: en las URLs sueltas el slug 'reclamos' viene en la
        del form, no en la landing) y, como respaldo, el DOM por los ids cc_*. El chequeo
        por DOM solo no alcanza si el form React todavía no montó cuando se evalúa.

        En estos forms NO se hace el clic de "enviar en vacío" que se usa en el resto para
        disparar las validaciones: se llena primero y recién ahí se envía.
        """
        _urls = (landing_url or "", getattr(self, "expected_form_url", "") or "")
        for _u in _urls:
            _u = _u.lower()
            if any(h in _u for h in self._LIBRO_RECLAMACIONES_URL_HINTS):
                return True
        return self._has_libro_reclamaciones_fields()

    def _fill_libro_reclamaciones_direct(self, form_data):
        """
        Fill directo para chevrolet.com.pe/libro-reclamaciones-virtual.
        IDs fijos: cc_name, cc_telephone, cc_ci, cc_email + dropdowns city/dealer aleatorios.
        """
        by_id = form_data.get("__by_id", {}) if isinstance(form_data, dict) else {}

        firstname = str(by_id.get("firstname") or by_id.get("name") or form_data.get("firstname") or "").strip()
        lastname  = str(by_id.get("lastname") or form_data.get("lastname") or "").strip()
        nombre    = f"{firstname} {lastname}".strip()
        phone     = str(by_id.get("telephone") or by_id.get("cellphone") or by_id.get("phone")
                        or form_data.get("phone") or "").strip()
        document  = str(by_id.get("ci") or by_id.get("document") or by_id.get("rut")
                        or form_data.get("document") or "").strip()
        email     = str(by_id.get("email") or form_data.get("email") or "").strip()

        def _fill_by_id(field_id, value):
            if not value:
                return False
            try:
                els = self.driver.find_elements(By.ID, field_id)
                for el in els:
                    if el.is_displayed() and el.is_enabled():
                        self._scroll_element_into_view(el)
                        try:
                            el.clear()
                            el.send_keys(value)
                        except Exception:
                            self.driver.execute_script(
                                "arguments[0].value = arguments[1];"
                                "arguments[0].dispatchEvent(new Event('input',{bubbles:true}));"
                                "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
                                el, value
                            )
                        # Estos IDs no están en el mapping del país, así que sin registrarlos
                        # acá no llegan al Excel de resultado: el lead viajaba con domicilio,
                        # monto y detalle del reclamo sin que quedara constancia de qué se
                        # envió. Todo campo que la herramienta llena es evidencia.
                        self._record_field_value(field_id, value)
                        print(f"  ✓ libro-reclamaciones: '{field_id}' = '{value}'")
                        return True
            except Exception as e:
                print(f"  ✗ libro-reclamaciones: error llenando '{field_id}': {e}")
            return False

        _fill_by_id("cc_name", nombre)
        _fill_by_id("cc_telephone", phone)
        
        if bool(self.config.get("solo_verificar_visual", False) or self.config.get("no_enviar_lead", False)):
            current_ss_number = getattr(self, 'ss_counter', 0)
            if self.screenshot_manager:
                self.screenshot_manager.take_form_screenshot(current_ss_number, "completado", full_page=True)
            return f"form_completado_{current_ss_number}.png"

        _fill_by_id("cc_ci", document)
        _fill_by_id("cc_email", email)

        # Resto de campos del reclamo: no están en el mapping del país y el form los valida
        # por JS (en el DOM figuran como required=false), así que hay que llenarlos acá.
        _fill_by_id("cc_address", "Av. Javier Prado Este 1234")
        _fill_by_id("cc_amount", "1000")
        _fill_by_id("cc_details", "Detalle de prueba automatizada del formulario.")
        _fill_by_id("cc_details_claim", "Detalle de prueba automatizada del formulario.")
        _fill_by_id("cc_order_claim", "Pedido de prueba automatizada.")

        # Ciudad aleatoria, luego esperar y seleccionar Concesionario aleatorio
        self.safe_select_option_if_visible("city", "", "Ciudad")
        time.sleep(0.5)
        if self._wait_for_dependent_dropdown_ready("dealer", parent_id="city"):
            self.safe_select_option_if_visible("dealer", "", "Concesionario")

        # Radios obligatorios (mayor de edad / bien contratado / reclamo-queja) y el
        # checkbox de términos del pie.
        try:
            self._handle_terms_checkboxes()
        except Exception as _e:
            print(f"  ⚠ libro-reclamaciones: radios/checkboxes — {_e}")

        current_ss_number = getattr(self, 'ss_counter', 0)
        if self.screenshot_manager:
            # full_page: el Libro de Reclamaciones es larguísimo y sin esto la captura salía
            # recortada al viewport (solo Pedido/VIN/Ciudad/Enviar, sin los datos personales).
            self._revalidar_campos_llenos()
            self.screenshot_manager.take_form_screenshot(current_ss_number, "completado",
                                                         full_page=True)
            print("Captura 2/3: Formulario completado")
        return f"form_completado_{current_ss_number}.png"

    def _verificar_cta_eletricos_br(self, landing_url, form_url):
        """Chequeo extra, sólo para el RAQ de Brasil (chevrolet.com.br/solicitar-contato).

        Esa pantalla ofrece dos cards: "Veículos a combustão" y "Veículos elétricos".
        El CTA "Formulário" de la card de eléctricos tiene que llevar a la landing de
        eléctricos, que es donde vive el form raq-eletricos. Se verifica que el link
        exista, que apunte a esa URL y que la URL responda 200.

        Debe llamarse ESTANDO DENTRO del iframe y ANTES de clickear #contact-by-form
        (ese click reemplaza la pantalla de las cards por el formulario).
        """
        pais = str(self.config.get("pais", "")).lower()
        blob = f"{landing_url or ''} {form_url or ''}".lower()
        if pais not in ("brasil", "brazil", "br"):
            return
        if "solicitar-contato" not in blob or "eletrico" in blob:
            return  # sólo la landing genérica; la de eléctricos no tiene estas cards

        try:
            href = self.driver.execute_script("""
                var sels = ["a.actions-link.actions-links-form[data-dtm='electric vehicles']",
                            "a.actions-links-form[data-dtm='electric vehicles']",
                            "a[data-dtm='electric vehicles'][href*='/eletrico/']"];
                for (var s = 0; s < sels.length; s++) {
                    var e = document.querySelector(sels[s]);
                    if (e) { return e.getAttribute('href') || ''; }
                }
                var links = document.querySelectorAll("a[href*='/eletrico/solicitar-contato']");
                return links.length ? (links[0].getAttribute('href') || '') : '';
            """) or ""
        except Exception as e:
            self._cta_eletricos = f"FAIL — no se pudo leer el CTA ({e})"
            print(f"CTA eléctricos: error leyendo el link: {e}")
            return

        href = str(href).strip()
        if not href:
            self._cta_eletricos = "FAIL — no se encontró el CTA 'Formulário' de Veículos elétricos"
            print("CTA eléctricos: NO se encontró el link")
            return

        from urllib.parse import urljoin
        href_abs = urljoin(self._CTA_ELETRICOS_URL, href)
        if href_abs.rstrip("/").lower() != self._CTA_ELETRICOS_URL.rstrip("/").lower():
            self._cta_eletricos = f"FAIL — el CTA apunta a {href_abs}"
            print(f"CTA eléctricos: apunta a {href_abs} (se esperaba {self._CTA_ELETRICOS_URL})")
            return

        try:
            from osocio.utils.url_status import check_url_status
            estado = check_url_status(href_abs) or {}
            etiqueta = estado.get("label", "")
            code = estado.get("code")
        except Exception as e:
            self._cta_eletricos = f"FAIL — no se pudo verificar la URL ({e})"
            print(f"CTA eléctricos: error verificando la URL: {e}")
            return

        if code == 200:
            self._cta_eletricos = f"PASS — {href_abs} ({etiqueta or '200 OK'})"
            print(f"CTA eléctricos OK: {href_abs} → {etiqueta or '200 OK'}")
        else:
            self._cta_eletricos = f"FAIL — {href_abs} responde {etiqueta or code}"
            print(f"CTA eléctricos: {href_abs} responde {etiqueta or code}")

    def _refill_brasil_doc_sendkeys(self, form_data) -> bool:
        """
        Re-ingresa los campos CPF/CNPJ/CEP carácter a carácter via send_keys,
        usando el MISMO valor del Excel. Solo para Brasil.
        Retorna True si re-llenó al menos un campo.
        """
        pais = str(self.config.get("pais", "")).lower()
        if pais not in ("brasil", "brazil", "br"):
            return False

        doc_keywords = ("cpf", "cnpj", "cep", "zip", "postal", "document", "ci")
        rellenados = 0
        for fc in (self.field_mapping or []):
            fid_raw = fc.get("id", "")
            fids    = fid_raw if isinstance(fid_raw, list) else [fid_raw]
            for fid in fids:
                if not fid or not any(k in str(fid).lower() for k in doc_keywords):
                    continue
                valor = self.get_form_value_str(form_data, fc.get("name", fid))
                if not valor:
                    continue
                # Normalizar: extraer dígitos y recuperar cero inicial comido por Excel numérico
                _fid_k = str(fid).lower()
                _min_k = 14 if "cnpj" in _fid_k else (8 if any(x in _fid_k for x in ("cep", "zip", "postal")) else 11)
                _digs = "".join(c for c in valor if c.isdigit())
                if _digs and len(_digs) == _min_k - 1:
                    _digs = _digs.zfill(_min_k)
                if _digs:
                    valor = _digs
                try:
                    els = self.driver.find_elements(By.ID, fid)
                    if not els or not els[0].is_displayed():
                        continue
                    el = els[0]
                    if not self._hard_clear_input(el):
                        print(f"  ⚠ No se pudo vaciar '{fid}' — se omite el re-ingreso "
                              f"para no duplicar el valor")
                        continue
                    self.driver.execute_script(
                        "arguments[0].dispatchEvent(new Event('input',{bubbles:true}));", el
                    )
                    time.sleep(0.05)
                    for char in str(valor):
                        el.send_keys(char)
                        time.sleep(0.005)
                    self.driver.execute_script(
                        "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));"
                        "arguments[0].dispatchEvent(new Event('blur',{bubbles:true}));",
                        el,
                    )
                    print(f"  ↺ Re-ingresado '{fid}' via send_keys (mismo valor): '{valor}'")
                    rellenados += 1
                except Exception as e:
                    print(f"  ✗ Error send_keys retry '{fid}': {e}")
        return rellenados > 0

    def _fetch_4devs(self, acao, extra_params=None, timeout=8):
        """Llama a la API de 4devs y retorna el texto de respuesta."""
        import urllib.request, urllib.parse
        data = {"acao": acao}
        if extra_params:
            data.update(extra_params)
        body = urllib.parse.urlencode(data).encode()
        req = urllib.request.Request(
            "https://www.4devs.com.br/ferramentas_online.php",
            data=body,
            headers={"Content-Type": "application/x-www-form-urlencoded",
                     "User-Agent": "Mozilla/5.0"},
        )
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.read().decode("utf-8").strip()

    def _generate_valid_cpf(self):
        """Fallback: genera CPF con dígitos verificadores válidos."""
        while True:
            d = [random.randint(0, 9) for _ in range(9)]
            if len(set(d)) == 1:
                continue
            r1 = sum(v * (10 - i) for i, v in enumerate(d)) % 11
            c1 = 0 if r1 < 2 else 11 - r1
            d2 = d + [c1]
            r2 = sum(v * (11 - i) for i, v in enumerate(d2)) % 11
            c2 = 0 if r2 < 2 else 11 - r2
            return "".join(map(str, d + [c1, c2]))

    def _generate_valid_cnpj(self):
        """Fallback: genera CNPJ con dígitos verificadores válidos (con puntuación)."""
        b = [random.randint(0, 9) for _ in range(8)] + [0, 0, 0, 1]
        w1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
        r1 = sum(v * w for v, w in zip(b, w1)) % 11
        c1 = 0 if r1 < 2 else 11 - r1
        w2 = [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
        r2 = sum(v * w for v, w in zip(b + [c1], w2)) % 11
        c2 = 0 if r2 < 2 else 11 - r2
        digits = "".join(map(str, b + [c1, c2]))
        return f"{digits[:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:]}"

    def _generate_brazil_document(self, field_id):
        """Genera CPF (11 díg), CNPJ (14 díg) o CEP (8 díg) via API 4devs con fallback local."""
        import re as _re
        fid = field_id.lower()
        is_cep  = "cep" in fid or "zip" in fid or "postal" in fid
        is_cnpj = "cnpj" in fid

        if is_cep:
            html = self._fetch_4devs("gerar_cep", {
                "estado": "", "cidade": "São Paulo",
                "bairro": "", "tipo_cep": "residencial",
            })
            m = _re.search(r'(\d{5}-\d{3})', html)
            if m:
                val = m.group(1).replace("-", "")
                if len(val) >= 8:
                    return val
            return random.choice(self._VALID_CEPS)

        if is_cnpj:
            raw = self._fetch_4devs("gerar_cnpj", {"pontuacao": "S"})
            digits = "".join(c for c in (raw or "") if c.isdigit())
            if len(digits) >= 14:
                return raw
            return self._generate_valid_cnpj()

        # CPF — solo dígitos, mínimo 11
        raw = self._fetch_4devs("gerar_cpf", {"pontuacao": "N"})
        digits = "".join(c for c in (raw or "") if c.isdigit())
        if len(digits) >= 11:
            return digits
        return self._generate_valid_cpf()

    def _sanitize_peru_document(self, doc_type_value, raw_value):
        """Corrige el número de documento según las reglas de Perú."""
        dt = (doc_type_value or "").lower()
        if "dni" in dt:
            required_len, no_leading_zero = 8, True
        elif "ruc" in dt:
            required_len, no_leading_zero = 11, False
        elif "pasaporte" in dt:
            # El pasaporte es ALFANUMÉRICO: se conservan las letras del valor original y se
            # completa con caracteres alfanuméricos, no con dígitos.
            alnum = "".join(c for c in str(raw_value or "") if c.isalnum()).upper()
            _pool = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
            while len(alnum) < 12:
                alnum += random.choice(_pool)
            return alnum[:12]
        elif "carn" in dt or "extran" in dt:
            required_len, no_leading_zero = 12, False
        else:
            return raw_value

        digits = "".join(c for c in str(raw_value or "") if c.isdigit())
        while len(digits) < required_len:
            digits += str(random.randint(0, 9))
        digits = digits[:required_len]
        if no_leading_zero and digits[0] == "0":
            digits = str(random.randint(1, 9)) + digits[1:]
        return digits
