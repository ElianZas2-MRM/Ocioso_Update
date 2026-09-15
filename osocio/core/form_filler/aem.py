"""Formularios 2.0 de AEM, que tienen su propia estructura de campos y de thank-you page.

Metodos extraidos de BaseFormFiller sin tocarles una linea. Siguen siendo parte de
la misma clase via herencia, asi que pueden llamar a sus hermanos y usar el mismo self.
"""


class FormulariosAEMMixin:
    """Se usa solo como mixin de BaseFormFiller."""

    def _is_aem_adaptive_form(self):
        """Detecta si la página actual es un Adobe AEM Adaptive Form (Guide)."""
        from osocio.utils import aem_fill
        return aem_fill.is_aem_adaptive_form(self.driver)

    def _aem_guide_data_paths(self):
        """data-path de los <div id="guideContainer-rootPanel…"> de la página actual."""
        try:
            return self.driver.execute_script("""
                var out = [];
                var els = document.querySelectorAll("div[id^='guideContainer-rootPanel']");
                for (var i = 0; i < els.length; i++) {
                    var p = els[i].getAttribute('data-path');
                    if (p) { out.push(p); }
                }
                return out;
            """) or []
        except Exception:
            return []

    def _aem_form_presente(self):
        """True si el guideContainer de la página corresponde al content path del Excel."""
        base = getattr(self, "_aem_content_path", "") or ""
        if not base:
            return False
        return any(str(p).strip().startswith(base) for p in self._aem_guide_data_paths())

    def _aem_ty_presente(self):
        """True si el form AEM ya mostró su Thank You page.

        Se aceptan tres señales, de la más fuerte a la más específica:
          1. el navegador quedó en la página de gracias (…/obrigado.html);
          2. el guideContainer conmutó su data-path a …guideThankYouPage.html;
          3. el mensaje de confirmación aparece dentro de un bloque .adv-col
             (es lo que hace hoy Cadillac Brasil: no cambia de URL ni de data-path,
             reemplaza el form por el bloque de agradecimiento en la misma página).
        """
        base = getattr(self, "_aem_content_path", "") or ""
        if not base:
            return False

        try:
            if "/obrigado" in (self.driver.current_url or "").lower():
                return True
        except Exception:
            pass

        objetivo = base + self._AEM_TY_SUFFIX
        if any(str(p).strip() == objetivo for p in self._aem_guide_data_paths()):
            return True

        try:
            return bool(self.driver.execute_script("""
                var els = document.querySelectorAll("[class*='adv-col']");
                var buscado = arguments[0];
                for (var i = 0; i < els.length; i++) {
                    var t = (els[i].innerText || '').toLowerCase()
                              .normalize('NFD').replace(/[\\u0300-\\u036f]/g, '');
                    if (t.indexOf(buscado) !== -1) { return true; }
                }
                return false;
            """, self._AEM_TY_TEXTO))
        except Exception:
            return False

    def _fill_aem_by_semantic_id(self, form_data):
        """Llena un AEM Adaptive Form delegando en utils.aem_fill (fuente única)."""
        from osocio.utils import aem_fill
        pais = str(self.config.get("pais", "")).lower()
        is_brasil = pais in ("brasil", "brazil", "br")
        fd = dict(form_data) if isinstance(form_data, dict) else {}
        
        # Modo rellenado parcial
        if bool(self.config.get("solo_verificar_visual", False) or self.config.get("no_enviar_lead", False)):
            fd = {
                "firstname": fd.get("firstname") or "Test",
                "lastname": fd.get("lastname") or "User"
            }
        else:
            # Asegurar cpf/cnpj/cep desde __by_id (por si el normalizado no los trae)
            by_id = fd.get("__by_id", {}) if isinstance(fd.get("__by_id"), dict) else {}
        for _src, _dst in (("cpf", "cpf"), ("cnpj", "cnpj"), ("cep", "cep"),
                           ("zip", "cep"), ("postal", "cep"),
                           ("vin", "vin"), ("vin-code", "vin"), ("chassis", "vin")):
            if not fd.get(_dst) and by_id.get(_src):
                fd[_dst] = by_id.get(_src)
        return aem_fill.fill_aem_form(
            self.driver, fd, is_brasil,
            gen_doc=self._generate_brazil_document,
            record=self._record_field_value,
        )
