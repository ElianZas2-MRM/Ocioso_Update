"""Formularios 2.0 de AEM, que tienen su propia estructura de campos y de thank-you page.

Metodos extraidos de BaseFormFiller sin tocarles una linea. Siguen siendo parte de
la misma clase via herencia, asi que pueden llamar a sus hermanos y usar el mismo self.
"""

from urllib.parse import unquote, urlsplit


def ruta_aem_de_url(url):
    """Ruta del componente AEM de una URL de form (la parte desde el path, sin dominio,
    parámetros ni selectores), o "" si la URL no es de un componente AEM.

    https://www.chevrolet.com.br/request-proposal/jcr:content/root/.../formcontainer.iframe.br?dataRef=
    → /request-proposal/jcr:content/root/.../formcontainer
    """
    path = unquote(urlsplit(str(url or "").strip()).path).rstrip("/")
    if "/jcr:content/" not in path:
        return ""
    carpeta, _, ultimo = path.rpartition("/")
    return f"{carpeta}/{ultimo.split('.')[0]}"


def rutas_aem_coinciden(encontrada, esperada):
    """¿Es el mismo componente AEM? Se compara por el FINAL de la ruta: en la página el
    componente figura con la ruta interna completa (/content/chevrolet/sa/br/pt/index/
    request-proposal/jcr:content/...) y en el Excel con la pública (/request-proposal/
    jcr:content/...). Como la ruta del Excel incluye la página, no matchea el mismo
    componente de otra página."""
    a, b = ruta_aem_de_url(encontrada).lower(), ruta_aem_de_url(esperada).lower()
    return bool(a and b) and a.endswith(b)


class FormulariosAEMMixin:
    """Se usa solo como mixin de BaseFormFiller."""

    def _aem_componentes_en_pagina(self):
        """Rutas de los formcontainer AEM embebidos en la página actual (sin iframe). AEM
        las deja en el JSON de un script inline: "aemFormComponentPath": "\\/content\\/…"."""
        try:
            return self.driver.execute_script("""
                var out = [], re = /"aemFormComponentPath"\\s*:\\s*"([^"]+)"/g, m;
                var scripts = document.getElementsByTagName('script');
                for (var i = 0; i < scripts.length; i++) {
                    var t = scripts[i].text || '';
                    while ((m = re.exec(t)) !== null) {
                        try { out.push(JSON.parse('"' + m[1] + '"')); } catch (e) { out.push(m[1]); }
                    }
                }
                return out;
            """) or []
        except Exception:
            return []

    def _aem_form_embebido(self, expected_form_url):
        """Ruta del formcontainer de la página que es el form del Excel, o "". Para los
        forms AEM cuya URL del Excel es la versión iframe del componente (…/formcontainer
        .iframe.br) pero que la landing muestra embebidos, sin iframe."""
        if not ruta_aem_de_url(expected_form_url):
            return ""
        for ruta in self._aem_componentes_en_pagina():
            if rutas_aem_coinciden(ruta, expected_form_url):
                return ruta
        return ""

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
                "firstname": fd.get("firstname") or "Prueba",
                "lastname": fd.get("lastname") or "Prueba"
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
