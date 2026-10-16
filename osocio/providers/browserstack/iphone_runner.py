"""
iphone_runner.py
================
El motor de los navegadores (BaseFormFiller), adaptado al iPhone de BrowserStack.

Safari en iPhone no deja entrar al iframe del form (otro dominio), así que el llenado
va por la capa nativa del dispositivo (ios_nativo.FormNativoIOS). Pero TODO lo demás es
el motor de la app, heredado tal cual:
  - lectura del Excel, nombres y carpetas de resultados (t1/t3, CDC/T3), colores,
    "Form coincide", estado de las URLs, motivo, resumen y fail_rows (Reintentar fallidos)
  - el valor de cada campo: _resolve_field_value (Excel), '-' para omitir, IDs dinámicos
    sólo si la celda está vacía, documentos de Brasil normalizados (nunca generados),
    documento de Perú según el tipo elegido, campos no mapeados (obligatorio -> valor
    plausible; opcional -> 50%)
  - el reintento por error de event_id del bucle principal

Lo único que se reemplaza son los métodos que tocan el DOM del iframe, y sólo cuando el
form está en un iframe: si está embebido en la página no hay bloqueo, y llena el motor
normal por DOM, igual que en cualquier navegador.
"""

import random
import time
from typing import Dict, Optional

from osocio.core.campos_inferencia import ids_para_etiqueta
from osocio.core.generic_country_base import GenericCountryBase
from osocio.providers.browserstack.ios_nativo import OMITIR, FormNativoIOS


class FormFillerIphone(GenericCountryBase):

    _nativo: Optional[FormNativoIOS] = None

    def begin_row_tracking(self):
        super().begin_row_tracking()
        self._nativo = None  # cada fila es otra landing: el iframe se vuelve a buscar

    # ── entrar al form: quedarse en la landing ───────────────────────────────
    def _entrar_al_iframe(self, iframe):
        """No se entra (Safari lo bloquea): el form se maneja desde la capa nativa."""
        self._nativo = FormNativoIOS(self.driver, iframe, log=print)

    def wait_for_form_ready_in_iframe(self, timeout=3):
        return True if self._nativo else super().wait_for_form_ready_in_iframe(timeout)

    def _resolve_submit_button(self, wait_seconds=3):
        # Con el form en iframe no hay envío vacío inicial: su único fin es la captura de
        # errores, y acá no hay capturas (la evidencia es el video).
        return (None, None) if self._nativo else super()._resolve_submit_button(wait_seconds)

    def _detectar_checkboxes_con_error(self):
        return None if self._nativo else super()._detectar_checkboxes_con_error()

    def _verificar_cta_eletricos_br(self, landing_url, form_url):
        return None if self._nativo else super()._verificar_cta_eletricos_br(landing_url, form_url)

    def _maybe_click_raq_cta(self, *urls):
        """RAQ de Brasil: antes del form hay una pantalla con cards y el CTA "Solicitar
        Contato" (#contact-by-form) que hay que tocar. En los navegadores se clickea por DOM
        adentro del iframe; en el iPhone se toca por la capa nativa, con las mismas
        condiciones (Brasil + RAQ, no eléctricos)."""
        if not self._nativo:
            return super()._maybe_click_raq_cta(*urls)
        blob = " ".join((u or "").lower() for u in urls)
        if str(self.config.get("pais", "")).lower() not in ("brasil", "brazil", "br"):
            return
        if "raq" not in blob or "eletricos" in blob:
            return
        if not ("solicitar-contato" in blob or "gm_forms/raq" in blob or "raq-revamp" in blob):
            return
        if self._nativo.tocar_texto("Solicitar Contato"):
            print("Tocado 'Solicitar Contato' (RAQ Brasil)")
        else:
            print("⚠ No se encontró el CTA 'Solicitar Contato' del RAQ de Brasil")

    def _rellenar_si_form_se_reseteo(self, form_data, es_libro_reclamaciones=False):
        if self._nativo:
            return None
        return super()._rellenar_si_form_se_reseteo(form_data, es_libro_reclamaciones)

    def _refill_brasil_doc_sendkeys(self, form_data) -> bool:
        return False if self._nativo else super()._refill_brasil_doc_sendkeys(form_data)

    # ── llenado ──────────────────────────────────────────────────────────────
    def fill_form_fields(self, form_data):
        """Llena el form dentro de la landing con los valores que decide la app."""
        if not self._nativo:
            return super().fill_form_fields(form_data)  # form sin iframe: DOM normal
        self._current_form_data = form_data
        self._snapshot_ios: Dict[str, Dict] = {}
        self._tipo_doc_ios = ""
        self._nativo.llenar(lambda tipo, etiqueta: self._valor_para_etiqueta(form_data, tipo, etiqueta),
                      al_llenar=self._registrar_campo_ios)
        return self.screenshot_manager.fname("form", "completado", 1)

    def _config_para_ids(self, ids):
        """El field_config del mapping del país que corresponde a esos ids (con alias)."""
        from osocio.utils.field_id_aliases import alias_ids_for
        buscados = set()
        for i in ids:
            buscados.add(i)
            buscados.update(alias_ids_for(i))
        for fc in self.field_mapping or []:
            fids = fc.get("id")
            fids = fids if isinstance(fids, list) else [fids]
            if any(str(f or "") in buscados for f in fids):
                return fc
        return None

    def _id_principal(self, ids, fc):
        if fc is not None:
            fid = fc.get("id")
            return str(fid[0] if isinstance(fid, list) else fid)
        return ids[0] if ids else ""

    def _valor_para_etiqueta(self, form_data, tipo, etiqueta):
        """El mismo camino de decisión que fill_form_fields, partiendo de la etiqueta."""
        if tipo == "radio":
            # Una opción al azar (la capa nativa no ve el name del grupo, así que la
            # columna del Excel con el name todavía no aplica en iPhone; ver README).
            return None
        ids = ids_para_etiqueta(etiqueta)
        fc = self._config_para_ids(ids) if ids else None
        fid = self._id_principal(ids, fc) or etiqueta
        valor = self._resolve_field_value(form_data, fc) if fc is not None else ""
        valor = "" if valor is None else str(valor).strip()

        if self._pide_omitir_campo(fid, valor):
            return OMITIR

        if tipo == "select":
            # Celda vacía o placeholder -> una opción real al azar (lo resuelve el motor nativo).
            return None if (not valor or self._is_placeholder_text(valor)) else valor

        # Texto: un ID dinámico sólo si la celda vino vacía.
        if not valor:
            ids_din = self._cargar_ids_dinamicos()
            for cand in [fid, *ids]:
                if self._usar_valor_fijo(cand, valor, ids_din):
                    valores = self._resolve_dynamic_id_values(ids_din[cand])
                    if valores:
                        valor = str(valores[0])
                        break
        # Brasil: el documento sale del Excel, normalizado (nunca generado).
        if str(self.config.get("pais", "")).lower() in ("brasil", "brazil", "br"):
            tipo_doc = self._tipo_documento_brasil(*(list(ids) + [fid]))
            if tipo_doc:
                valor = self._normalizar_documento_brasil(tipo_doc, valor)
        # Perú: el documento se ajusta al tipo elegido antes en el form.
        if fid in ("ci", "document") and self._tipo_doc_ios and \
                str(self.config.get("pais", "")).lower() in ("peru", "pe"):
            valor = self._sanitize_peru_document(self._tipo_doc_ios, valor)

        if valor:
            return valor
        if fc is not None:
            return None  # campo mapeado sin dato: queda vacío, como en los navegadores
        # Campo que el mapping no conoce: obligatorio -> valor plausible; opcional -> 50%,
        # decidido una vez por campo y por fila (README, "Campos que no están en el Excel").
        if "*" not in etiqueta:
            clave = f"ios::{etiqueta}"
            if clave not in self._decision_opcionales:
                self._decision_opcionales[clave] = random.random() < 0.5
            if not self._decision_opcionales[clave]:
                return None
        from osocio.utils.data_generator import valor_plausible_por_nombre
        return valor_plausible_por_nombre(etiqueta, etiqueta) or "Prueba"

    def _registrar_campo_ios(self, reg):
        """Lo que quedó en cada campo: tracking del Excel y foto para 'Datos vs Excel'."""
        quedo = reg.get("quedo") or ""
        if quedo.startswith(("(", "no ", "sin ", "ERROR")) or reg.get("tipo") in ("checkbox", "radio"):
            return
        ids = ids_para_etiqueta(reg.get("etiqueta", ""))
        fc = self._config_para_ids(ids) if ids else None
        fid = self._id_principal(ids, fc) or reg.get("etiqueta", "")
        self._record_field_value(fid, quedo)
        if "document-type" in ids:
            self._tipo_doc_ios = quedo
        entrada = {"id": fid, "realId": fid, "name": fid, "value": quedo, "text": quedo}
        for clave in {fid, *ids}:
            self._snapshot_ios[clave] = entrada

    def _snapshot_form_state(self):
        """'Datos vs Excel' compara contra lo que MUESTRA la pantalla del iPhone."""
        if not self._nativo:
            return super()._snapshot_form_state()
        return dict(getattr(self, "_snapshot_ios", {}) or {})

    # ── envío ────────────────────────────────────────────────────────────────
    def submit_and_verify_form(self, current_ss_number, expected_form_url):
        if bool(self.config.get("solo_verificar_visual", False)
                or self.config.get("no_enviar_lead", False)
                or self.config.get("chequeo_campos", False)):
            return "PASS (Verificación visual)", "Verificación visual"
        nativo = self._nativo
        if nativo is None:
            return super().submit_and_verify_form(current_ss_number, expected_form_url)
        try:
            self._sync_tracked_with_dom_before_submit(getattr(self, "_current_form_data", {}))
        except Exception as e:
            print(f"  ⚠ Datos vs Excel: {e}")
        if not nativo.tocar_enviar():
            return "Error al enviar: no se encontró el botón Enviar", None
        self._marcar_checkboxes_que_pide_el_form(nativo)
        estado, texto = nativo.esperar_ty()
        if estado == "ty":
            print(f"✓ TY: {texto}")
            nativo.mostrar_ty()
            self._ty_cta = "-"
            return "Lead enviado correctamente", self.screenshot_manager.fname("form", "typage",
                                                                                current_ss_number)
        if estado == "event_id":
            print(f"  ⚠ Error event_id: {texto}")
            return "ERROR_EVENT_ID: Error de servidor al envío - formulario recargado automáticamente", None
        return "Formulario sigue visible (TY Page no detectada)", None

    def _marcar_checkboxes_que_pide_el_form(self, nativo) -> bool:
        """Checkboxes obligatorios por JS (la lógica de la app, #27, en el iPhone).

        En Chrome se detectan con el "enviar vacío" del principio. En el iPhone se hace al
        revés, para no arriesgar un envío vacío en un form sin obligatorios: se llena todo,
        se toca Enviar y, si el form frena con un aviso debajo de un checkbox o de su grupo,
        se marcan esos (todas las opciones del grupo, como en Chrome) y se reenvía una vez.
        Si nada frena el envío, es que no eran obligatorios. Pendiente: las columnas SI/NO
        del Excel para checkboxes opcionales (la capa nativa no ve el name del checkbox).
        """
        time.sleep(2)
        try:
            rechazados = nativo.checkboxes_rechazados()
        except Exception as e:
            print(f"  ⚠ No se pudo revisar si el form pidió checkboxes: {e}")
            return False
        if not rechazados:
            return False
        print("  ↺ El form pide marcar: " + ", ".join(
            f"{c.get('grupo') + ' > ' if c.get('grupo') else ''}{c['etiqueta'] or '?'}" for c in rechazados))
        for c in rechazados:
            print(f"    · checkbox {c['etiqueta'] or '?':28} -> {nativo.marcar(c)}")
        # Hay forms (Chile) que dejan "términos" deshabilitado hasta que todo lo demás es
        # válido: el toque del llenado no entró. Ahora que el resto está, se vuelve a marcar.
        try:
            for c in nativo.checkboxes_pendientes():
                print(f"    · checkbox {c['etiqueta'] or '?':28} -> {nativo.marcar(c)}")
        except Exception as e:
            print(f"  ⚠ No se pudieron revisar los checkboxes de términos: {e}")
        return nativo.tocar_enviar()

    def _esperar_recarga_tras_event_id(self, use_iframe, expected_form_url, timeout=12):
        if self._nativo is None:
            return super()._esperar_recarga_tras_event_id(use_iframe, expected_form_url, timeout)
        return self._nativo.esperar_recarga(timeout)
