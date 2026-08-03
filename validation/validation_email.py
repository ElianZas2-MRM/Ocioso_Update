"""
validation_email.py — Envía el reporte de validación de campos por email.
Formatea el resumen (campos testeados, errores encontrados, resultados UI) y lo manda
usando la misma infraestructura de email de la interfaz (cola Outlook asíncrona).
"""
from datetime import datetime


def send_validation_report_email(excel_path, summary, recipient=None):
    """Envia el reporte de validacion usando la infraestructura de mails existente."""
    from interface.helpers_interface import (
        _encolar_email,
        cargar_config_global,
        obtener_email_destinatario,
        _build_url_table_html,
        _FIRMA_HTML,
    )

    config = cargar_config_global()
    if not bool(config.get("enviar_mail", False)):
        return True, "El envio de mails esta deshabilitado en la configuracion global."

    destinatario = (recipient or obtener_email_destinatario() or "").strip()
    if not destinatario:
        return False, "No hay destinatario configurado."

    summary = summary or {}
    fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")

    errores = int(summary.get("errors", 0))
    ui_errores = int(summary.get("ui_errors", 0))
    urls_error = int(summary.get("urls_error", 0))
    urls_ok = int(summary.get("urls_ok", 0))
    # Veredicto global: mismo criterio que las otras pestañas, para poder filtrar el asunto
    # por [PASS] / [FAILED] sin abrir el mail.
    hay_fallas = bool(errores or ui_errores or urls_error)
    resultado_global = "FAILED" if hay_fallas else "PASS"

    asunto = (f"[{resultado_global}] Osocio — Validacion de campos {fecha_actual} — "
              f"{urls_ok} URL(s) OK / {urls_error} con error")

    icono = "❌" if hay_fallas else "✅"
    cuerpo = (
        f"{icono} RESULTADO GLOBAL: {resultado_global}\n"
        f"Fecha: {fecha_actual}\n\n"
        f"URLs validadas OK: {urls_ok}\n"
        f"URLs con error: {urls_error}\n"
        f"Campos validados: {summary.get('fields', 0)}\n"
        f"Caracteres evaluados: {summary.get('characters', 0)}\n"
        f"OK: {summary.get('ok', 0)}\n"
        f"Errores: {errores}\n"
        f"Tests UI de error: {summary.get('ui_error_tests', 0)}\n"
        f"UI OK: {summary.get('ui_ok', 0)}\n"
        f"UI Errores: {ui_errores}\n"
        f"Regex final OK: {'SI' if summary.get('regex_ok') else 'NO'}\n"
    )

    # Tabla landing / form por URL, con el motivo del fallo cuando lo hay. Se reusa el mismo
    # builder que Envio de Leads para que los mails se lean igual en todas las pestañas.
    detalles = summary.get("url_details") or []
    items = [{
        "linea": f"{d.get('pais', '')} {d.get('browser', '')}/{d.get('viewport', '')}".strip(),
        "url": d.get("landing", ""),
        "url_secure": d.get("form", ""),
        "ok": bool(d.get("ok")),
        "error": d.get("error", ""),
    } for d in detalles]
    html_extra = _build_url_table_html(items) + _FIRMA_HTML

    _encolar_email(destinatario, asunto, cuerpo, [excel_path], html_extra=html_extra)
    return True, f"Reporte encolado para {destinatario}."
