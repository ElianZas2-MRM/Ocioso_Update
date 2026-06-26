from datetime import datetime


def send_validation_report_email(excel_path, summary, recipient=None):
    """Envia el reporte de validacion usando la infraestructura de mails existente."""
    from interface.helpers_interface import _encolar_email, cargar_config_global, obtener_email_destinatario

    config = cargar_config_global()
    if not bool(config.get("enviar_mail", False)):
        return True, "El envio de mails esta deshabilitado en la configuracion global."

    destinatario = (recipient or obtener_email_destinatario() or "").strip()
    if not destinatario:
        return False, "No hay destinatario configurado."

    fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")
    asunto = f"Validacion de campos {fecha_actual}"
    cuerpo = (
        "Se adjunta el resultado de la validacion de campos.\n\n"
        f"Campos validados: {summary.get('fields', 0)}\n"
        f"Caracteres evaluados: {summary.get('characters', 0)}\n"
        f"OK: {summary.get('ok', 0)}\n"
        f"Errores: {summary.get('errors', 0)}\n"
        f"Tests UI de error: {summary.get('ui_error_tests', 0)}\n"
        f"UI OK: {summary.get('ui_ok', 0)}\n"
        f"UI Errores: {summary.get('ui_errors', 0)}\n"
        f"Regex final OK: {'SI' if summary.get('regex_ok') else 'NO'}\n"
    )

    _encolar_email(destinatario, asunto, cuerpo, [excel_path])
    return True, f"Reporte encolado para {destinatario}."