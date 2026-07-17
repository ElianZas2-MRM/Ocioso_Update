"""
helpers_interface.py — Utilidades de soporte para la interfaz gráfica.
Incluye: sistema de envío de emails via Outlook (cola asíncrona + callbacks de estado),
gestión de config global, lectura/escritura de Excels, IDs dinámicos y dependencias.
"""
import os
import sys
import ast
import builtins
import pandas as pd
import zipfile
from datetime import datetime
from glob import glob
from tkinter import messagebox
import re
from tkinter import ttk
import json
import time
import queue
import threading

from utils.fixed_field_mapping_store import (
    build_excel_columns_for_country,
    get_fixed_mapping_ids,
    infer_country_from_excel_filename,
)
from utils.popup_logger import log_runtime
from utils.paths import BASE_DIR, DATA_DIR, TEMPORALES_DIR, JSON_DIR

# === RUTAS ===
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(TEMPORALES_DIR, exist_ok=True)
os.makedirs(JSON_DIR, exist_ok=True)

GLOBAL_CONFIG_PATH = os.path.join(JSON_DIR, "config_global.json")
DEFAULT_EMAIL_DESTINATARIO = "ariel.melgratti@mrm.com"

# === COLA DE ENVÍO DE EMAILS ===
_cola_emails = queue.Queue()
_outlook_instance = None
_worker_thread = None
_email_ui_callback = None  # Callback registrado por la UI para mostrar estado de envío


def registrar_callback_ui_email(callback):
    """Registra un handler de la UI que recibe ('pending'|'success'|'error', msg)."""
    global _email_ui_callback
    _email_ui_callback = callback


def print(*args, **kwargs):
    """Centraliza trazas internas: silencioso para usuario final y persistente en runtime.log."""
    message = " ".join(str(arg) for arg in args).strip()
    if message:
        log_runtime(message, level="INFO")

    # Solo mostrar por consola si se habilita explícitamente para soporte/debug.
    if os.environ.get("FORM_AUTOMATION_VERBOSE", "").strip().lower() in {"1", "true", "yes", "on"}:
        target = kwargs.get("file", sys.stdout)
        encoding = getattr(target, "encoding", None) or "utf-8"
        try:
            return builtins.print(*args, **kwargs)
        except UnicodeEncodeError:
            safe_args = []
            for arg in args:
                text = str(arg)
                safe_args.append(text.encode(encoding, errors="replace").decode(encoding, errors="replace"))
            return builtins.print(*safe_args, **kwargs)

    return None


_PAIS_ABREV = {
    "Argentina": "AR", "Bolivia": "BO", "Brasil": "BR",
    "Chile": "CL", "Colombia": "CO", "Ecuador": "EC",
    "Paraguay": "PY", "Peru": "PE", "Uruguay": "UY",
}

def _abrev_paises(lista):
    return " ".join(_PAIS_ABREV.get(p, p[:2].upper()) for p in lista)


def _cuerpo_a_html(texto, html_extra=""):
    """Convierte cuerpo de texto plano a HTML con fuente Segoe UI Emoji para Outlook clásico."""
    import html as _html
    escaped = _html.escape(texto).replace('\n', '<br>\n')
    return (
        '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
        '<body style="font-family:\'Segoe UI Emoji\',\'Segoe UI\',Arial,sans-serif;'
        'font-size:13px;color:#222;line-height:1.6;">'
        f'{escaped}{html_extra}</body></html>'
    )


def _enviar_via_smtp(destinatarios, asunto, cuerpo, adjuntos):
    """Envía email via SMTP usando configuración en config_global.json."""
    import smtplib
    import ssl
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders

    cfg = cargar_config_global().get("smtp", {})
    host = cfg.get("host", "")
    port = int(cfg.get("port", 587))
    user = cfg.get("user", "")
    password = cfg.get("password", "")

    if not host or not user or not password:
        raise RuntimeError(
            "SMTP no configurado. Agregar en json/config_global.json la clave 'smtp' con: "
            "host, port, user, password."
        )

    msg = MIMEMultipart()
    msg["From"] = user
    msg["To"] = destinatarios if isinstance(destinatarios, str) else "; ".join(destinatarios)
    msg["Subject"] = asunto
    msg.attach(MIMEText(_cuerpo_a_html(cuerpo), "html", "utf-8"))

    if adjuntos:
        for archivo in adjuntos:
            if archivo and os.path.exists(archivo):
                with open(archivo, "rb") as f:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f'attachment; filename="{os.path.basename(archivo)}"')
                msg.attach(part)

    ctx = ssl.create_default_context()
    with smtplib.SMTP(host, port) as server:
        server.starttls(context=ctx)
        server.login(user, password)
        server.sendmail(user, msg["To"].split("; "), msg.as_string())


def _safe_print(msg):
    try:
        builtins.print(msg)
    except Exception:
        try:
            enc = sys.stdout.encoding or "utf-8"
            builtins.print(str(msg).encode(enc, errors="replace").decode(enc))
        except Exception:
            pass

def _worker_envio_emails():
    """Worker que mantiene Outlook abierto y envía emails desde la cola."""
    global _outlook_instance
    pythoncom = None
    com_initialized = False

    try:
        import win32com.client as win32
        import pythoncom
        pythoncom.CoInitialize()
        com_initialized = True
        try:
            _outlook_instance = win32.GetActiveObject("Outlook.Application")
        except Exception as _get_err:
            _safe_print(f"[WARN] GetActiveObject falló ({_get_err}), intentando Dispatch...")
            _outlook_instance = win32.Dispatch("Outlook.Application")

        while True:
            try:
                data = _cola_emails.get(timeout=5)
                if data is None:
                    break

                # Soporta 4-tupla (legacy), 5-tupla (callback) y 6-tupla (html_extra)
                if len(data) == 6:
                    destinatarios, asunto, cuerpo, adjuntos, _callback, _html_extra = data
                elif len(data) == 5:
                    destinatarios, asunto, cuerpo, adjuntos, _callback = data
                    _html_extra = ""
                else:
                    destinatarios, asunto, cuerpo, adjuntos = data
                    _callback = None
                    _html_extra = ""

                try:
                    mail = _outlook_instance.CreateItem(0)
                    mail.To = destinatarios if isinstance(destinatarios, str) else "; ".join(destinatarios)
                    mail.Subject = asunto
                    mail.HTMLBody = _cuerpo_a_html(cuerpo, html_extra=_html_extra)
                    if adjuntos:
                        for i, archivo in enumerate(adjuntos, 1):
                            if archivo and os.path.exists(archivo):
                                mail.Attachments.Add(os.path.abspath(archivo))
                            else:
                                _safe_print(f"      [WARN] [{i}] Adjunto no encontrado: {archivo}")
                    _to = mail.To
                    _subject = mail.Subject
                    mail.Send()
                    _safe_print(f"[SUCCESS] Email enviado (Outlook) a: {_to} | Asunto: {_subject}")
                    time.sleep(0.5)

                    if _callback:
                        try:
                            _callback(True, "")
                        except Exception:
                            pass
                    if _email_ui_callback:
                        try:
                            _email_ui_callback("success", "")
                        except Exception:
                            pass

                    if adjuntos:
                        for archivo in adjuntos:
                            try:
                                if archivo and archivo.endswith(".zip") and os.path.exists(archivo):
                                    os.remove(archivo)
                            except Exception:
                                pass

                except Exception as e:
                    import traceback
                    _safe_print(f"[ERROR] Error al enviar email via Outlook: {e}")
                    log_runtime(traceback.format_exc(), level="ERROR")
                    if _callback:
                        try:
                            _callback(False, str(e))
                        except Exception:
                            pass
                    if _email_ui_callback:
                        try:
                            _email_ui_callback("error", str(e))
                        except Exception:
                            pass

                finally:
                    _cola_emails.task_done()

            except queue.Empty:
                pass
            except Exception as e:
                _safe_print(f"[ERROR] Error en worker de email: {e}")
                log_runtime(str(e), level="ERROR")

    except Exception as e:
        _safe_print(f"[ERROR] ERROR CRÍTICO inicializando worker de email: {e}")
        log_runtime(str(e), level="ERROR")
        while True:
            try:
                _cola_emails.get_nowait()
                _cola_emails.task_done()
            except queue.Empty:
                break

    finally:
        if com_initialized and pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def _iniciar_worker_envio():
    """Inicia el thread worker si no está ya corriendo."""
    global _worker_thread
    
    if _worker_thread is None or not _worker_thread.is_alive():
        _worker_thread = threading.Thread(target=_worker_envio_emails, daemon=True)
        _worker_thread.start()


def _encolar_email(destinatarios, asunto, cuerpo, adjuntos=None, callback=None, html_extra=""):
    """Encola un email para ser enviado por el worker."""
    _iniciar_worker_envio()
    if _email_ui_callback:
        try:
            _email_ui_callback("pending", "")
        except Exception:
            pass
    _cola_emails.put((destinatarios, asunto, cuerpo, adjuntos, callback, html_extra))


def cargar_config_global():
    """Carga configuración global persistida (json/config_global.json)."""
    try:
        if os.path.exists(GLOBAL_CONFIG_PATH):
            with open(GLOBAL_CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        return {}
    except Exception as e:
        print(f"⚠️ No se pudo cargar config global: {e}")
        return {}


def guardar_config_global(config):
    """Guarda configuración global en json/config_global.json."""
    try:
        if not isinstance(config, dict):
            raise ValueError("config debe ser dict")
        os.makedirs(JSON_DIR, exist_ok=True)
        with open(GLOBAL_CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"⚠️ No se pudo guardar config global: {e}")
        return False


def obtener_email_destinatario():
    """Devuelve lista de emails destinatarios (soporta varios separados por coma)."""
    config = cargar_config_global()
    raw = (config.get("email_destinatario") or "").strip()
    if not raw:
        raw = DEFAULT_EMAIL_DESTINATARIO
    return [e.strip() for e in raw.replace(';', ',').split(',') if e.strip()]


# === IDs DINÁMICOS ===
DYNAMIC_IDS_PATH = os.path.join(JSON_DIR, "ids_dinamicos.json")


def _normalizar_paises_ids_dinamicos(raw_paises):
    if raw_paises is None:
        return []
    if isinstance(raw_paises, str):
        candidatos = [raw_paises]
    elif isinstance(raw_paises, (list, tuple, set)):
        candidatos = list(raw_paises)
    else:
        return []

    normalizados = []
    for pais in candidatos:
        texto = str(pais).strip()
        if texto and texto not in normalizados:
            normalizados.append(texto)
    return normalizados


def _normalizar_estructura_ids_dinamicos(data):
    """Convierte formato legacy a {'version': 2, 'entries': [...]}"""
    entries = []

    def _append_entry(raw_entry):
        if not isinstance(raw_entry, dict):
            return
        entry_id = str(raw_entry.get("id") or "").strip()
        if not entry_id:
            return

        valor = raw_entry.get("valor")
        if valor is None:
            valor = raw_entry.get("valores", raw_entry.get("value", raw_entry.get("values")))

        nombre_campo = str(
            raw_entry.get("nombre_campo")
            or raw_entry.get("nombre")
            or raw_entry.get("campo")
            or ""
        ).strip()
        paises = _normalizar_paises_ids_dinamicos(raw_entry.get("paises", raw_entry.get("countries")))

        new_entry = {
            "id": entry_id,
            "valor": valor,
            "paises": paises,
        }
        if nombre_campo:
            new_entry["nombre_campo"] = nombre_campo
        entries.append(new_entry)

    if isinstance(data, dict) and isinstance(data.get("entries"), list):
        for item in data.get("entries", []):
            _append_entry(item)
    elif isinstance(data, dict):
        for key, raw_value in data.items():
            if key in {"version", "entries"}:
                continue
            if not str(key).strip():
                continue

            if isinstance(raw_value, dict):
                raw_entry = dict(raw_value)
                raw_entry["id"] = str(key).strip()
            else:
                raw_entry = {
                    "id": str(key).strip(),
                    "valor": raw_value,
                    "paises": [],
                }
            _append_entry(raw_entry)

    return {"version": 2, "entries": entries}


def cargar_ids_dinamicos():
    """Carga IDs dinámicos persistidos en formato {'version': 2, 'entries': [...]}"""
    try:
        if os.path.exists(DYNAMIC_IDS_PATH):
            with open(DYNAMIC_IDS_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            return _normalizar_estructura_ids_dinamicos(data)
        return {"version": 2, "entries": []}
    except Exception as e:
        print(f"⚠️ No se pudo cargar ids_dinamicos: {e}")
        return {"version": 2, "entries": []}


def guardar_ids_dinamicos(data):
    """Guarda IDs dinámicos en formato {'version': 2, 'entries': [...]}"""
    try:
        if not isinstance(data, dict):
            raise ValueError("data debe ser dict")
        data_normalizada = _normalizar_estructura_ids_dinamicos(data)
        # Preserve the dependencies key if it exists in the original or current file
        if "dependencies" in data:
            data_normalizada["dependencies"] = data["dependencies"]
        elif os.path.exists(DYNAMIC_IDS_PATH):
            try:
                with open(DYNAMIC_IDS_PATH, "r", encoding="utf-8") as f:
                    existing = json.load(f)
                if "dependencies" in existing:
                    data_normalizada["dependencies"] = existing["dependencies"]
            except Exception:
                pass
        os.makedirs(JSON_DIR, exist_ok=True)
        with open(DYNAMIC_IDS_PATH, "w", encoding="utf-8") as f:
            json.dump(data_normalizada, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"⚠️ No se pudo guardar ids_dinamicos: {e}")
        return False


def cargar_dependencias():
    """Carga las dependencias padre→hijo almacenadas en ids_dinamicos.json."""
    try:
        if os.path.exists(DYNAMIC_IDS_PATH):
            with open(DYNAMIC_IDS_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            deps = data.get("dependencies", [])
            if isinstance(deps, list):
                return deps
        return []
    except Exception as e:
        print(f"⚠️ No se pudo cargar dependencias: {e}")
        return []


def guardar_dependencias(deps):
    """Persiste la lista de dependencias en ids_dinamicos.json (clave 'dependencies')."""
    try:
        data = {}
        if os.path.exists(DYNAMIC_IDS_PATH):
            with open(DYNAMIC_IDS_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
        data["dependencies"] = deps
        os.makedirs(JSON_DIR, exist_ok=True)
        with open(DYNAMIC_IDS_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"⚠️ No se pudo guardar dependencias: {e}")
        return False


def obtener_ids_mapeados_normales():
    """Devuelve el conjunto de IDs declarados en los field_mapping fijos efectivos."""
    return get_fixed_mapping_ids()

# === FUNCIONES AUXILIARES ===

def obtener_ultimo_archivo(prefijo, extension=".xlsx"):
    """Devuelve el último archivo que coincide con el patrón, según el número final."""
    archivos = glob(f"{prefijo}*{extension}")
    if not archivos:
        return None
    
    def extract_number(filename):
        match = re.search(r'(\d+)(?=\.\w+$)', filename)
        return int(match.group()) if match else 0
    
    try:
        archivos_ordenados = sorted(archivos, key=extract_number, reverse=True)
        return archivos_ordenados[0]
    except Exception:
        # Si falla el ordenamiento por números, usar modificación de archivo
        archivos_ordenados = sorted(archivos, key=os.path.getmtime, reverse=True)
        return archivos_ordenados[0]

def obtener_ultima_carpeta(prefijo):
    """Devuelve la última carpeta que coincide con el patrón, según el número final."""
    # Buscar de manera case-insensitive
    carpetas = [c for c in os.listdir() if os.path.isdir(c) and c.lower().startswith(prefijo.lower())]

    if not carpetas:
        return None
    
    def extract_number(foldername):
        match = re.search(r'(\d+)$', foldername)
        return int(match.group()) if match else 0
    
    try:
        carpetas_ordenadas = sorted(carpetas, key=extract_number, reverse=True)
        return carpetas_ordenadas[0]
    except Exception as e:
        # Si falla el ordenamiento por números, usar modificación de carpeta
        carpetas_ordenadas = sorted(carpetas, key=lambda x: os.path.getmtime(x) if os.path.exists(x) else 0, reverse=True)
        return carpetas_ordenadas[0]

def limpiar_mensaje_error(mensaje_error):
    """Limpia el mensaje de error eliminando el stacktrace y información técnica."""
    if not isinstance(mensaje_error, str):
        return str(mensaje_error)
    
    # Patrones comunes para dividir el error
    patrones_divisores = [
        "(Session info:",
        "Stacktrace:",
        "Call Stack:",
        "Build info:",
        "System info:",
        "Driver info:",
        "\nStack trace:",
        "\n\tat ",
    ]
    
    for patron in patrones_divisores:
        if patron in mensaje_error:
            mensaje_error = mensaje_error.split(patron)[0].strip()
    
    # Limpiar espacios extras y saltos de línea
    mensaje_error = re.sub(r'\s+', ' ', mensaje_error).strip()
    
    # Remover múltiples emojis de error al inicio
    mensaje_error = re.sub(r'^(❌\s*)+', '', mensaje_error).strip()
    
    return mensaje_error

def analizar_errores_excel(ruta_excel):
    """
    Analiza el Excel y reporta los errores indicando la URL y el mensaje de error.
    
    Retorna un diccionario con:
    - total: número total de filas
    - exitosos: número de filas sin error
    - con_errores: número de filas con error
    - detalles: lista de diccionarios con {url, error, linea}
    - mensaje: mensaje formateado para mostrar
    """
    if not os.path.exists(ruta_excel):
        return {
            'total': 0,
            'exitosos': 0,
            'con_errores': 0,
            'detalles': [],
            'mensaje': f"❌ No se encontró el archivo {ruta_excel}"
        }

    try:
        df = pd.read_excel(ruta_excel)

        # Ignorar filas completamente vacías para no inflar el conteo
        def fila_tiene_datos(fila):
            for valor in fila:
                if pd.notna(valor) and str(valor).strip().lower() not in ("", "nan"):
                    return True
            return False

        if not df.empty:
            df = df[df.apply(fila_tiene_datos, axis=1)]

        # Asegurar que las columnas clave existan
        columnas = [c.lower() for c in df.columns]
        col_url = next((c for c in df.columns if c.lower() == "url"), None)
        col_resultado = next((c for c in df.columns if c.lower() == "resultado"), None)
        col_form_url_esperada   = next((c for c in df.columns if c.lower() in ("formulario", "form url esperada")), None)
        col_form_url_encontrada = next((c for c in df.columns if c.lower() == "form url encontrada"), None)
        col_form_coincide       = next((c for c in df.columns if c.lower() == "form coincide"), None)
        col_link_issue          = next((c for c in df.columns if c.lower() == "link issue typ"), None)

        total_filas = len(df)

        # Si no existe la columna de resultados, no analizamos nada
        if not col_resultado:
            return {
                'total': total_filas,
                'exitosos': 0,
                'con_errores': 0,
                'no_procesados': total_filas,
                'detalles': [],
                'mensaje': f"ℹ️ No se encontró columna 'Resultado' en {os.path.basename(ruta_excel)}"
            }

        resultados = df[col_resultado].astype(str).fillna("").str.strip()
        procesados_mask = resultados != ""
        # Éxito REAL = el runner (local y LambdaTest) marcó "Lead enviado correctamente".
        # Cualquier otra fila procesada (TY Page no detectada, formulario sigue visible,
        # intentos fallidos, error de event id, error de servidor, etc.) cuenta como ERROR,
        # aunque el texto no contenga literalmente la palabra "error".
        _ok_mask = resultados.str.contains("Lead enviado correctamente", case=False, na=False)
        # "Campos sin completar" = quedaron campos sin valor asignado que no se pudieron
        # llenar: cuenta como error aunque el lead se haya enviado (el usuario debe ir a
        # ⚙ IDs Dinámicos a asignarles un valor).
        _sin_completar_mask = resultados.str.contains("Campos sin completar", case=False, na=False)
        errores_mask = procesados_mask & (~_ok_mask | _sin_completar_mask)

        errores = df[errores_mask]
        con_errores = int(errores_mask.sum())
        procesados = int(procesados_mask.sum())
        exitosos = procesados - con_errores
        no_procesados = total_filas - procesados

        exitosos_df = df[procesados_mask & ~errores_mask]
        detalles_ok = []
        for _i, _fila in exitosos_df.iterrows():
            _url = str(_fila[col_url]) if col_url and pd.notna(_fila[col_url]) else ""
            _url_sec = str(_fila[col_form_url_esperada]) if col_form_url_esperada and pd.notna(_fila.get(col_form_url_esperada, float('nan'))) else ""
            detalles_ok.append({'url': _url, 'url_secure': _url_sec, 'linea': _i + 2})

        # Construir detalles estructurados
        detalles = []
        detalles_texto = []
        
        for i, fila in errores.iterrows():
            url = str(fila[col_url]) if col_url and pd.notna(fila[col_url]) else ""
            url_sec = str(fila[col_form_url_esperada]) if col_form_url_esperada and pd.notna(fila.get(col_form_url_esperada, float('nan'))) else ""
            mensaje_error = str(fila[col_resultado]).strip()

            # Limpiar el mensaje de error
            mensaje_limpio = limpiar_mensaje_error(mensaje_error)

            linea_excel = i + 2  # +2 porque Excel empieza en 1 y la fila 1 son los encabezados

            detalles.append({
                'url': url,
                'url_secure': url_sec,
                'error': mensaje_limpio,
                'linea': linea_excel
            })
            detalles_texto.append(
                f"➡️ En la URL: {url}\n"
                f"   Línea: {linea_excel}\n"
                f"   ❌ {mensaje_limpio}"
            )

        # Detectar formularios no insertados (Form coincide == NO)
        detalles_form = []
        if col_form_coincide:
            _fc = df[col_form_coincide].astype(str).str.strip().str.upper()
            form_no_coincide_mask = _fc.isin(["NO", "FORMULARIO INCORRECTO", "FAIL"])
            for idx, fila in df[form_no_coincide_mask].iterrows():
                url_landing    = str(fila[col_url]) if col_url and pd.notna(fila[col_url]) else "(sin URL)"
                url_esperada   = str(fila[col_form_url_esperada]) if col_form_url_esperada and pd.notna(fila.get(col_form_url_esperada, float('nan'))) else "?"
                url_encontrada = str(fila[col_form_url_encontrada]) if col_form_url_encontrada and pd.notna(fila.get(col_form_url_encontrada, float('nan'))) else "ninguna"
                detalles_form.append({
                    'linea': idx + 2,
                    'url_landing': url_landing,
                    'url_esperada': url_esperada,
                    'url_encontrada': url_encontrada,
                })

        # Detectar campos opcionales que quedaron vacíos por no tener valor asignado
        # (aviso, no error: el usuario puede asignarles valor en ⚙ IDs Dinámicos)
        detalles_sin_valor = []
        _aviso_sin_valor_mask = procesados_mask & resultados.str.contains(
            "campos opcionales vacíos", case=False, na=False
        )
        for idx, fila in df[_aviso_sin_valor_mask].iterrows():
            _url_sv = str(fila[col_url]) if col_url and pd.notna(fila[col_url]) else "(sin URL)"
            _res_sv = str(fila[col_resultado])
            _m = re.search(r"campos opcionales vacíos \(sin valor asignado\):\s*(.+?)(?:\s+—|$)", _res_sv)
            detalles_sin_valor.append({
                'linea': idx + 2,
                'url': _url_sv,
                'campos': _m.group(1).strip() if _m else "(ver columna Resultado)",
            })

        # Detectar links raros en la TY (columna 'LINK ISSUE TYP' con valor distinto de '-')
        detalles_link_issue = []
        if col_link_issue:
            _li = df[col_link_issue].astype(str).str.strip()
            link_issue_mask = _li.notna() & ~_li.isin(["", "-", "nan", "None"])
            for idx, fila in df[link_issue_mask].iterrows():
                url_landing = str(fila[col_url]) if col_url and pd.notna(fila[col_url]) else "(sin URL)"
                detalles_link_issue.append({
                    'linea': idx + 2,
                    'url_landing': url_landing,
                    'detalle': str(fila[col_link_issue]).strip(),
                })

        if con_errores == 0:
            if no_procesados > 0:
                mensaje = (
                    f"⚠️ Se detectaron {no_procesados} fila(s) no procesadas en {os.path.basename(ruta_excel)}. "
                    f"Verifique la columna 'Resultado'."
                )
            else:
                mensaje = f"✅ Sin errores detectados en {os.path.basename(ruta_excel)}"

            return {
                'total': total_filas,
                'exitosos': exitosos,
                'con_errores': 0,
                'no_procesados': no_procesados,
                'detalles': [],
                'detalles_ok': detalles_ok,
                'detalles_form': detalles_form,
                'detalles_link_issue': detalles_link_issue,
                'detalles_sin_valor': detalles_sin_valor,
                'mensaje': mensaje
            }

        resumen = "\n\n".join(detalles_texto)

        nota_no_procesados = ""
        if no_procesados > 0:
            nota_no_procesados = f"\n\n⚠️ Hay {no_procesados} fila(s) no procesadas en el Excel."

        mensaje = (
            f"⚠️ Se han detectado errores en {os.path.basename(ruta_excel)}:\n\n"
            f"{resumen}"
            f"{nota_no_procesados}"
        )

        return {
            'total': total_filas,
            'exitosos': exitosos,
            'con_errores': con_errores,
            'no_procesados': no_procesados,
            'detalles': detalles,
            'detalles_ok': detalles_ok,
            'detalles_form': detalles_form,
            'detalles_link_issue': detalles_link_issue,
            'detalles_sin_valor': detalles_sin_valor,
            'mensaje': mensaje
        }

    except Exception as e:
        return {
            'total': 0,
            'exitosos': 0,
            'con_errores': 0,
            'detalles': [],
            'detalles_ok': [],
            'mensaje': f"❌ No se pudo analizar {ruta_excel}: {e}"
        }

def _sincronizar_excel_si_cambia(ruta, country_name):
    """
    Compara las columnas del Excel existente contra las esperadas para el país.
    Si hay diferencias, muestra un diálogo con el detalle y reorganiza el archivo
    manteniendo los datos en sus columnas correspondientes.
    Columnas nuevas quedan vacías; columnas eliminadas del mapping se descartan.
    Hace un backup silencioso en temporales/ antes de sobrescribir.
    Retorna True si se reorganizó, False si no hay cambios o se canceló.
    """
    import shutil

    try:
        expected_cols = build_excel_columns_for_country(country_name)
    except Exception:
        return False

    try:
        df_existing = pd.read_excel(ruta, dtype=str)
    except Exception:
        return False

    if "Mensajes Error" in df_existing.columns:
        backup_name = (
            f"{os.path.splitext(os.path.basename(ruta))[0]}_sin_mensajes_error_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        )
        try:
            shutil.copy2(ruta, os.path.join(TEMPORALES_DIR, backup_name))
        except Exception:
            pass
        df_existing = df_existing.drop(columns=["Mensajes Error"])
        try:
            df_existing.to_excel(ruta, index=False)
        except PermissionError:
            messagebox.showerror(
                "Archivo en uso",
                "No se pudo quitar la columna obsoleta 'Mensajes Error'.\n\n"
                "Cerrá el Excel en otra aplicación y volvé a intentar.",
            )
            return False
        print(f"Columna obsoleta 'Mensajes Error' eliminada de {os.path.basename(ruta)}")

    existing_cols = list(df_existing.columns)
    existing_set = set(existing_cols)
    expected_set = set(expected_cols)
    added = [c for c in expected_cols if c not in existing_set]
    removed = [c for c in existing_cols if c not in expected_set]
    print("====================================")
    print("EXISTING COLS:", existing_cols)
    print("EXPECTED COLS:", expected_cols)
    print("ADDED:", added)
    print("REMOVED:", removed)
    print("====================================")

    moved = []
    for col in expected_cols:
        if col in existing_set:
            old_pos = existing_cols.index(col)
            new_pos = expected_cols.index(col)
            if old_pos != new_pos:
                moved.append((col, old_pos + 1, new_pos + 1))  # 1-based para mostrar

    # Detectar renombres por índice (misma posición, distinto nombre):
    # útil cuando cambia la descripción/nombre de columna en IDs fijos.
    renamed_by_index = []
    fixed_columns = 2  # A=URL, B=Formulario
    upper_bound = min(len(existing_cols), len(expected_cols))
    for idx in range(fixed_columns, upper_bound):
        old_col = existing_cols[idx]
        new_col = expected_cols[idx]
        if old_col == new_col:
            continue
        if old_col not in expected_set and new_col not in existing_set:
            renamed_by_index.append((old_col, new_col, idx + 1))

    if not added and not removed and not moved and not renamed_by_index:
        return False  # Sin cambios

    partes = []
    if added:
        partes.append("+ Columnas nuevas (quedarán vacías):\n  " + ", ".join(added))
    if removed:
        partes.append("- Columnas eliminadas (sus datos se descartan):\n  " + ", ".join(removed))
    if moved:
        lineas = [f"  {col}: col {old} → col {new}" for col, old, new in moved]
        partes.append("↔ Columnas reordenadas (los datos se mueven con ellas):\n" + "\n".join(lineas))
    if renamed_by_index:
        lineas = [f"  col {pos}: {old} → {new}" for old, new, pos in renamed_by_index]
        partes.append("✎ Columnas renombradas (se conserva la información por índice):\n" + "\n".join(lineas))

    mensaje = (
        "Se detectaron diferencias entre el Excel y la estructura esperada:\n\n"
        + "\n\n".join(partes)
        + "\n\n¿Querés reorganizar el archivo automáticamente?\n"
        "Los datos existentes se conservarán en las columnas que coincidan."
    )

    if not messagebox.askyesno("Estructura de Excel modificada", mensaje):
        return False

    # Backup silencioso en temporales/
    backup_name = f"{os.path.splitext(os.path.basename(ruta))[0]}_backup_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    try:
        shutil.copy2(ruta, os.path.join(TEMPORALES_DIR, backup_name))
    except Exception:
        pass

    # Reconciliar: construir nuevo DataFrame con columnas esperadas.
    # 1) Si el nombre coincide, copiar por nombre (permite mover Email de col 5 a 8).
    # 2) Si no coincide, intentar conservar por índice para renombres en misma posición.
    n_rows = len(df_existing)
    df_new = pd.DataFrame("", index=range(n_rows), columns=expected_cols)
    consumed_existing_indices = set()

    # 1) Copia por nombre
    for col in expected_cols:
        if col in df_existing.columns:
            src_idx = existing_cols.index(col)
            df_new[col] = df_existing.iloc[:, src_idx].values
            consumed_existing_indices.add(src_idx)

    # 2) Fallback por índice (solo columnas de datos, no URL/Formulario)
    for idx, col in enumerate(expected_cols):
        if col in df_existing.columns:
            continue
        if idx < fixed_columns or idx >= len(existing_cols):
            continue

        old_col = existing_cols[idx]
        if idx in consumed_existing_indices:
            continue
        if old_col in expected_set:
            continue

        df_new[col] = df_existing.iloc[:, idx].values
        consumed_existing_indices.add(idx)

    try:
        df_new.to_excel(ruta, index=False)
        return True
    except PermissionError:
        messagebox.showerror(
            "Archivo en uso",
            "No se pudo guardar el archivo reorganizado porque está abierto en otro programa.\n\n"
            "Cerrá el Excel y volvé a intentarlo."
        )
        return False


def sincronizar_excels_de_pais(country_name):
    """Revisa y sincroniza los Excels existentes de un país dentro de data/."""
    country_name = str(country_name or "").strip()
    if not country_name:
        return []

    rutas_sincronizadas = []
    try:
        for nombre_archivo in sorted(os.listdir(DATA_DIR)):
            if not nombre_archivo.lower().endswith(".xlsx"):
                continue

            if infer_country_from_excel_filename(nombre_archivo) != country_name:
                continue

            ruta_excel = os.path.join(DATA_DIR, nombre_archivo)
            if _sincronizar_excel_si_cambia(ruta_excel, country_name):
                rutas_sincronizadas.append(ruta_excel)
    except Exception as e:
        print(f"⚠️ No se pudieron sincronizar Excels de {country_name}: {e}")

    return rutas_sincronizadas


def crear_excel_si_no_existe(nombre_archivo, mostrar_mensaje=True):
    """
    Crea el Excel en data/ si no existe, usando la estructura esperada del país.
    Retorna (ruta, creado).
    """
    if not nombre_archivo.endswith('.xlsx'):
        nombre_archivo = f"{nombre_archivo}.xlsx"

    ruta = os.path.join(DATA_DIR, nombre_archivo)
    if os.path.exists(ruta):
        return ruta, False

    country_name = infer_country_from_excel_filename(nombre_archivo)
    columnas = build_excel_columns_for_country(country_name) if country_name else [
        "URL", "Formulario", "Modelo", "Nombre", "Apellido",
        "Documento", "Celular", "Email", "Region", "Ciudad", "Concesionario", "Fecha de compra",
        "Tipo de documento (Perú)", "Patente", "Evento", "Chasis/Vin", "Código asesor", "Comentario"
    ]

    df = pd.DataFrame(columns=columnas)
    df.to_excel(ruta, index=False)

    if mostrar_mensaje:
        messagebox.showinfo("Archivo creado", f"Se creó automáticamente:\n{ruta}")

    return ruta, True


def abrir_excel(nombre_archivo):
    """Abre el Excel si existe; si no, crea uno nuevo con la estructura base en la carpeta data."""
    ruta, creado = crear_excel_si_no_existe(nombre_archivo, mostrar_mensaje=True)
    if not creado:
        # Archivo ya existe: verificar si las columnas cambiaron y sincronizar si el usuario acepta
        country_name = infer_country_from_excel_filename(nombre_archivo)
        if country_name:
            _sincronizar_excel_si_cambia(ruta, country_name)

    # Abrir archivo (ya existente o recién creado/sincronizado)
    os.startfile(ruta)

#palmail
def crear_zip_de_carpeta(carpeta):
    """Comprime una carpeta en uno o más ZIPs, cada uno con un tamaño máximo de ~18 MB comprimido."""
    if not os.path.exists(carpeta):
        print(f"❌ La carpeta no existe: {carpeta}")
        return None

    MAX_ZIP_SIZE = 18 * 1024 * 1024  # 18 MB en bytes (límite seguro para Outlook corporativo)
    zip_files = []

    # Obtener lista de archivos
    archivos = []
    for root_dir, _, files in os.walk(carpeta):
        for file in files:
            full_path = os.path.join(root_dir, file)
            archivos.append((full_path, os.path.relpath(full_path, carpeta)))
    
    if not archivos:
        print(f"❌ No se encontraron archivos en la carpeta: {carpeta}")
        return None

    # Ordenar por nombre
    archivos.sort(key=lambda x: x[1])

    # Trackear qué archivos van en cada ZIP
    # Guardar temporales fuera de data/ para facilitar limpieza manual si hiciera falta
    os.makedirs(TEMPORALES_DIR, exist_ok=True)
    zip_name_base = os.path.join(TEMPORALES_DIR, f"{os.path.basename(carpeta)}_{datetime.now():%Y%m%d_%H%M%S}")

    def open_zip(part_number):
        zip_path = f"{zip_name_base}_parte{part_number}.zip"
        return zip_path, zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED)

    def current_zip_size_bytes(zf):
        try:
            # Suma de tamaños comprimidos (aprox; no incluye headers)
            return sum(info.compress_size for info in zf.infolist())
        except Exception:
            return 0

    def close_zip_safely(zf):
        try:
            zf.close()
        except Exception:
            pass

    zip_count = 1
    zip_path, zipf = open_zip(zip_count)
    current_part_entries = []

    try:
        for full_path, relative_path in archivos:
            try:
                zipf.write(full_path, relative_path)
                current_part_entries.append((full_path, relative_path))
            except Exception as e:
                print(f"⚠️ No se pudo agregar a ZIP: {relative_path} ({e})")
                continue

            size_now = current_zip_size_bytes(zipf)
            if size_now <= MAX_ZIP_SIZE:
                continue

            # Si se excede, partimos el ZIP actual dejando el último archivo para la siguiente parte.
            if len(current_part_entries) > 1:
                overflow_entry = current_part_entries.pop()

                close_zip_safely(zipf)
                try:
                    if os.path.exists(zip_path):
                        os.remove(zip_path)
                except Exception:
                    pass

                zipf_rebuild = zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED)
                for prev_full, prev_rel in current_part_entries:
                    try:
                        zipf_rebuild.write(prev_full, prev_rel)
                    except Exception:
                        continue
                close_zip_safely(zipf_rebuild)

                if os.path.exists(zip_path) and os.path.getsize(zip_path) > 0:
                    zip_files.append(zip_path)

                zip_count += 1
                zip_path, zipf = open_zip(zip_count)
                current_part_entries = []

                try:
                    zipf.write(overflow_entry[0], overflow_entry[1])
                    current_part_entries.append(overflow_entry)
                except Exception as e:
                    print(f"⚠️ No se pudo agregar a ZIP nuevo: {overflow_entry[1]} ({e})")
                    continue

                # Si un único archivo supera el límite comprimido, se conserva igual en su propia parte.
                single_size = current_zip_size_bytes(zipf)
                if single_size > MAX_ZIP_SIZE:
                    close_zip_safely(zipf)
                    if os.path.exists(zip_path) and os.path.getsize(zip_path) > 0:
                        zip_files.append(zip_path)
                        final_size = os.path.getsize(zip_path)
                        print(
                            f"⚠️ ZIP parte {zip_count} supera el límite por archivo individual grande "
                            f"({final_size / (1024 * 1024):.2f} MB): {os.path.basename(zip_path)}"
                        )
                    zip_count += 1
                    zip_path, zipf = open_zip(zip_count)
                    current_part_entries = []
            else:
                # Caso de un único archivo que ya excede el límite.
                close_zip_safely(zipf)
                if os.path.exists(zip_path) and os.path.getsize(zip_path) > 0:
                    zip_files.append(zip_path)
                    final_size = os.path.getsize(zip_path)
                    print(
                        f"⚠️ ZIP parte {zip_count} supera el límite por archivo individual grande "
                        f"({final_size / (1024 * 1024):.2f} MB): {os.path.basename(zip_path)}"
                    )

                zip_count += 1
                zip_path, zipf = open_zip(zip_count)
                current_part_entries = []

        close_zip_safely(zipf)

        if os.path.exists(zip_path) and os.path.getsize(zip_path) > 0:
            zip_files.append(zip_path)
        else:
            try:
                if os.path.exists(zip_path):
                    os.remove(zip_path)
            except Exception:
                pass

        return zip_files if zip_files else None

    except Exception as e:
        print(f"❌ Error al crear ZIPs: {e}")
        close_zip_safely(zipf)
        try:
            if os.path.exists(zip_path) and os.path.getsize(zip_path) == 0:
                os.remove(zip_path)
        except Exception:
            pass
        return None

def _url_short(url):
    """Extrae solo el path de una URL para mostrar en resumen."""
    try:
        from urllib.parse import urlparse
        path = urlparse(url).path.rstrip('/')
        parts = [p for p in path.split('/') if p]
        return '/' + '/'.join(parts[-2:]) if parts else url
    except Exception:
        return url


def _build_url_table_html(items):
    """Construye un bloque HTML con tabla URL LANDING | URL FORM para el email.
    items: lista de {ok, url, url_secure, error, linea}
    Falls first, then passes.
    """
    import html as _h
    if not items:
        return ""

    # Fails first, then passes; within each group keep original order
    sorted_items = [i for i in items if not i['ok']] + [i for i in items if i['ok']]

    _TD = 'style="padding:7px 10px;border:1px solid #ddd;vertical-align:top"'
    _TD_IC = 'style="padding:7px 8px;border:1px solid #ddd;vertical-align:top;text-align:center;width:28px"'

    rows_html = ""
    for item in sorted_items:
        landing = _h.escape(item.get('url', '') or '—')
        form_url = _h.escape(item.get('url_secure', '') or '—')
        error   = _h.escape(item.get('error', ''))
        icon    = "✅" if item['ok'] else "❌"
        row_bg  = "#f0fff4" if item['ok'] else "#fff5f5"

        rows_html += (
            f'<tr style="background:{row_bg}">'
            f'<td {_TD_IC}>{icon}</td>'
            f'<td {_TD}>{landing}</td>'
            f'<td {_TD}>{form_url}</td>'
            '</tr>'
        )
        if not item['ok'] and error:
            rows_html += (
                f'<tr style="background:{row_bg}">'
                '<td style="border:1px solid #ddd;border-top:none"></td>'
                f'<td colspan="2" style="padding:3px 10px 8px 10px;border:1px solid #ddd;'
                f'border-top:none;color:#c0392b;font-size:12px">↳ {error}</td>'
                '</tr>'
            )

    _TH = 'style="text-align:left;padding:7px 10px;border:1px solid #ddd;background:#f0f0f0;color:#555;font-weight:bold"'
    return (
        '<hr style="border:none;border-top:1px solid #ddd;margin:18px 0">'
        '<table style="border-collapse:collapse;width:100%;font-size:13px">'
        '<thead><tr>'
        f'<th {_TH} style="width:28px;text-align:center"> </th>'
        f'<th {_TH}>URL LANDING</th>'
        f'<th {_TH}>URL FORM</th>'
        '</tr></thead>'
        f'<tbody>{rows_html}</tbody>'
        '</table>'
    )


def enviar_email_resultados(pais, excel_path, screenshots_dir, browser=None, viewport=None):
    """
    Envía un email con los resultados de la ejecución del formulario.

    Args:
        pais: Nombre del país (ej: "Argentina")
        excel_path: Ruta al archivo Excel de resultados
        screenshots_dir: Ruta a la carpeta de screenshots
        browser: Navegador usado (opcional, para mostrarlo en el email)
        viewport: Viewport usado (opcional, para mostrarlo en el email)
    """
    try:
        if not os.path.exists(excel_path):
            print(f"❌ No se encontró el archivo Excel: {excel_path}")
            return False

        if screenshots_dir and not os.path.exists(screenshots_dir):
            print(f"⚠️ No se encontró la carpeta de screenshots: {screenshots_dir}")

        errores = analizar_errores_excel(excel_path)
        exitosos = int(errores.get('exitosos', 0))
        con_errores = int(errores.get('con_errores', 0))
        no_procesados = int(errores.get('no_procesados', 0))
        total_filas = int(errores.get('total', 0))

        nav_label, vp_label = _label_navegador(browser or "", viewport or "")
        fecha_actual = datetime.now().strftime("%d/%m/%Y")

        resultado_global = "PASS" if con_errores == 0 else "FAILED"
        asunto = f"[{resultado_global}] Osocio {_PAIS_ABREV.get(pais, pais)} {fecha_actual} — {exitosos} OK / {con_errores} errores"

        encabezado = pais
        if browser:
            encabezado += f" — {nav_label} / {vp_label}"

        icono = "✅" if con_errores == 0 else "❌"
        if con_errores == 0:
            estado_texto = "OK — todos los formularios insertados y leads enviados correctamente"
        else:
            _hay_form_error = bool(errores.get('detalles_form'))
            _hay_envio_error = any(
                not d.get('error', '').startswith('[Error Form]')
                for d in (errores.get('detalles') or [])
            )
            if _hay_form_error and not _hay_envio_error:
                estado_texto = f"FORMULARIO NO INSERTO en {len(errores['detalles_form'])} fila(s)"
            elif _hay_form_error and _hay_envio_error:
                estado_texto = f"ERRORES MÚLTIPLES: formulario(s) no inserto + {con_errores - len(errores['detalles_form'])} error(es) de envío"
            else:
                estado_texto = f"LEAD NO ENVIADO en {con_errores} formulario(s)"

        cuerpo = f"{icono} {encabezado}\nFecha: {fecha_actual}\nEstado: {estado_texto}\n\n"

        _fail_urls = [_url_short(d.get('url', '')) for d in (errores.get('detalles') or [])]
        _pass_urls = [_url_short(d.get('url', '')) for d in (errores.get('detalles_ok') or [])]
        if _fail_urls:
            cuerpo += f"FAILED ({len(_fail_urls)}): {pais}\n"
        if _pass_urls:
            cuerpo += f"PASSED ({len(_pass_urls)}): {pais}\n"

        _link_issues = errores.get('detalles_link_issue') or []
        if _link_issues:
            cuerpo += f"\n⚠️ LINK ISSUE en la TY page ({len(_link_issues)} fila/s) — link raro fuera de un <a> normal:\n"
            for _li in _link_issues:
                cuerpo += f"  • Línea {_li['linea']} — {_url_short(_li.get('url_landing',''))}: {_li.get('detalle','')}\n"

        _sin_valor = errores.get('detalles_sin_valor') or []
        if _sin_valor:
            cuerpo += (f"\n⚠️ CAMPOS SIN VALOR ASIGNADO ({len(_sin_valor)} fila/s) — quedaron vacíos al enviar. "
                       f"Asignales un valor en Osocio → Envío de Leads → ⚙ IDs Dinámicos:\n")
            for _sv in _sin_valor:
                cuerpo += f"  • Línea {_sv['linea']} — {_url_short(_sv.get('url',''))}: {_sv.get('campos','')}\n"

        cuerpo += "\nSaludos,\nAutomación de Formularios"

        _ok_items = [{'linea': d['linea'], 'url': d.get('url',''), 'url_secure': d.get('url_secure',''), 'ok': True, 'error': ''} for d in (errores.get('detalles_ok') or [])]
        _err_items = [{'linea': d['linea'], 'url': d.get('url',''), 'url_secure': d.get('url_secure',''), 'ok': False, 'error': d.get('error', '')} for d in (errores.get('detalles') or [])]
        _url_table_html = _build_url_table_html(_ok_items + _err_items)

        config = cargar_config_global()
        enviar_mail = bool(config.get("enviar_mail", False))
        adjuntar_resultados = bool(config.get("adjuntar_resultados", True))
        adjuntar_screenshots = bool(config.get("adjuntar_screenshots", True))

        if not enviar_mail:
            print("📭 Envío de email deshabilitado por Configuración Global. Se omite el envío.")
            return True

        adjuntos = []

        if adjuntar_resultados:
            if os.path.getsize(excel_path) < 24 * 1024 * 1024:
                adjuntos.append(excel_path)
            else:
                print(f"⚠️ Excel excede 24 MB, no se adjuntará")
                cuerpo += f"\n⚠️ NOTA: El archivo Excel es muy grande y no se pudo adjuntar. Ubicación: {excel_path}"

        if adjuntar_screenshots and screenshots_dir and os.path.exists(screenshots_dir):
            zip_files = crear_zip_de_carpeta(screenshots_dir)
            if zip_files:
                adjuntos.extend(zip_files)
            else:
                print(f"⚠️ No se pudieron comprimir los screenshots")
                cuerpo += f"\n⚠️ NOTA: Los screenshots no se pudieron comprimir. Ubicación: {screenshots_dir}"

        destinatario = obtener_email_destinatario()

        try:
            _encolar_email(destinatario, asunto, cuerpo, adjuntos, html_extra=_url_table_html)
        except Exception as e:
            print(f"❌ Error encolando email: {e}")
            return False

        return True

    except Exception as e:
        print(f"❌ Error al enviar email de resultados: {e}")
        return False


def _label_navegador(navegador, viewport):
    """Convierte los IDs internos a nombres legibles para el email."""
    nav_map = {
        "chrome": "Chrome",
        "firefox": "Firefox",
        "edge": "Edge",
        "lambdatest_mac": "LambdaTest Mac",
        "lambdatest_android": "LambdaTest Android",
    }
    vp_map = {
        "desktop": "Desktop",
        "mobile": "Mobile",
        "mac": "Mac (Safari)",
        "android": "Android (Chrome)",
        "fullscreen": "Desktop",
        "600x738": "Mobile",
    }
    nav_label = nav_map.get(str(navegador).lower(), str(navegador))
    vp_label = vp_map.get(str(viewport).lower(), str(viewport))
    return nav_label, vp_label


def enviar_email_resultados_consolidados(resultados_ejecucion):
    """
    Envía un único email consolidado con resultados de múltiples ejecuciones.

    Args:
        resultados_ejecucion: lista de dicts con claves esperadas:
            - pais
            - excel_path
            - screenshots_dir
            - navegador (opcional)
            - viewport (opcional)
            - estado (opcional)
    """
    try:
        if not resultados_ejecucion:
            print("❌ No hay resultados para enviar en el consolidado.")
            return False

        config = cargar_config_global()

        enviar_mail = bool(config.get("enviar_mail", False))
        adjuntar_resultados = bool(config.get("adjuntar_resultados", True))
        adjuntar_screenshots = bool(config.get("adjuntar_screenshots", True))

        if not enviar_mail:
            print("📭 Envío de email DESHABILITADO por Configuración Global. Se omite el envío consolidado.")
            return True

        fecha_actual = datetime.now().strftime("%d/%m/%Y")

        # Acumular resultados por país para el resumen PASSED/FAILED
        pais_errores = {}   # pais -> total con_errores sumado
        pais_exitosos = {}  # pais -> total exitosos sumado
        adjuntos = []
        detalles_fallidos = []   # bloques de detalle solo para países FAILED
        detalles_url_pais = []   # [(pais, nav_label, vp_label, todos_urls)] para tabla por URL
        avisos_sin_valor = []    # [(pais, nav_label, vp_label, detalles_sin_valor)]

        for idx, resultado in enumerate(resultados_ejecucion, start=1):
            pais = resultado.get("pais", "N/A")
            navegador = resultado.get("navegador", "N/A")
            viewport = resultado.get("viewport", "N/A")
            excel_path = resultado.get("excel_path")
            screenshots_dir = resultado.get("screenshots_dir")

            nav_label, vp_label = _label_navegador(navegador, viewport)

            if not excel_path or not os.path.exists(excel_path):
                print(f"      ⚠️ ADVERTENCIA: Excel no encontrado - {excel_path}")
                pais_errores[pais] = pais_errores.get(pais, 0) + 1
                continue

            errores = analizar_errores_excel(excel_path)
            exitosos = int(errores.get("exitosos", 0))
            con_errores = int(errores.get("con_errores", 0))

            pais_errores[pais] = pais_errores.get(pais, 0) + con_errores
            pais_exitosos[pais] = pais_exitosos.get(pais, 0) + exitosos

            _ok_items_c = [{'linea': d['linea'], 'url': d.get('url',''), 'url_secure': d.get('url_secure',''), 'ok': True, 'error': ''} for d in (errores.get('detalles_ok') or [])]
            _err_items_c = [{'linea': d['linea'], 'url': d.get('url',''), 'url_secure': d.get('url_secure',''), 'ok': False, 'error': d.get('error', '')} for d in (errores.get('detalles') or [])]
            if _ok_items_c or _err_items_c:
                detalles_url_pais.append((pais, nav_label, vp_label, _err_items_c + _ok_items_c))

            if con_errores > 0:
                bloque = f"❌ {pais.upper()} — {nav_label} / {vp_label}\n"
                bloque += f"   Errores: {con_errores} leads NO enviados, {exitosos} OK\n\n"
                if errores.get("detalles"):
                    for detalle in errores["detalles"]:
                        url = detalle.get("url", "N/A")
                        error = detalle.get("error", "Sin descripción")
                        linea = detalle.get("linea", "?")
                        bloque += f"   ❌ Linea {linea} | URL: {url}\n      Error: {error}\n\n"
                detalles_fallidos.append(bloque)

            _sv_items = errores.get("detalles_sin_valor") or []
            if _sv_items:
                avisos_sin_valor.append((pais, nav_label, vp_label, _sv_items))

            # Adjuntos
            if adjuntar_resultados and os.path.getsize(excel_path) < 24 * 1024 * 1024:
                adjuntos.append(excel_path)
            if adjuntar_screenshots and screenshots_dir and os.path.exists(screenshots_dir):
                zip_files = crear_zip_de_carpeta(screenshots_dir)
                if zip_files:
                    adjuntos.extend(zip_files)

        # Clasificar países
        todos_paises = list(dict.fromkeys(
            [r.get("pais", "N/A") for r in resultados_ejecucion]
        ))
        paises_failed = [p for p in todos_paises if pais_errores.get(p, 0) > 0]
        paises_passed = [p for p in todos_paises if pais_errores.get(p, 0) == 0]

        total_exitosos = sum(pais_exitosos.values())
        total_errores = sum(pais_errores.values())
        resultado_global = "PASS" if not paises_failed else "FAILED"

        partes_asunto = []
        if paises_passed:
            partes_asunto.append(f"{_abrev_paises(paises_passed)} ✓")
        if paises_failed:
            partes_asunto.append(f"{_abrev_paises(paises_failed)} ✗")
        asunto = f"[{resultado_global}] Osocio — {fecha_actual} — {' | '.join(partes_asunto)}"

        icono_global = "✅" if resultado_global == "PASS" else "❌"
        cuerpo = f"{icono_global} RESULTADO GLOBAL: {resultado_global}\n\n"

        if paises_failed:
            cuerpo += f"FAILED ({len(paises_failed)}): {' | '.join(paises_failed)}\n"
        if paises_passed:
            cuerpo += f"PASSED ({len(paises_passed)}): {' | '.join(paises_passed)}\n"

        if avisos_sin_valor:
            cuerpo += ("\n⚠️ CAMPOS SIN VALOR ASIGNADO — quedaron vacíos al enviar. "
                       "Asignales un valor en Osocio → Envío de Leads → ⚙ IDs Dinámicos:\n")
            for _p, _nav, _vp, _items in avisos_sin_valor:
                for _sv in _items:
                    cuerpo += f"  • {_p} ({_nav}/{_vp}) — Línea {_sv['linea']}: {_sv.get('campos','')}\n"

        cuerpo += "\nSaludos,\nAutomacion de Formularios"

        _consolidado_html = ""
        if detalles_url_pais:
            for _pais, _nav, _vp, _urls in detalles_url_pais:
                import html as _h
                _consolidado_html += (
                    f'<p style="font-weight:bold;margin:14px 0 6px 0">'
                    f'{_h.escape(_pais.upper())} — {_h.escape(_nav)} / {_h.escape(_vp)}</p>'
                )
                _consolidado_html += _build_url_table_html(_urls)

        destinatario = obtener_email_destinatario()

        if not destinatario or destinatario == "correo@example.com":
            print(f"❌ ERROR: Email destinatario no configurado o es el default. No se puede enviar.")
            return False

        try:
            _encolar_email(destinatario, asunto, cuerpo, adjuntos, html_extra=_consolidado_html)
            return True
        except Exception as e_cola:
            print(f"❌ Error al encolar email: {e_cola}")
            return False

    except Exception as e:
        print(f"\n❌ ERROR CRÍTICO al enviar email consolidado de resultados: {e}")
        import traceback
        print(traceback.format_exc())
        return False


def enviar_email_comparador_dealers(pais, carpeta_reporte, excel_path, counts,
                                    form_url="", landing_url="", incluir_capturas=True):
    """Envía por email el resultado de UNA corrida del Comparador de Dealers, con el mismo
    criterio que el resto de la app:
      - Respeta el flag global 'enviar_mail' (si está apagado, no envía).
      - incluir_capturas=True (modo 'Excel + Capturas'): adjunta un ZIP con TODA la carpeta
        del reporte (Excel + capturas). incluir_capturas=False (modo 'Solo Excel'): adjunta
        sólo el Excel.
      - Destinatario: el mismo 'email_destinatario' global que Envío de Leads.
    Devuelve True si se encoló (o si el envío está deshabilitado), False ante error.
    """
    try:
        config = cargar_config_global()
        if not bool(config.get("enviar_mail", False)):
            print("📭 Envío de email deshabilitado por Configuración Global. Se omite (Comparador).")
            return True

        c = counts or {}
        pass_n = int(c.get("PASS", 0))
        fail_n = int(c.get("FAIL", 0))
        extra_n = int(c.get("EXTRA", 0))
        dup_n = int(c.get("DUPLICADO", 0))
        oculto_n = int(c.get("OCULTO", 0))
        nota_n = int(c.get("NOTA", 0))

        fecha_actual = datetime.now().strftime("%d/%m/%Y")
        estado = "PASS" if fail_n == 0 else "FAILED"
        icono = "✅" if fail_n == 0 else "❌"
        asunto = (f"[{estado}] Comparador Dealers {_PAIS_ABREV.get(pais, pais)} {fecha_actual} — "
                  f"{pass_n} OK / {fail_n} FAIL")

        cuerpo = f"{icono} Comparador de Dealers — {pais}\nFecha: {fecha_actual}\n"
        if form_url:
            cuerpo += f"Formulario: {form_url}\n"
        if landing_url:
            cuerpo += f"Landing: {landing_url}\n"
        cuerpo += (
            f"\nResumen:\n"
            f"  🟢 PASS: {pass_n}\n"
            f"  🔴 FAIL: {fail_n}\n"
            f"  🟡 EXTRA (en el form, no en el Excel): {extra_n}\n"
            f"  🔵 DUPLICADO (repetido en el dropdown): {dup_n}\n"
            f"  🟠 OCULTO (solo en filas ocultas del Excel): {oculto_n}\n"
            f"  🔷 NOTA (mismo dealer, nombre con diferencias menores): {nota_n}\n"
        )
        cuerpo += "\nSaludos,\nAutomación de Formularios"

        adjuntos = []
        if incluir_capturas and carpeta_reporte and os.path.isdir(carpeta_reporte):
            # ZIP con toda la carpeta (Excel + capturas)
            zip_files = crear_zip_de_carpeta(carpeta_reporte)
            if zip_files:
                adjuntos.extend(zip_files)
            elif excel_path and os.path.exists(excel_path):
                adjuntos.append(excel_path)
        elif excel_path and os.path.exists(excel_path):
            if os.path.getsize(excel_path) < 24 * 1024 * 1024:
                adjuntos.append(excel_path)
            else:
                cuerpo += f"\n⚠️ NOTA: El Excel es muy grande y no se pudo adjuntar. Ubicación: {excel_path}"

        destinatario = obtener_email_destinatario()
        try:
            _encolar_email(destinatario, asunto, cuerpo, adjuntos)
        except Exception as e:
            print(f"❌ Error encolando email (Comparador): {e}")
            return False
        return True

    except Exception as e:
        print(f"❌ Error al enviar email del Comparador de Dealers: {e}")
        return False