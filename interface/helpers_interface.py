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
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

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


def _worker_envio_emails():
    """Worker que envía emails via Outlook COM. Si Outlook no está instalado, loguea el error."""
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
            builtins.print(f"⚠️ GetActiveObject falló ({_get_err}), intentando Dispatch...")
            _outlook_instance = win32.Dispatch("Outlook.Application")
    except Exception as e:
        builtins.print(f"❌ No se puede enviar email: {e}")
        log_runtime(f"Outlook no disponible: {e}", level="ERROR")
        while True:
            try:
                _cola_emails.get_nowait()
                _cola_emails.task_done()
            except queue.Empty:
                break
        return

    while True:
        try:
            data = _cola_emails.get(timeout=5)
            if data is None:
                break

            destinatarios, asunto, cuerpo, adjuntos = data

            try:
                mail = _outlook_instance.CreateItem(0)
                mail.To = destinatarios if isinstance(destinatarios, str) else "; ".join(destinatarios)
                mail.Subject = asunto
                mail.Body = cuerpo
                if adjuntos:
                    for archivo in adjuntos:
                        if archivo and os.path.exists(archivo):
                            mail.Attachments.Add(os.path.abspath(archivo))
                mail.Send()
                builtins.print(f"✅ Email enviado (Outlook) a: {mail.To} | Asunto: {mail.Subject}")
                time.sleep(0.5)

                if adjuntos:
                    for archivo in adjuntos:
                        try:
                            if archivo and archivo.endswith(".zip") and os.path.exists(archivo):
                                os.remove(archivo)
                        except Exception:
                            pass

            except Exception as e:
                builtins.print(f"❌ Error al enviar email via Outlook: {e}")
                log_runtime(str(e), level="ERROR")

            finally:
                _cola_emails.task_done()

        except queue.Empty:
            pass
        except Exception as e:
            builtins.print(f"❌ Error en worker de email: {e}")
            log_runtime(str(e), level="ERROR")

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


def _encolar_email(destinatarios, asunto, cuerpo, adjuntos=None):
    """Encola un email para ser enviado por el worker."""
    _iniciar_worker_envio()
    _cola_emails.put((destinatarios, asunto, cuerpo, adjuntos))


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
    """Devuelve el email destinatario actual (persistido) o el default."""
    config = cargar_config_global()
    email = (config.get("email_destinatario") or "").strip()
    return email if email else DEFAULT_EMAIL_DESTINATARIO


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
        errores_mask = resultados.str.contains("error", case=False, na=False)

        errores = df[errores_mask]
        con_errores = int(errores_mask.sum())
        procesados = int(procesados_mask.sum())
        exitosos = procesados - con_errores
        no_procesados = total_filas - procesados

        # Construir detalles estructurados
        detalles = []
        detalles_texto = []
        
        for i, fila in errores.iterrows():
            url = str(fila[col_url]) if col_url and pd.notna(fila[col_url]) else "(sin URL)"
            mensaje_error = str(fila[col_resultado]).strip()
            
            # Limpiar el mensaje de error
            mensaje_limpio = limpiar_mensaje_error(mensaje_error)
            
            linea_excel = i + 2  # +2 porque Excel empieza en 1 y la fila 1 son los encabezados

            detalles.append({
                'url': url,
                'error': mensaje_limpio,
                'linea': linea_excel
            })
            detalles_texto.append(
                f"➡️ En la URL: {url}\n"
                f"   Línea: {linea_excel}\n"
                f"   ❌ {mensaje_limpio}"
            )

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
            'mensaje': mensaje
        }

    except Exception as e:
        return {
            'total': 0,
            'exitosos': 0,
            'con_errores': 0,
            'detalles': [],
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

def enviar_email_resultados(pais, excel_path, screenshots_dir):
    """
    Envía un email con los resultados de la ejecución del formulario.
    
    Args:
        pais: Nombre del país (ej: "Argentina")
        excel_path: Ruta al archivo Excel de resultados
        screenshots_dir: Ruta a la carpeta de screenshots
    """
    try:
        # Validar que existan los archivos
        if not os.path.exists(excel_path):
            print(f"❌ No se encontró el archivo Excel: {excel_path}")
            return False
        
        if not os.path.exists(screenshots_dir):
            print(f"❌ No se encontró la carpeta de screenshots: {screenshots_dir}")
            return False
        
        # Analizar errores del Excel
        errores = analizar_errores_excel(excel_path)
        total_filas = errores.get('total', 0)
        exitosos = errores.get('exitosos', 0)
        con_errores = errores.get('con_errores', 0)
        no_procesados = errores.get('no_procesados', 0)
        
        # Preparar asunto
        fecha_actual = datetime.now().strftime("%d/%m/%Y")
        asunto = f"Resultados Osocio {pais} {fecha_actual}"
        
        # Preparar cuerpo del mensaje
        cuerpo = f"""Resultados de la ejecución de formularios - {pais}
Fecha: {fecha_actual}

=== RESUMEN ===
Exitosos: {exitosos}
Con errores: {con_errores}

"""
        
        if con_errores > 0 and errores.get('detalles'):
            cuerpo += "=== ERRORES DETECTADOS ===\n\n"
            for detalle in errores['detalles'][:10]:  # Máximo 10 errores en preview
                url = detalle.get('url', 'N/A')
                error = detalle.get('error', 'Sin descripción')
                cuerpo += f"URL: {url}\nError: {error}\n\n"
            
            if len(errores['detalles']) > 10:
                cuerpo += f"... y {len(errores['detalles']) - 10} errores más (ver Excel adjunto)\n\n"
        else:
            cuerpo += "✅ Todos los formularios se completaron exitosamente.\n\n"
        
        cuerpo += "Saludos,\nAutomación de Formularios"
        
        # Respetar configuración global (enviar/no enviar + adjuntos)
        config = cargar_config_global()
        enviar_mail = bool(config.get("enviar_mail", False))
        adjuntar_resultados = bool(config.get("adjuntar_resultados", True))
        adjuntar_screenshots = bool(config.get("adjuntar_screenshots", True))

        if not enviar_mail:
            print("📭 Envío de email deshabilitado por Configuración Global. Se omite el envío.")
            return True

        # Preparar adjuntos
        adjuntos = []
        
        # Agregar Excel
        if adjuntar_resultados:
            if os.path.getsize(excel_path) < 24 * 1024 * 1024:  # Menor a 24 MB
                adjuntos.append(excel_path)
            else:
                print(f"⚠️ Excel excede 24 MB, no se adjuntará")
                cuerpo += f"\n⚠️ NOTA: El archivo Excel es muy grande y no se pudo adjuntar. Ubicación: {excel_path}"
        else:
            pass
        
        # Comprimir y agregar screenshots
        zip_files = None
        if adjuntar_screenshots:
            zip_files = crear_zip_de_carpeta(screenshots_dir)
        
        if adjuntar_screenshots:
            if zip_files:
                # Agregar todos los ZIPs (límite de 18 MB por archivo ya se respeta en crear_zip_de_carpeta)
                for zip_file in zip_files:
                    adjuntos.append(zip_file)
            else:
                print(f"⚠️ No se pudieron comprimir los screenshots")
                cuerpo += f"\n⚠️ NOTA: Los screenshots no se pudieron comprimir. Ubicación: {screenshots_dir}"
        else:
            pass
        
        # Enviar email
        destinatario = obtener_email_destinatario()
        
        try:
            _encolar_email(destinatario, asunto, cuerpo, adjuntos)
        except Exception as e:
            print(f"❌ Error encolando email: {e}")
            return False
        
        return True
        
    except Exception as e:
        print(f"❌ Error al enviar email de resultados: {e}")
        return False


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
        asunto = f"Resultados Osocio Programación {fecha_actual}"

        total_corridas = 0
        total_filas = 0
        total_exitosos = 0
        total_errores = 0
        total_no_procesados = 0
        detalles_cuerpo = []
        adjuntos = []
        for idx, resultado in enumerate(resultados_ejecucion, start=1):
            pais = resultado.get("pais", "N/A")
            navegador = resultado.get("navegador", "N/A")
            viewport = resultado.get("viewport", "N/A")
            estado = resultado.get("estado", "desconocido")
            excel_path = resultado.get("excel_path")
            screenshots_dir = resultado.get("screenshots_dir")

            if not excel_path or not os.path.exists(excel_path):
                print(f"      ⚠️ ADVERTENCIA: Excel no encontrado - {excel_path}")
                detalles_cuerpo.append(
                    f"[{idx}] {pais} ({navegador}/{viewport})\n"
                    f"Estado: {estado}\n"
                    "⚠️ No se encontró archivo Excel de resultados.\n"
                )
                continue

            total_corridas += 1
            errores = analizar_errores_excel(excel_path)
            filas = int(errores.get("total", 0))
            exitosos = int(errores.get("exitosos", 0))
            con_errores = int(errores.get("con_errores", 0))
            no_procesados = int(errores.get("no_procesados", 0))

            total_filas += filas
            total_exitosos += exitosos
            total_errores += con_errores
            total_no_procesados += no_procesados

            bloque = f"Resultados de la ejecución de formularios - {pais} ({navegador}/{viewport})\nFecha: {fecha_actual}\n\n"
            bloque += "=== RESUMEN ===\n"
            bloque += f"Exitosos: {exitosos}\n"
            bloque += f"Con errores: {con_errores}\n\n"

            if con_errores > 0 and errores.get("detalles"):
                bloque += "=== ERRORES DETECTADOS ===\n\n"
                for detalle in errores["detalles"][:10]:
                    url = detalle.get("url", "N/A")
                    error = detalle.get("error", "Sin descripción")
                    bloque += f"URL: {url}\nError: {error}\n\n"
                if len(errores["detalles"]) > 10:
                    bloque += f"... y {len(errores['detalles']) - 10} errores más (ver Excel adjunto)\n\n"
            else:
                bloque += "✅ Todos los formularios se completaron exitosamente.\n\n"

            bloque += "---\n\n"
            detalles_cuerpo.append(bloque)

            if adjuntar_resultados:
                if os.path.getsize(excel_path) < 24 * 1024 * 1024:
                    adjuntos.append(excel_path)
                else:
                    print(f"      ⚠️ Excel muy grande (>24MB), no se adjunta")
                    detalles_cuerpo.append(
                        f"⚠️ Excel muy grande y no adjuntado: {excel_path}\n"
                    )

            if adjuntar_screenshots and screenshots_dir and os.path.exists(screenshots_dir):
                zip_files = crear_zip_de_carpeta(screenshots_dir)
                if zip_files:
                    adjuntos.extend(zip_files)
                else:
                    print(f"      ⚠️ No se pudieron comprimir screenshots")
                    detalles_cuerpo.append(
                        f"⚠️ No se pudieron comprimir screenshots de: {screenshots_dir}\n"
                    )

        cuerpo = (
            f"{''.join(detalles_cuerpo)}"
            "Saludos,\nAutomación de Formularios"
        )

        destinatario = obtener_email_destinatario()
        
        if not destinatario or destinatario == "correo@example.com":
            print(f"❌ ERROR: Email destinatario no configurado o es el default. No se puede enviar.")
            return False
        
        try:
            _encolar_email(destinatario, asunto, cuerpo, adjuntos)
            return True
        except Exception as e_cola:
            print(f"❌ Error al encolar email: {e_cola}")
            return False

    except Exception as e:
        print(f"\n❌ ERROR CRÍTICO al enviar email consolidado de resultados: {e}")
        import traceback
        print(traceback.format_exc())
        return False