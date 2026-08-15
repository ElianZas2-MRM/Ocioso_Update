"""
scheduling.py — Gestión de programación de tests automáticos.
Carga, guarda y limpia el archivo json/programacion_test.json que define
qué países, en qué horarios y con qué browser se ejecutan automáticamente.
"""
import os
import sys
import json
from datetime import datetime

# === RUTAS BASE ===
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JSON_DIR = os.path.join(BASE_DIR, "json")

def guardar_programacion(programacion, filename="programacion_test.json"):
    """Guarda la programación en archivo JSON. None = eliminar. Soporta esquema semanal y legado."""
    try:
        if programacion is None:
            json_path = os.path.join(JSON_DIR, filename)
            if os.path.exists(json_path):
                os.remove(json_path)
                print(f"Archivo {filename} eliminado")
            return True

        if not os.path.exists(JSON_DIR):
            os.makedirs(JSON_DIR)

        if programacion.get("tipo") == "semanal":
            # Se parte de la programación completa para no perder los campos propios
            # de cada modo (revisión masiva, t3_also, columnas del Excel matriz…) y
            # después se normalizan los campos base con sus valores por defecto.
            serializable = {k: v for k, v in programacion.items() if k != "fecha_hora"}
            serializable.update({
                "tipo": "semanal",
                "modo_tarea": programacion.get("modo_tarea", "leads"),
                "horarios":    programacion["horarios"],
                "paises":      programacion["paises"],
                "navegadores": programacion.get("navegadores", []),
                "viewports":   programacion.get("viewports", []),
                "dispositivo": programacion.get("dispositivo", "local"),
                "modo_excel":  programacion.get("modo_excel", "consecutivo"),
                "modo_mercados": programacion.get("modo_mercados", "consecutivo"),
            })
        else:
            serializable = {
                "fecha_hora": programacion["fecha_hora"].strftime("%Y-%m-%d %H:%M:%S"),
                "paises":      programacion["paises"],
                "navegadores": programacion["navegadores"],
                "viewports":   programacion["viewports"],
            }

        json_path = os.path.join(JSON_DIR, filename)
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(serializable, f, indent=2, ensure_ascii=False)

        # Sin emojis: este módulo también corre desde el ejecutor headless, donde la
        # consola de Windows usa cp1252 y un ✅ tira UnicodeEncodeError.
        print(f"Programacion guardada en {json_path}")
        return True
    except Exception as e:
        print(f"Error guardando programacion: {e}")
        return False

def cargar_programacion(filename="programacion_test.json"):
    """Carga la programación desde archivo JSON. Retorna dict con tipo='semanal' o None."""
    try:
        json_path = os.path.join(JSON_DIR, filename)
        if not os.path.exists(json_path):
            return None
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        if data.get("tipo") == "semanal":
            out = dict(data)
            out.update({
                "tipo":        "semanal",
                "modo_tarea":  data.get("modo_tarea", "leads"),
                "horarios":    data["horarios"],
                "paises":      data["paises"],
                "navegadores": data.get("navegadores", []),
                "viewports":   data.get("viewports", []),
                "dispositivo": data.get("dispositivo", "local"),
                "modo_excel":  data.get("modo_excel", "consecutivo"),
                "modo_mercados": data.get("modo_mercados", "consecutivo"),
            })
            return out

        return None
    except Exception as e:
        print(f"Error cargando programación ({filename}): {e}")
        return None

def limpiar_programacion(filename="programacion_test.json"):
    """Elimina el archivo de programación de carpeta json/"""
    try:
        json_path = os.path.join(JSON_DIR, filename)
        if os.path.exists(json_path):
            os.remove(json_path)
            print(f"Archivo {filename} eliminado")
        return True
    except Exception as e:
        print(f"Error limpiando programación ({filename}): {e}")
        return False