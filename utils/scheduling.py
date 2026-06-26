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

def guardar_programacion(programacion):
    """Guarda la programación en un archivo JSON en carpeta json/ o limpia el archivo si programacion es None"""
    try:
        # Si programacion es None, simplemente limpiar el archivo
        if programacion is None:
            json_path = os.path.join(JSON_DIR, "programacion_test.json")
            if os.path.exists(json_path):
                os.remove(json_path)
                print("✅ Archivo de programación eliminado (guardar_programacion(None))")
            return True
        
        # Crear carpeta json si no existe
        if not os.path.exists(JSON_DIR):
            os.makedirs(JSON_DIR)
            print(f"📁 Carpeta '{JSON_DIR}' creada")

        # Convertir datetime a string para JSON
        programacion_serializable = {
            "fecha_hora": programacion["fecha_hora"].strftime("%Y-%m-%d %H:%M:%S"),
            "paises": programacion["paises"],
            "navegadores": programacion["navegadores"],
            "viewports": programacion["viewports"]
        }

        json_path = os.path.join(JSON_DIR, "programacion_test.json")
        with open(json_path, "w", encoding='utf-8') as f:
            json.dump(programacion_serializable, f, indent=2)

        print(f"✅ Programación guardada en {json_path}")
        return True
    except Exception as e:
        print(f"❌ Error guardando programación: {e}")
        return False

def cargar_programacion():
    """Carga la programación desde archivo JSON en carpeta json/"""
    try:
        json_path = os.path.join(JSON_DIR, "programacion_test.json")
        if os.path.exists(json_path):
            with open(json_path, "r", encoding='utf-8') as f:
                data = json.load(f)

            # Convertir string a datetime
            programacion = {
                "fecha_hora": datetime.strptime(data["fecha_hora"], "%Y-%m-%d %H:%M:%S"),
                "paises": data["paises"],
                "navegadores": data["navegadores"],
                "viewports": data["viewports"]
            }

            print(f"Programación cargada desde {json_path}")
            return programacion
        return None
    except Exception as e:
        print(f"Error cargando programación: {e}")
        return None

def limpiar_programacion():
    """Elimina el archivo de programación de carpeta json/"""
    try:
        json_path = os.path.join(JSON_DIR, "programacion_test.json")
        if os.path.exists(json_path):
            os.remove(json_path)
            print("Archivo de programación eliminado")
        return True
    except Exception as e:
        print(f"Error limpiando programación: {e}")
        return False