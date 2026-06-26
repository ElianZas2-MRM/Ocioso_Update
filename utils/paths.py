import os
import sys


def get_base_dir() -> str:
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    # Sube un nivel desde utils/ para llegar a la raíz del proyecto
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def get_bundle_dir() -> str:
    if getattr(sys, 'frozen', False):
        return getattr(sys, '_MEIPASS', os.path.join(get_base_dir(), '_internal'))
    return get_base_dir()


BASE_DIR = get_base_dir()
BUNDLE_DIR = get_bundle_dir()
FORMS_DIR = os.path.join(BASE_DIR, "forms")
DATA_DIR = os.path.join(BASE_DIR, "data")
ASSET_DIR = os.path.join(BUNDLE_DIR, "Asset")
RESULTS_DIR = os.path.join(BASE_DIR, "resultados")
JSON_DIR = os.path.join(BASE_DIR, "json")
TEMPORALES_DIR = os.path.join(BASE_DIR, "temporales")
DRIVERS_DIR = os.path.join(BASE_DIR, "drivers")
