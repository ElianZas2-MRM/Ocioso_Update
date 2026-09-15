"""
paths.py — Resolución de rutas del proyecto.
Detecta si la app corre como script Python normal o como EXE empaquetado con PyInstaller
y devuelve las rutas correctas para datos, drivers, resultados y archivos JSON.
"""
import os
import sys


def get_base_dir() -> str:
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    # Sube desde osocio/utils/paths.py hasta la raíz del proyecto: utils -> osocio -> raíz
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def get_bundle_dir() -> str:
    if getattr(sys, 'frozen', False):
        return getattr(sys, '_MEIPASS', os.path.join(get_base_dir(), '_internal'))
    return get_base_dir()


BASE_DIR = get_base_dir()
BUNDLE_DIR = get_bundle_dir()
FORMS_DIR = os.path.join(BASE_DIR, "osocio", "forms")
DATA_DIR = os.path.join(BASE_DIR, "data")
ASSET_DIR = os.path.join(BUNDLE_DIR, "Asset")
RESULTS_DIR = os.path.join(BASE_DIR, "resultados")

# Todo lo que la app produce vive adentro de resultados/, cada cosa en su subcarpeta.
# Antes estas tres colgaban sueltas de la raiz (resultados_lambdatestmac/,
# resultados_lambdatest_android/ y Dealerscheck_resultados/), que ensuciaba la raiz
# y obligaba a buscar los resultados de una corrida en cuatro lugares distintos.
RESULTS_LT_MAC_DIR = os.path.join(RESULTS_DIR, "lambdatest_mac")
RESULTS_LT_ANDROID_DIR = os.path.join(RESULTS_DIR, "lambdatest_android")
RESULTS_DEALERS_DIR = os.path.join(RESULTS_DIR, "dealers")
RESULTS_MASIVA_DIR = os.path.join(RESULTS_DIR, "resultado_urlsinsertas")
JSON_DIR = os.path.join(BASE_DIR, "json")
TEMPORALES_DIR = os.path.join(BASE_DIR, "temporales")
DRIVERS_DIR = os.path.join(BASE_DIR, "drivers")
