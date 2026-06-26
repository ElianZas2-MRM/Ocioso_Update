import argparse
import importlib
import importlib.util
import os
import sys

from interface.main_interface import iniciar_interfaz
from utils.paths import BASE_DIR, FORMS_DIR


ENVIRONMENTS = {
    "chrome_desktop": {"browser": "chrome", "viewport": "fullscreen"},
    "chrome_mobile": {"browser": "chrome", "viewport": "600x738"},
    "firefox_desktop": {"browser": "firefox", "viewport": "fullscreen"},
    "firefox_mobile": {"browser": "firefox", "viewport": "600x738"},
    "edge_desktop": {"browser": "edge", "viewport": "fullscreen"},
    "edge_mobile": {"browser": "edge", "viewport": "600x738"},
}


def _ensure_form_paths():
    for path in (BASE_DIR, FORMS_DIR, os.path.join(BASE_DIR, "core")):
        if path not in sys.path:
            sys.path.insert(0, path)


def _load_country_run_function(country_name):
    _ensure_form_paths()

    module_name = f"Formulario_{country_name}_Main"
    script_path = os.path.join(FORMS_DIR, f"{module_name}.py")

    if os.path.exists(script_path):
        spec = importlib.util.spec_from_file_location(module_name, script_path)
        if spec is None or spec.loader is None:
            raise ImportError(f"No se pudo crear spec para {script_path}")
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
    else:
        module = importlib.import_module(module_name)

    run_func = getattr(module, f"run_formularios_{country_name}", None)
    if not callable(run_func):
        raise AttributeError(f"No se encontró run_formularios_{country_name} en {module_name}")

    return run_func


def _run_country(country_name, environment, headless=False, enviar_email=True):
    config = ENVIRONMENTS.get(environment)
    if config is None:
        raise ValueError(f"Entorno '{environment}' no reconocido")

    run_func = _load_country_run_function(country_name)
    return run_func(
        browser=config["browser"],
        viewport=config["viewport"],
        headless=headless,
        enviar_email=enviar_email,
    )


def _run_autonomous():
    from autonomous_runner import main as autonomous_main

    return autonomous_main()


def _parse_args():
    parser = argparse.ArgumentParser(add_help=True)
    parser.add_argument("--autonomous", action="store_true", help="Ejecuta el planificador autónomo")
    parser.add_argument("--run-country", dest="country_name", help="Ejecuta un país puntual sin abrir la UI")
    parser.add_argument("--environment", default="chrome_desktop", choices=sorted(ENVIRONMENTS.keys()))
    parser.add_argument("--no-email", action="store_true", help="No envía email al finalizar")
    return parser.parse_args()


def _ensure_runtime_dirs():
    for folder_name in ("drivers", "resultados", "data", "json", "temporales"):
        os.makedirs(os.path.join(BASE_DIR, folder_name), exist_ok=True)

if __name__ == "__main__":
    _ensure_runtime_dirs()
    args = _parse_args()
    if args.autonomous:
        _run_autonomous()
    elif args.country_name:
        _run_country(
            args.country_name,
            args.environment,
            headless=False,
            enviar_email=not args.no_email,
        )
    else:
        iniciar_interfaz()
    #python "E:\Ariel\Scripts\Form GDCP\Form_Automation_Project\run.py"
    
    # venv   cd "E:\Ariel\Scripts\Form GDCP\Form_Automation_Project"
    # # .\venv\Scripts\activate
    