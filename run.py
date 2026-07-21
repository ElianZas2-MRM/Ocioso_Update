"""
run.py — Punto de entrada principal de la aplicación.
Lanza la interfaz gráfica (por defecto), permite correr un país directo por línea de comandos
o iniciar el modo autónomo (ejecuta tests programados en background).
"""
import argparse
import importlib
import importlib.util
import os
import sys

# Fix de conexión LambdaTest en redes de oficina con proxy (Netskope, Zscaler, etc.):
# Windows ya confía en el certificado que pone ese proxy, pero Python trae su propia lista
# de certificados (certifi) que NO lo conoce, y rechaza la conexión con LambdaTest.
# truststore hace que Python use la MISMA lista de confianza que ya usa Windows (no
# desactiva ninguna verificación, solo iguala a Python con Windows). Tiene que ir acá,
# antes de cualquier otro import, para que aplique desde la primera conexión de red.
try:
    import truststore
    truststore.inject_into_ssl()
except Exception:
    pass

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

    try:
        try:
            from forms._runner_common import get_runner
        except ImportError:
            from _runner_common import get_runner  # type: ignore[import-not-found]
        return get_runner(country_name)
    except (ImportError, AttributeError):
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


def _run_country(country_name, environment, headless=False, enviar_email=True, is_scheduled=False):
    config = ENVIRONMENTS.get(environment)
    if config is None:
        raise ValueError(f"Entorno '{environment}' no reconocido")

    run_func = _load_country_run_function(country_name)
    return run_func(
        browser=config["browser"],
        viewport=config["viewport"],
        headless=headless,
        enviar_email=enviar_email,
        is_scheduled=is_scheduled,
    )


def _run_autonomous():
    from autonomous_runner import main as autonomous_main

    return autonomous_main()


def _run_lambdatest(lt_type, pais, build_name=""):
    """Ejecuta LambdaTest Mac o Android para un país y muestra el resumen."""
    lt_mac_dir = os.path.join(BASE_DIR, "lambdatest_mac")
    lt_android_dir = os.path.join(BASE_DIR, "lambdatest_android")

    if lt_type == "mac":
        if lt_mac_dir not in sys.path:
            sys.path.insert(0, lt_mac_dir)
        import lt_controller  # type: ignore[import]
        summary = lt_controller.run(pais=pais, build_name=build_name)
    elif lt_type == "android":
        if lt_android_dir not in sys.path:
            sys.path.insert(0, lt_android_dir)
        import lt_android_controller  # type: ignore[import]
        summary = lt_android_controller.run(pais=pais, build_name=build_name)
    else:
        raise ValueError(f"Tipo LambdaTest no reconocido: {lt_type!r}")

    ok = summary.get("ok", 0)
    failed = summary.get("failed", 0)
    total = summary.get("total", 0)
    print(f"LambdaTest {lt_type} — {pais}: {ok}/{total} OK, {failed} errores")
    return summary


def _parse_args():
    parser = argparse.ArgumentParser(add_help=True)
    parser.add_argument("--autonomous", action="store_true", help="Ejecuta el planificador autónomo")
    parser.add_argument("--run-country", dest="country_name", help="Ejecuta un país puntual sin abrir la UI")
    parser.add_argument("--environment", default="chrome_desktop", choices=sorted(ENVIRONMENTS.keys()))
    parser.add_argument("--no-email", action="store_true", help="No envía email al finalizar")
    parser.add_argument("--scheduled", action="store_true", help="Indica que es una ejecución programada")
    parser.add_argument("--run-lambdatest", dest="lt_type", choices=["mac", "android"],
                        help="Ejecuta LambdaTest (mac o android) para el país indicado con --pais")
    parser.add_argument("--pais", dest="lt_pais", help="País para --run-lambdatest")
    parser.add_argument("--build-name", dest="lt_build_name", default="",
                        help="Nombre del build en LambdaTest (opcional)")
    return parser.parse_args()


def _ensure_runtime_dirs():
    for folder_name in ("drivers", "resultados", "data", "json", "temporales"):
        os.makedirs(os.path.join(BASE_DIR, folder_name), exist_ok=True)

if __name__ == "__main__":
    _ensure_runtime_dirs()
    args = _parse_args()
    if args.autonomous:
        _run_autonomous()
    elif args.lt_type:
        if not args.lt_pais:
            print("Error: --run-lambdatest requiere también --pais <nombre_país>")
            sys.exit(1)
        _run_lambdatest(args.lt_type, args.lt_pais, build_name=args.lt_build_name)
    elif args.country_name:
        _run_country(
            args.country_name,
            args.environment,
            headless=False,
            enviar_email=not args.no_email,
            is_scheduled=args.scheduled,
        )
    else:
        iniciar_interfaz()
    #python "E:\Ariel\Scripts\Form GDCP\Form_Automation_Project\run.py"
    
    # venv   cd "E:\Ariel\Scripts\Form GDCP\Form_Automation_Project"
    # # .\venv\Scripts\activate
    