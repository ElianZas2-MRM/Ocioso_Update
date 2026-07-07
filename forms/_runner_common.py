"""
_runner_common.py — Lógica compartida de ejecución para todos los países.
Contiene run_country_form() (crea el formulario y lo corre), get_runner() (devuelve
la función de runner dinámico por nombre de país) y run_cli() (soporte para línea de comandos).
"""
import os
import sys

# Asegurar rutas — compatible con script directo y EXE empaquetado
if getattr(sys, 'frozen', False):
    PROJECT_ROOT = os.path.dirname(sys.executable)
else:
    PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

CORE_DIR = os.path.join(PROJECT_ROOT, "core")

for _p in (PROJECT_ROOT, CORE_DIR):
    if _p not in sys.path:
        sys.path.insert(0, _p)


ENVIRONMENTS = {
    "chrome_desktop": {"browser": "chrome", "viewport": "fullscreen"},
    "chrome_mobile": {"browser": "chrome", "viewport": "600x738"},
    "firefox_desktop": {"browser": "firefox", "viewport": "fullscreen"},
    "firefox_mobile": {"browser": "firefox", "viewport": "600x738"},
    "edge_desktop": {"browser": "edge", "viewport": "fullscreen"},
    "edge_mobile": {"browser": "edge", "viewport": "600x738"},
}


def run_country_form(form_class, country_name, browser="chrome", viewport="fullscreen", headless=False, enviar_email=True, background=True, progress_callback=None, email_callback=None, stop_event=None):
    formulario = form_class(browser=browser, viewport=viewport, headless=headless, background=background)
    formulario.run(progress_callback=progress_callback)

    if not enviar_email:
        return formulario

    # Si el usuario presionó Detener no enviamos email — no es un fallo real
    if stop_event and stop_event.is_set():
        return formulario

    resultados_path = getattr(formulario, "RESULTADOS_PATH", None)
    screenshot_dir = getattr(formulario, "SCREENSHOT_DIR", None)
    if not resultados_path or not screenshot_dir:
        return formulario

    try:
        from interface.helpers_interface import enviar_email_resultados
        if email_callback:
            try:
                email_callback("sending")
            except Exception:
                pass
        enviar_email_resultados(country_name, resultados_path, screenshot_dir, browser=browser, viewport=viewport)
        if email_callback:
            try:
                email_callback("ok")
            except Exception:
                pass
    except Exception:
        if email_callback:
            try:
                email_callback("fail")
            except Exception:
                pass

    return formulario


def get_runner(country_name: str):
    """
    Devuelve la función run_formularios_<País> para el país dado.
    Reemplaza la carga dinámica de los 9 Formulario_*_Main.py eliminados.
    """
    from generic_country_base import GenericCountryBase

    def _runner(browser="chrome", viewport="fullscreen", headless=False, enviar_email=True, background=True, progress_callback=None, email_callback=None, stop_event=None, is_scheduled=False):
        class _DynamicCountry(GenericCountryBase):
            def __init__(self, browser=browser, viewport=viewport, headless=headless, background=background):
                super().__init__(country_name, browser=browser, viewport=viewport, headless=headless, background=background, is_scheduled=is_scheduled)

        return run_country_form(
            _DynamicCountry,
            country_name,
            browser=browser,
            viewport=viewport,
            headless=headless,
            enviar_email=enviar_email,
            background=background,
            progress_callback=progress_callback,
            email_callback=email_callback,
            stop_event=stop_event,
        )

    _runner.__name__ = f"run_formularios_{country_name}"
    return _runner


def run_cli(run_func):
    env = sys.argv[1] if len(sys.argv) > 1 else "chrome_desktop"
    headless = len(sys.argv) > 2 and sys.argv[2].lower() == "true"
    enviar_email = True
    if len(sys.argv) > 3:
        enviar_email = sys.argv[3].lower() == "true"

    if env not in ENVIRONMENTS:
        raise SystemExit(f"Entorno '{env}' no reconocido. Opciones válidas: {', '.join(ENVIRONMENTS.keys())}")

    config = ENVIRONMENTS[env]
    return run_func(
        browser=config["browser"],
        viewport=config["viewport"],
        headless=headless,
        enviar_email=enviar_email,
    )