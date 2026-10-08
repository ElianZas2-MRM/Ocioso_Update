"""
bs_driver.py
============
Todo lo que es propio de BrowserStack Automate y no de Selenium: dónde conectarse, con
qué credenciales, qué dispositivo pedir, cómo marcar la sesión como pasada/fallada y cómo
traer el link del video al terminar.

El llenado de formularios NO vive acá: es Selenium puro y se reutiliza tal cual.

No se usa el SDK de BrowserStack (browserstack.yml + pytest): la app arma el driver a mano
desde la interfaz, igual que con LambdaTest, así que se conecta directo al hub.
"""

import json
import os
from typing import Dict, Tuple

from selenium.webdriver import Remote
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.remote.client_config import ClientConfig
from selenium.webdriver.safari.options import Options as SafariOptions

from osocio.paths import BASE_DIR as _OSOCIO_DIR

BS_HUB = "https://hub.browserstack.com/wd/hub"
BS_API = "https://api.browserstack.com/automate"

# Proyecto donde caen todas las corridas en el dashboard. BrowserStack descarta los
# caracteres que no acepta (paréntesis, guiones largos, etc.): usar [] en vez de ().
PROJECT_NAME = "Osocio Form Automation [Local]"

# Carpeta de este módulo: se usa para buscar un browserstack_credentials.txt puesto al lado.
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))


# ══════════════════════════════════════════════════════════════════════════════
# DISPOSITIVOS
# ══════════════════════════════════════════════════════════════════════════════

# Fijos a propósito (no regex): para QA conviene que cada corrida caiga en el mismo
# equipo y los resultados sean comparables. BrowserStack acepta regex en deviceName
# ("iPhone.*") si algún día la cola obliga a aflojar esto.
DISPOSITIVOS: Dict[str, Dict] = {
    "iphone": {
        "browserName": "safari",
        "platformName": "iOS",
        # Safari en iPhone bloquea switch_to.frame() hacia un iframe de otro dominio
        # ("Blocked a frame ... cross-origin frame", probado en iOS 16 y 17). El form se
        # llena igual DENTRO de la landing (nunca abriéndolo aparte): ver capa nativa.
        "iframe_cross_origin": False,
        "bstack": {
            "deviceName": "iPhone 14 Pro Max",
            "osVersion": "16",
            "realMobile": "true",
        },
    },
    "android": {
        "browserName": "chrome",
        "platformName": "android",
        "bstack": {
            "deviceName": "Samsung Galaxy S26 Ultra",
            "osVersion": "16.0",
            "realMobile": "true",
        },
    },
    "mac": {
        "browserName": "safari",
        "platformName": "mac",
        "bstack": {
            "os": "OS X",
            "osVersion": "Sequoia",
            "resolution": "1920x1080",
        },
    },
}


def plataformas() -> Tuple[str, ...]:
    return tuple(DISPOSITIVOS)


def puede_entrar_a_iframe(platform: str) -> bool:
    """False si el navegador no deja entrar a un iframe de otro dominio (iOS Safari)."""
    return DISPOSITIVOS[platform].get("iframe_cross_origin", True)


# ══════════════════════════════════════════════════════════════════════════════
# CREDENCIALES
# ══════════════════════════════════════════════════════════════════════════════

#: Mismos nombres que usa la propia documentación de BrowserStack.
_ENV_USER_VARS = ("BROWSERSTACK_USERNAME",)
_ENV_KEY_VARS = ("BROWSERSTACK_ACCESS_KEY",)

CREDENTIALS_FILENAME = "browserstack_credentials.txt"


def _first_env(names):
    for n in names:
        v = (os.environ.get(n) or "").strip()
        if v:
            return v
    return ""


def load_credentials(filepath: str = "") -> Tuple[str, str]:
    """Credenciales de BrowserStack: primero variables de entorno, si no el archivo.

    Mismo esquema que LambdaTest: el entorno evita dejar la access key en claro en el
    disco, y browserstack_credentials.txt (username=/access_key=) queda como fallback.
    """
    if not filepath:
        env_user, env_key = _first_env(_ENV_USER_VARS), _first_env(_ENV_KEY_VARS)
        if env_user and env_key:
            return env_user, env_key

        for candidate in [
            os.path.join(_THIS_DIR, CREDENTIALS_FILENAME),
            os.path.join(_OSOCIO_DIR, CREDENTIALS_FILENAME),
        ]:
            if os.path.exists(candidate):
                filepath = candidate
                break

    if not os.path.exists(filepath or ""):
        raise FileNotFoundError(
            "No se encontraron credenciales de BrowserStack.\n"
            "Opción 1 (recomendada) — variables de entorno:\n"
            "    setx BROWSERSTACK_USERNAME tu_usuario\n"
            "    setx BROWSERSTACK_ACCESS_KEY tu_key\n"
            "  setx no afecta a las terminales ya abiertas: si corrés desde VS Code, hay que\n"
            "  cerrar y abrir VS Code entero. Para probar ya, en la misma terminal PowerShell:\n"
            "    $env:BROWSERSTACK_USERNAME = \"tu_usuario\"\n"
            "    $env:BROWSERSTACK_ACCESS_KEY = \"tu_key\"\n"
            f"Opción 2 — archivo {CREDENTIALS_FILENAME} en la raíz del proyecto con:\n"
            "    username=tu_usuario\n    access_key=tu_key"
        )

    username = access_key = None
    with open(filepath, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            k, v = k.strip().lower(), v.strip()
            if k == "username":
                username = v
            elif k == "access_key":
                access_key = v

    if not username or not access_key:
        raise ValueError("El archivo debe tener 'username' y 'access_key'.")
    return username, access_key


# ══════════════════════════════════════════════════════════════════════════════
# DRIVER
# ══════════════════════════════════════════════════════════════════════════════

def build_bstack_options(platform: str, build_name: str, session_name: str,
                         os_version: str = "") -> Dict:
    """Arma el bloque bstack:options. Separado del driver para poder testearlo sin red.

    os_version pisa la versión del catálogo (para probar otra sin tocar DISPOSITIVOS).
    """
    if platform not in DISPOSITIVOS:
        raise ValueError(f"Plataforma '{platform}' no soportada. Opciones: {', '.join(DISPOSITIVOS)}")
    opts = dict(DISPOSITIVOS[platform]["bstack"])
    if os_version:
        opts["osVersion"] = os_version
    opts.update({
        "projectName": PROJECT_NAME,
        "buildName":   build_name,
        "sessionName": session_name,
        # El video es la evidencia de la corrida (va linkeado en el Excel de resultados).
        "video":       "true",
        # Capturas por comando: cada una es un round-trip más al device real.
        "debug":       "false",
        "networkLogs": "false",
        "consoleLogs": "errors",
    })
    return opts


def create_bs_driver(username: str, access_key: str,
                     platform: str = "iphone",
                     build_name: str = "Osocio BS",
                     session_name: str = "Form Run",
                     os_version: str = "") -> Remote:
    """Crea una sesión de BrowserStack Automate para iphone, android o mac."""
    bstack = build_bstack_options(platform, build_name, session_name, os_version)
    browser = DISPOSITIVOS[platform]["browserName"]
    options = ChromeOptions() if browser == "chrome" else SafariOptions()
    # SafariOptions trae platformName "mac" de fábrica: en el iPhone hay que pisarlo, si no
    # la sesión queda pedida como Mac y BrowserStack la resuelve sólo por el deviceName.
    options.set_capability("platformName", DISPOSITIVOS[platform]["platformName"])
    options.set_capability("bstack:options", bstack)

    cfg = ClientConfig(remote_server_addr=BS_HUB, username=username, password=access_key)
    driver = Remote(command_executor=BS_HUB, options=options, client_config=cfg)
    if platform == "mac":
        driver.maximize_window()
    driver.implicitly_wait(0)
    return driver


# ══════════════════════════════════════════════════════════════════════════════
# ESTADO DE LA SESIÓN Y VIDEO
# ══════════════════════════════════════════════════════════════════════════════

def status_script(passed: bool, reason: str = "") -> str:
    """Comando de BrowserStack para marcar la sesión (equivalente a lambda-status)."""
    payload = {
        "action": "setSessionStatus",
        "arguments": {"status": "passed" if passed else "failed", "reason": (reason or "")[:255]},
    }
    return "browserstack_executor: " + json.dumps(payload)


def mark_bs_status(driver, passed: bool = True, reason: str = "") -> None:
    try:
        driver.execute_script(status_script(passed, reason))
    except Exception:
        pass


def fetch_bs_session(session_id: str, username: str, access_key: str) -> Dict[str, str]:
    """Datos de la sesión terminada vía API REST.

    Devuelve (vacíos si la API no responde):
      video_url   — link al video grabado: es la evidencia que va al Excel de resultados
      device / os_version / status
    """
    vacio = {"video_url": "", "device": "", "os_version": "", "status": ""}
    if not session_id:
        return vacio
    try:
        import requests
        resp = requests.get(f"{BS_API}/sessions/{session_id}.json",
                            auth=(username, access_key), timeout=15)
        data = resp.json().get("automation_session", {}) or {}
    except Exception:
        return vacio
    return {k: str(data.get(k) or "") for k in vacio}
