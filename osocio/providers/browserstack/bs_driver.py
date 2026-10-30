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
    if username.upper().startswith("TU_") or access_key.upper().startswith("TU_"):
        raise ValueError(
            f"Completá {os.path.basename(filepath)} con tu usuario y access key de BrowserStack "
            "(https://www.browserstack.com/accounts/settings). Si no tenés usuario, pedí que te "
            "inviten a la organización de BrowserStack.")
    return username, access_key


# ══════════════════════════════════════════════════════════════════════════════
# TÚNEL LOCAL (VPN)
# ══════════════════════════════════════════════════════════════════════════════

def tunel_local() -> Tuple[bool, str]:
    """(activo, identificador) del túnel de BrowserStack Local.

    Con el túnel, el navegador de la nube sale a internet por la PC donde corre
    BrowserStackLocal.exe, o sea por su VPN: así llega a sitios que solo se ven desde la
    red de la empresa (ej. stage). El binario lo abre la persona a mano, la app no lanza
    nada: `BrowserStackLocal.exe --key <access key> --force-local --local-identifier X`.
    Sin --force-local el túnel solo lleva hosts locales/privados: un dominio público que
    solo deja entrar desde la VPN (stage) se abre igual desde la nube y da 403. El
    identificador evita usar el túnel de otro equipo (la cuenta es compartida).
    Se prende en json/config_global.json:
    {"browserstack": {"local": true, "local_identifier": "opcional"}}. Apagado por
    defecto: con el túnel pedido y cerrado, BrowserStack rechaza la sesión.
    """
    try:
        from osocio.paths import JSON_DIR
        with open(os.path.join(JSON_DIR, "config_global.json"), "r", encoding="utf-8") as f:
            cfg = json.load(f).get("browserstack") or {}
    except Exception:
        return False, ""
    return bool(cfg.get("local", False)), str(cfg.get("local_identifier") or "").strip()


def mensaje_tunel_cerrado(identificador: str = "") -> str:
    extra = f" --local-identifier {identificador}" if identificador else ""
    return ("BrowserStack no encuentra el túnel local (la VPN). Abrí BrowserStackLocal.exe en "
            "una terminal y dejalo corriendo:\n"
            f"    BrowserStackLocal.exe --key <tu access key> --force-local{extra}\n"
            "o apagá el túnel en json/config_global.json -> \"browserstack\": {\"local\": false}.")


# ══════════════════════════════════════════════════════════════════════════════
# DRIVER
# ══════════════════════════════════════════════════════════════════════════════

def build_bstack_options(platform: str, build_name: str, session_name: str,
                         os_version: str = "", local=None) -> Dict:
    """Arma el bloque bstack:options. Separado del driver para poder testearlo sin red.

    os_version pisa la versión del catálogo (para probar otra sin tocar DISPOSITIVOS).
    local: (activo, identificador) del túnel; None = lo que diga config_global.json.
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
    activo, identificador = tunel_local() if local is None else local
    if activo:
        opts["local"] = "true"
        if identificador:
            opts["localIdentifier"] = identificador
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
    if bstack.get("local") == "true":
        # Por el túnel el tráfico pasa por la red de la empresa, que inspecciona HTTPS y
        # firma los sitios con su propio certificado: el device de la nube no lo conoce y
        # mostraba "This Connection Is Not Private". Probado: iPhone y Android lo aceptan
        # con esto; Safari de Mac lo ignora.
        options.set_capability("acceptInsecureCerts", True)
    # Igual que los navegadores locales: no esperar trackers/píxeles de terceros. Con
    # "normal", en un device real la landing puede tardar más que el timeout de red.
    options.page_load_strategy = "eager"

    # Devices reales: algunos comandos tardan más que en un navegador local.
    cfg = ClientConfig(remote_server_addr=BS_HUB, username=username, password=access_key,
                       timeout=300)
    try:
        driver = Remote(command_executor=BS_HUB, options=options, client_config=cfg)
    except Exception as e:
        if bstack.get("local") == "true" and "local" in str(e).lower():
            raise RuntimeError(mensaje_tunel_cerrado(bstack.get("localIdentifier", ""))) from e
        raise
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


def mark_bs_status_api(session_id: str, username: str, access_key: str,
                       passed: bool, reason: str = "") -> bool:
    """Igual que mark_bs_status pero por la API REST: sirve cuando el driver ya se
    cerró (el motor de los navegadores lo cierra en su cleanup)."""
    if not session_id:
        return False
    try:
        import requests
        resp = requests.put(f"{BS_API}/sessions/{session_id}.json", auth=(username, access_key),
                            json={"status": "passed" if passed else "failed",
                                  "reason": (reason or "")[:255]}, timeout=15)
        return resp.ok
    except Exception:
        return False


def fetch_bs_session(session_id: str, username: str, access_key: str) -> Dict[str, str]:
    """Datos de la sesión terminada vía API REST.

    Devuelve (vacíos si la API no responde):
      video_url   — la evidencia que va al Excel de resultados: la página de la sesión
                    DENTRO de su build (video, logs y el build del día, con la cuenta
                    de BrowserStack). El link crudo al archivo de video vence y no lleva
                    a la corrida: queda sólo como respaldo.
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
    info = {k: str(data.get(k) or "") for k in vacio}
    info["video_url"] = str(data.get("browser_url") or "") or info["video_url"]
    return info
