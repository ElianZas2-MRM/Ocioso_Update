"""
browser_manager.py — Fábrica de browsers Selenium (Chrome, Firefox, Edge).
Crea y configura el navegador con las opciones correctas (headless, tamaño de ventana,
anti-detección de bots). Usa solo los drivers locales de la carpeta /drivers/.
"""
import os
import sys
import threading
import subprocess
from selenium import webdriver

# ── Registro de PIDs de drivers activos ───────────────────────────────────────
_active_pids: list = []
_pids_lock = threading.Lock()


def _reg_pid(service):
    """Registra el PID del proceso driver (chromedriver, geckodriver, etc.)."""
    try:
        pid = service.process.pid
        with _pids_lock:
            if pid not in _active_pids:
                _active_pids.append(pid)
    except Exception:
        pass


def clear_active_drivers():
    """Limpia la lista de PIDs (llamar al iniciar una nueva ejecución)."""
    with _pids_lock:
        _active_pids.clear()


def kill_active_drivers():
    """Mata los procesos driver + browser (árbol de procesos) que abrió esta app."""
    _NO_WIN = 0x08000000  # CREATE_NO_WINDOW
    with _pids_lock:
        pids = list(_active_pids)
        _active_pids.clear()
    for pid in pids:
        try:
            subprocess.run(
                ["taskkill", "/F", "/T", "/PID", str(pid)],
                capture_output=True,
                creationflags=_NO_WIN,
            )
        except Exception:
            pass
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.firefox.service import Service as FirefoxService

def _resolve_driver(local_name, drivers_dir):
    """Busca el driver en la carpeta local /drivers/. No descarga nada de internet."""
    local = os.path.join(drivers_dir, local_name)
    if os.path.exists(local):
        return local
    raise FileNotFoundError(
        f"No se encontró el driver '{local_name}' en la carpeta /drivers/. "
        "Descargá el driver manualmente desde el sitio oficial y colocálo en /drivers/."
    )


class BrowserManager:
    """Gestiona la configuración y creación de navegadores."""

    COMMON_BROWSER_ARGS = [
        "--force-device-scale-factor=1",
        "--disable-gpu",
        "--hide-scrollbars",
        "--disable-dev-shm-usage",
    ]
    
    @staticmethod
    def _apply_common_args(options):
        """Aplica argumentos comunes a las opciones del navegador"""
        for arg in BrowserManager.COMMON_BROWSER_ARGS:
            options.add_argument(arg)
    
    @staticmethod
    def create_browser(browser_type="chrome", viewport="fullscreen", headless=False, background=True):
        """
        Crea y configura un navegador según los parámetros.
        background=True (default): el browser arranca fuera de pantalla para no molestar al usuario.
        background=False: el browser abre en pantalla normalmente.
        """
        browser_type = browser_type.lower()

        if browser_type == "chrome":
            return BrowserManager._create_chrome(viewport, headless, background)
        elif browser_type == "firefox":
            return BrowserManager._create_firefox(viewport, headless, background)
        elif browser_type == "edge":
            return BrowserManager._create_edge(viewport, headless, background)
        else:
            raise ValueError(f"Navegador no soportado: {browser_type}")
    
    @staticmethod
    def _get_drivers_dir():
        try:
            from utils.paths import DRIVERS_DIR
            return DRIVERS_DIR
        except ImportError:
            if getattr(sys, 'frozen', False):
                base = os.path.dirname(sys.executable)
            else:
                base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            return os.path.normpath(os.path.join(base, 'drivers'))

    @staticmethod
    def _get_driver_path(driver_name):
        return os.path.join(BrowserManager._get_drivers_dir(), driver_name)

    @staticmethod
    def _create_driver_with_message(create_fn, driver_name):
        """Crea el driver y traduce errores comunes de versión a mensajes claros"""
        try:
            return create_fn()
        except Exception as e:
            error_text = str(e).lower()
            if "expected browser binary location" in error_text or "no 'moz:firefoxoptions.binary' capability provided" in error_text:
                raise Exception(
                    "Firefox no encontrado en esta PC.\n"
                    "Instalá Mozilla Firefox o configurá la variable FIREFOX_BINARY con la ruta de firefox.exe"
                ) from e
            if "session not created" in str(e).lower():
                raise Exception(
                    f"Driver desactualizado para {driver_name}.\n"
                    "Descargá la versión correcta del driver y reemplazálo en la carpeta /drivers/."
                ) from e
            raise

    @staticmethod
    def _get_firefox_binary_path():
        """Resuelve la ruta de Firefox en Windows (env, rutas comunes y registro)."""
        env_binary = os.environ.get("FIREFOX_BINARY")
        if env_binary and os.path.exists(env_binary):
            return env_binary

        if os.name != "nt":
            return None

        common_paths = [
            os.path.join(os.environ.get("ProgramFiles", ""), "Mozilla Firefox", "firefox.exe"),
            os.path.join(os.environ.get("ProgramFiles(x86)", ""), "Mozilla Firefox", "firefox.exe"),
        ]
        for path in common_paths:
            if path and os.path.exists(path):
                return path

        try:
            import winreg

            registry_keys = [
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\firefox.exe"),
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\App Paths\\firefox.exe"),
                (winreg.HKEY_CURRENT_USER, r"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\firefox.exe"),
            ]

            for hive, key_path in registry_keys:
                try:
                    with winreg.OpenKey(hive, key_path) as key:
                        value, _ = winreg.QueryValueEx(key, None)
                        if value and os.path.exists(value):
                            return value
                except OSError:
                    continue
        except Exception:
            pass

        return None
    
    @staticmethod
    def _fix_headless_user_agent(driver):
        """
        En headless, Chrome/Edge se anuncian como "HeadlessChrome"/"HeadlessEdg" en el
        user-agent. chevrolet.com.br detecta eso y NO carga el iframe del formulario (queda
        solo el data-src, sin src) → la app no encontraba ningún iframe y no llenaba nada.

        Le saca el "Headless" al UA real del navegador (sin hardcodear una versión, que se
        desactualizaría). Firefox no lo necesita: su UA headless es idéntico al normal.
        """
        try:
            ua = driver.execute_script("return navigator.userAgent") or ""
            if "Headless" not in ua:
                return
            clean = ua.replace("HeadlessChrome", "Chrome").replace("HeadlessEdg", "Edg")
            driver.execute_cdp_cmd("Network.setUserAgentOverride", {"userAgent": clean})
        except Exception:
            pass

    @staticmethod
    def _create_chrome(viewport, headless, background=True):
        options = ChromeOptions()
        # "eager" devuelve el control apenas el DOM está listo, sin esperar a que
        # terminen de cargar trackers/pixels de terceros (doubleclick, demdex, etc.)
        # que pueden tardar decenas de segundos y no afectan si el form ya está en el DOM.
        options.page_load_strategy = "eager"

        if headless:
            # --start-maximized es ignorado en headless; forzar tamaño explícito
            if viewport == "fullscreen":
                options.add_argument("--window-size=1920,1080")
            else:
                width, height = viewport.split('x')
                options.add_argument(f"--window-size={width},{height}")
            options.add_argument("--headless")
            # Reducir huellas de automatización para que el JS de la página no bloquee el formulario
            options.add_argument("--disable-blink-features=AutomationControlled")
            try:
                options.add_experimental_option("excludeSwitches", ["enable-automation"])
                options.add_experimental_option("useAutomationExtension", False)
            except Exception:
                pass
        else:
            if viewport == "fullscreen":
                options.add_argument("--window-size=1366,768")
            else:
                width, height = viewport.split('x')
                options.add_argument(f"--window-size={width},{height}")
            if background:
                # Fuera de pantalla: no roba el foco al usuario
                options.add_argument("--window-position=10000,0")

        # Aplicar argumentos comunes
        BrowserManager._apply_common_args(options)

        driver_path = _resolve_driver(
            'chromedriver.exe',
            BrowserManager._get_drivers_dir(),
        )
        service = ChromeService(driver_path)
        driver = BrowserManager._create_driver_with_message(
            lambda: webdriver.Chrome(service=service, options=options),
            "chromedriver.exe",
        )
        if headless:
            # Ocultar navigator.webdriver para evitar detección anti-bot
            try:
                driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
                    "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
                })
            except Exception:
                pass
            BrowserManager._fix_headless_user_agent(driver)
        _reg_pid(service)
        return driver

    @staticmethod
    def _create_firefox(viewport, headless, background=True):
        options = FirefoxOptions()
        options.page_load_strategy = "eager"

        firefox_binary = BrowserManager._get_firefox_binary_path()
        if firefox_binary:
            options.binary_location = firefox_binary
        elif os.name == "nt":
            raise FileNotFoundError(
                "No se encontró Firefox en esta PC. Instalá Mozilla Firefox o definí FIREFOX_BINARY con la ruta de firefox.exe"
            )

        if headless:
            options.add_argument("--headless")
            if viewport == "fullscreen":
                options.add_argument("--width=1920")
                options.add_argument("--height=1080")
            else:
                width, height = viewport.split('x')
                options.add_argument(f"--width={width}")
                options.add_argument(f"--height={height}")

        # Aplicar argumentos comunes (excepto --disable-dev-shm-usage que no aplica a Firefox)
        for arg in BrowserManager.COMMON_BROWSER_ARGS:
            if arg != "--disable-dev-shm-usage":
                options.add_argument(arg)

        driver_path = _resolve_driver(
            'geckodriver.exe',
            BrowserManager._get_drivers_dir(),
        )
        service = FirefoxService(driver_path)
        driver = BrowserManager._create_driver_with_message(
            lambda: webdriver.Firefox(service=service, options=options),
            "geckodriver.exe",
        )
        _reg_pid(service)
        # Aplicar tamaño después de crear el driver
        if headless:
            if viewport == "fullscreen":
                driver.set_window_size(1920, 1080)
            else:
                width, height = viewport.split('x')
                driver.set_window_size(int(width), int(height))
        else:
            if viewport == "fullscreen":
                driver.set_window_size(1366, 768)
            else:
                width, height = viewport.split('x')
                driver.set_window_size(int(width), int(height))
            if background:
                try:
                    driver.set_window_position(10000, 0)
                except Exception:
                    pass

        return driver

    @staticmethod
    def _create_edge(viewport, headless, background=True):
        options = EdgeOptions()
        options.page_load_strategy = "eager"

        if headless:
            if viewport == "fullscreen":
                options.add_argument("--window-size=1920,1080")
            else:
                width, height = viewport.split('x')
                options.add_argument(f"--window-size={width},{height}")
            options.add_argument("--headless")
        else:
            if viewport == "fullscreen":
                options.add_argument("--window-size=1366,768")
            else:
                width, height = viewport.split('x')
                options.add_argument(f"--window-size={width},{height}")
            if background:
                options.add_argument("--window-position=10000,0")

        # Aplicar argumentos comunes
        BrowserManager._apply_common_args(options)

        driver_path = _resolve_driver(
            'msedgedriver.exe',
            BrowserManager._get_drivers_dir(),
        )
        service = EdgeService(driver_path)
        driver = BrowserManager._create_driver_with_message(
            lambda: webdriver.Edge(service=service, options=options),
            "msedgedriver.exe",
        )
        if headless:
            BrowserManager._fix_headless_user_agent(driver)
        _reg_pid(service)
        return driver