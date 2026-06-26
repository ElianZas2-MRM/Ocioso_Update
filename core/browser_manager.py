import os
import sys
from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.firefox.service import Service as FirefoxService

# Intentar usar webdriver-manager para descarga automática; si no está instalado
# o falla (sin internet, EXE offline), se usa el driver local de /drivers/.
try:
    from webdriver_manager.chrome import ChromeDriverManager
    from webdriver_manager.firefox import GeckoDriverManager
    from webdriver_manager.microsoft import EdgeChromiumDriverManager
    _WDM_AVAILABLE = True
except ImportError:
    _WDM_AVAILABLE = False


def _resolve_driver(wdm_fn, local_name, drivers_dir):
    """Descarga con webdriver-manager o cae al driver local si falla."""
    if _WDM_AVAILABLE:
        try:
            return wdm_fn()
        except Exception:
            pass
    local = os.path.join(drivers_dir, local_name)
    if os.path.exists(local):
        return local
    raise FileNotFoundError(
        f"No se encontró el driver '{local_name}'. "
        "Instalá webdriver-manager (`pip install webdriver-manager`) o colocá el driver en /drivers/."
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
    def create_browser(browser_type="chrome", viewport="fullscreen", headless=False):
        """
        Crea y configura un navegador según los parámetros
        """
        browser_type = browser_type.lower()
        
        if browser_type == "chrome":
            return BrowserManager._create_chrome(viewport, headless)
        elif browser_type == "firefox":
            return BrowserManager._create_firefox(viewport, headless)
        elif browser_type == "edge":
            return BrowserManager._create_edge(viewport, headless)
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
                    "El sistema intentará descargarlo automáticamente la próxima vez si tiene internet."
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
    def _create_chrome(viewport, headless):
        options = ChromeOptions()

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
                options.add_argument("--start-maximized")
            else:
                width, height = viewport.split('x')
                options.add_argument(f"--window-size={width},{height}")

        # Aplicar argumentos comunes
        BrowserManager._apply_common_args(options)

        driver_path = _resolve_driver(
            lambda: ChromeDriverManager().install() if _WDM_AVAILABLE else None,
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
        else:
            # Forzar tamaño desktop y mover fuera de pantalla para no interrumpir al usuario
            try:
                if viewport == "fullscreen":
                    driver.set_window_size(1366, 768)
                driver.set_window_position(10000, 0)
            except Exception:
                pass
        return driver
    
    @staticmethod
    def _create_firefox(viewport, headless):
        options = FirefoxOptions()

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
            lambda: GeckoDriverManager().install() if _WDM_AVAILABLE else None,
            'geckodriver.exe',
            BrowserManager._get_drivers_dir(),
        )
        service = FirefoxService(driver_path)
        driver = BrowserManager._create_driver_with_message(
            lambda: webdriver.Firefox(service=service, options=options),
            "geckodriver.exe",
        )
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
            try:
                driver.set_window_position(10000, 0)
            except Exception:
                pass

        return driver

    @staticmethod
    def _create_edge(viewport, headless):
        options = EdgeOptions()

        if headless:
            # --start-maximized es ignorado en headless; forzar tamaño explícito
            if viewport == "fullscreen":
                options.add_argument("--window-size=1920,1080")
            else:
                width, height = viewport.split('x')
                options.add_argument(f"--window-size={width},{height}")
            options.add_argument("--headless")
        else:
            if viewport == "fullscreen":
                options.add_argument("--start-maximized")
            else:
                width, height = viewport.split('x')
                options.add_argument(f"--window-size={width},{height}")

        # Aplicar argumentos comunes
        BrowserManager._apply_common_args(options)
        
        driver_path = _resolve_driver(
            lambda: EdgeChromiumDriverManager().install() if _WDM_AVAILABLE else None,
            'msedgedriver.exe',
            BrowserManager._get_drivers_dir(),
        )
        service = EdgeService(driver_path)
        driver = BrowserManager._create_driver_with_message(
            lambda: webdriver.Edge(service=service, options=options),
            "msedgedriver.exe",
        )
        if not headless:
            try:
                if viewport == "fullscreen":
                    driver.set_window_size(1366, 768)
                driver.set_window_position(10000, 0)
            except Exception:
                pass

        return driver