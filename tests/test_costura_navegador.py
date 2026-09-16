"""El motor recibe con qué navegar en vez de fabricarlo.

Antes, `initialize_browser()` llamaba directo a `BrowserManager.create_browser()`. Esa
línea sola era la razón por la que ninguna de las 6.378 líneas de BaseFormFiller se podía
probar: para ejercitar el motor había que abrir Chrome de verdad.

Ahora la fábrica de navegadores se inyecta. Por defecto sigue siendo la de siempre, así
que nada cambia en producción; en los tests se le pasa un doble. Es composición, y es lo
que habilita los smoke tests de cada funcionalidad.
"""
import pytest

from osocio.core.base_form_filler import BaseFormFiller
from osocio.core.browser_manager import BrowserManager
from osocio.core.generic_country_base import GenericCountryBase


class DriverFalso:
    """Lo mínimo que initialize_browser() necesita para no explotar."""

    name = "chrome-falso"

    def __init__(self):
        self.urls_visitadas = []
        self.cerrado = False

    def get(self, url):
        self.urls_visitadas.append(url)

    def find_element(self, by, value):
        raise LookupError(f"no existe: {by}={value}")

    def find_elements(self, by, value):
        return []

    def set_window_size(self, *a, **k):
        pass

    def set_window_position(self, *a, **k):
        pass

    def maximize_window(self):
        pass

    def execute_script(self, *a, **k):
        return None

    def quit(self):
        self.cerrado = True


@pytest.fixture
def fabrica_espia():
    """Devuelve (fabrica, llamadas) — la lista registra con qué se la invocó."""
    llamadas = []

    def fabrica(**kwargs):
        llamadas.append(kwargs)
        return DriverFalso()

    return fabrica, llamadas


def _pais_de_ejemplo():
    from osocio.core.country_configs import COUNTRY_CONFIGS
    if not COUNTRY_CONFIGS:
        pytest.skip("no hay países configurados en este entorno")
    return sorted(COUNTRY_CONFIGS)[0]


# --- el comportamiento por defecto no cambia ---------------------------------------

def test_por_defecto_usa_el_browser_manager_de_siempre():
    """Lo más importante de esta costura: en produccion no cambia nada."""
    form = GenericCountryBase(_pais_de_ejemplo())
    assert form._browser_factory is BrowserManager.create_browser


# --- pero ahora se puede inyectar --------------------------------------------------

def test_generic_country_base_deja_pasar_la_fabrica(fabrica_espia):
    fabrica, _ = fabrica_espia
    form = GenericCountryBase(_pais_de_ejemplo(), browser_factory=fabrica)
    assert form._browser_factory is fabrica


def test_initialize_browser_usa_la_fabrica_inyectada(fabrica_espia, tmp_path):
    fabrica, llamadas = fabrica_espia
    form = GenericCountryBase(_pais_de_ejemplo(), browser_factory=fabrica)
    form.SCREENSHOT_DIR = str(tmp_path)

    form.initialize_browser()

    assert len(llamadas) == 1, "se tenia que llamar a la fabrica exactamente una vez"
    assert isinstance(form.driver, DriverFalso), "el driver tiene que ser el doble"


def test_la_fabrica_recibe_la_configuracion_del_pais(fabrica_espia, tmp_path):
    """No alcanza con que se llame: tiene que llegarle lo que el usuario eligió."""
    fabrica, llamadas = fabrica_espia
    form = GenericCountryBase(
        _pais_de_ejemplo(), browser="firefox", viewport="600x738",
        headless=True, background=False, browser_factory=fabrica,
    )
    form.SCREENSHOT_DIR = str(tmp_path)

    form.initialize_browser()

    assert llamadas[0] == {
        "browser_type": "firefox",
        "viewport": "600x738",
        "headless": True,
        "background": False,
    }


def test_base_form_filler_tambien_acepta_la_fabrica():
    """La costura vive en BaseFormFiller, no solo en la subclase."""
    import inspect
    parametros = inspect.signature(BaseFormFiller.__init__).parameters
    assert "browser_factory" in parametros
    assert parametros["browser_factory"].default is None, (
        "tiene que ser opcional: si fuera obligatorio romperia todas las llamadas actuales"
    )
