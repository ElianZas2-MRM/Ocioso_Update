"""Un form que tarda en aparecer no es un FORM NO INSERTO.

En Firefox el iframe de Cadillac Vistiq apareció después de las dos esperas (~17 s) y la
fila quedó como "FORM NO INSERTO" con el form ahí. Hay una tercera espera, más larga.
"""
import pytest
from selenium.common.exceptions import TimeoutException

from osocio.core import base_form_filler as bff

URL = "https://secure-developments.com/commonwealth/gm_frontend/cadillac/t1/brasil/form/vistiq?model=vistiq"


class _Iframe:
    location = {"y": 0}

    def get_attribute(self, nombre):
        return URL if nombre == "src" else None


class _Driver:
    def execute_script(self, *_a):
        return 800


@pytest.fixture
def motor(monkeypatch):
    monkeypatch.setattr(bff.time, "sleep", lambda _s: None)
    m = bff.BaseFormFiller.__new__(bff.BaseFormFiller)
    m.driver = _Driver()
    m.screenshot_manager = None
    m.pre_scroll_for_dynamic_content = lambda: None
    m._scrollear_a_iframe_sin_cargar = lambda _u: None
    return m


def _espera_que_aparece_en(intento, esperas):
    """WebDriverWait de mentira: el iframe aparece recién en el intento pedido."""
    class _Wait:
        def __init__(self, _driver, segundos):
            esperas.append(segundos)

        def until(self, _cond):
            if len(esperas) < intento:
                raise TimeoutException()
            return _Iframe()
    return _Wait


def test_form_que_aparece_en_la_tercera_espera_se_encuentra(motor, monkeypatch):
    esperas = []
    monkeypatch.setattr(bff, "WebDriverWait", _espera_que_aparece_en(3, esperas))
    assert motor.find_and_position_to_form(URL) is not None
    assert esperas == [6, 8, 12]
    assert motor._url_form_encontrado == URL


def test_form_rapido_no_espera_de_mas(motor, monkeypatch):
    esperas = []
    monkeypatch.setattr(bff, "WebDriverWait", _espera_que_aparece_en(1, esperas))
    assert motor.find_and_position_to_form(URL) is not None
    assert esperas == [6]


def test_sin_form_se_rinde_tras_las_tres_esperas(motor, monkeypatch):
    esperas = []
    monkeypatch.setattr(bff, "WebDriverWait", _espera_que_aparece_en(99, esperas))
    assert motor.find_and_position_to_form(URL) is None
    assert esperas == [6, 8, 12]
