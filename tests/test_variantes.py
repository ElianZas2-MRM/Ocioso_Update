"""Una variante es el mismo mercado corrido contra otro tipo de formulario.

Hasta ahora habia una sola (`_T3`, los formularios 2.0 de Adobe AEM) y el codigo aprovecho
eso para tomar un atajo: preguntaba "¿tiene sufijo?" y respondia "entonces es T3".

    self.ES_T3 = bool(config.get('excel_suffix'))        # base_form_filler
    _es_t3 = bool(excel_suffix)                          # autonomous_runner
    "pais": _etiqueta_t3(pais) if excel_suffix else pais  # autonomous_runner

Con una sola variante el atajo funcionaba. Con dos se rompe, y de la peor manera: en
silencio. Cadillac (`_CDC`) es T1, pero por tener sufijo sus resultados se guardarian en
`resultados/t3/` y los reportes lo etiquetarian como T3.

El otro problema que fijan estos tests es la duplicacion. El nombre del archivo de una
corrida se armaba en DOS lugares que tenian que coincidir a mano, y el docstring de uno de
ellos documenta que ya se desincronizaron una vez:

    "Buscar con el prefijo equivocado hacia que el ejecutor nunca detectara los resultados
     y, por lo tanto, nunca enviara el email consolidado."
"""
import os

import pytest

from osocio.core import variantes


# --- que variante es -------------------------------------------------------------------

def test_t3_es_t3():
    assert variantes.es_t3(variantes.T3)


@pytest.mark.parametrize("sufijo", ["", None, variantes.CDC, "_OTRA"])
def test_lo_demas_no_es_t3(sufijo):
    """La regresion: `_CDC` tiene sufijo pero NO es T3."""
    assert not variantes.es_t3(sufijo)


@pytest.mark.parametrize("escrito", ["_t3", "_T3", " _T3 "])
def test_no_importa_como_se_escriba(escrito):
    assert variantes.es_t3(escrito)


# --- como se llama en los reportes -----------------------------------------------------

def test_sin_variante_es_el_mercado_a_secas():
    assert variantes.etiqueta_de_corrida("Brasil") == "Brasil"
    assert variantes.etiqueta_de_corrida("Brasil", "") == "Brasil"


def test_t3_lleva_el_mercado_adelante():
    """Antes Brasil T3 se llamaba "CADILLAC BR", que era el enredo que hay que deshacer:
    Cadillac no es el T3 de Brasil, es otra cosa y ahora tiene su propia variante."""
    assert variantes.etiqueta_de_corrida("Brasil", variantes.T3) == "Brasil T3"
    assert variantes.etiqueta_de_corrida("Peru", variantes.T3) == "Peru T3"


def test_cadillac_se_llama_cdc():
    assert variantes.etiqueta_de_corrida("Brasil", variantes.CDC) == "CDC"


# --- el nombre de los archivos de una corrida ------------------------------------------

def test_una_corrida_normal():
    excel, ss = variantes.basenames_de_corrida("Brasil", "chrome")
    assert excel == "resultados_Brasil_Chrome"
    assert ss == "screenshots_Brasil_Chrome"


def test_una_corrida_programada_usa_otro_prefijo():
    """Lo programado se distingue de lo que corriste a mano."""
    excel, ss = variantes.basenames_de_corrida("Brasil", "chrome", programada=True)
    assert excel == "Automatizacion_Brasil_Chrome"
    assert ss == "Automatizacion_screenshots_Brasil_Chrome"


@pytest.mark.parametrize("navegador, esperado", [
    ("chrome", "Chrome"), ("firefox", "Firefox"), ("edge", "Edge"),
])
def test_el_dispositivo_va_en_el_nombre(navegador, esperado):
    excel, _ = variantes.basenames_de_corrida("Peru", navegador)
    assert excel.endswith(f"_{esperado}")


def test_la_variante_va_en_el_nombre():
    """Sin esto, una corrida de CDC escribe `resultados_Brasil_Chrome1.xlsx` y se pisa en
    la numeracion con las corridas normales de Brasil, en la misma carpeta."""
    normal, _ = variantes.basenames_de_corrida("Brasil", "chrome")
    cdc, _ = variantes.basenames_de_corrida("Brasil", "chrome", sufijo=variantes.CDC)
    t3, _ = variantes.basenames_de_corrida("Brasil", "chrome", sufijo=variantes.T3)

    assert normal != cdc != t3 and normal != t3, "las tres tienen que dar nombres distintos"
    assert "CDC" in cdc
    assert "T3" in t3


def test_sin_navegador_no_queda_un_guion_colgando():
    excel, _ = variantes.basenames_de_corrida("Brasil", "")
    assert not excel.endswith("_")


# --- que los dos lugares que arman el nombre sigan coincidiendo ------------------------

def _filler(excel_suffix="", browser="chrome", programada=False):
    from osocio.core.base_form_filler import BaseFormFiller

    config = {
        "pais": "Brasil", "browser": browser, "is_scheduled": programada,
        "excel_file": "x.xlsx", "excel_suffix": excel_suffix,
        "wait_kits_after_model": 0, "auto_step_max_iterations": 1,
        "dependency_dropdown_timeout": 1, "dependency_dropdown_poll_interval": 1,
        "dependency_selection_retries": 1,
    }
    f = object.__new__(BaseFormFiller)
    BaseFormFiller.__init__(f, config)
    return f


@pytest.mark.parametrize("sufijo", ["", "_T3", "_CDC"])
@pytest.mark.parametrize("programada", [False, True])
def test_el_motor_nombra_igual_que_la_pieza_compartida(sufijo, programada):
    """El motor ESCRIBE los archivos. Si se aparta, nadie mas los encuentra."""
    f = _filler(sufijo, programada=programada)
    excel, ss = variantes.basenames_de_corrida(
        "Brasil", "chrome", programada=programada, sufijo=sufijo)
    assert f.RESULTADOS_BASENAME == excel
    assert f.SCREENSHOT_BASENAME == ss


@pytest.mark.parametrize("navegador", ["chrome", "firefox", "edge"])
def test_el_ejecutor_autonomo_busca_con_el_mismo_nombre(navegador):
    """El ejecutor autonomo BUSCA los archivos para adjuntarlos al mail consolidado.

    Es la desincronizacion que ya paso una vez: buscaba con el prefijo equivocado y el
    mail nunca salia.
    """
    from osocio import autonomous_runner

    excel, ss = autonomous_runner._basenames_resultados("Brasil", navegador)
    esperado = variantes.basenames_de_corrida("Brasil", navegador, programada=True)
    assert (excel, ss) == esperado


# --- y lo que motiva todo: donde caen los resultados ----------------------------------

def test_una_corrida_t3_va_a_la_carpeta_t3():
    from osocio import paths

    f = _filler("_T3")
    assert f.ES_T3 is True
    assert os.path.normpath(f.RESULTADOS_DIR) == os.path.normpath(paths.RESULTS_T3_DIR)


def test_una_corrida_cdc_va_a_t1_aunque_tenga_sufijo():
    """LA regresion. Cadillac es T1; con el atajo viejo terminaba en `resultados/t3/`."""
    from osocio import paths

    f = _filler("_CDC")
    assert f.ES_T3 is False
    assert os.path.normpath(f.RESULTADOS_DIR) == os.path.normpath(paths.RESULTS_T1_DIR)
