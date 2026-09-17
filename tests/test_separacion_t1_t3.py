"""Los formularios T1 y T3 guardan sus resultados por separado.

T3 son los formularios 2.0 de Adobe AEM. Se corren con un Excel propio
(`Lead_information_..._T3.xlsx`) y hasta ahora escribían los resultados **en la misma
carpeta** que los T1: `resultados_Argentina_Chrome1.xlsx` y `..._Chrome2.xlsx`, sin nada
que dijera cuál era cuál.

El sufijo `_T3` ya existía, pero solo decidía qué Excel leer. Ahora también decide dónde
caen los resultados.
"""
import os

import pytest

from osocio import paths


def test_t1_y_t3_no_comparten_carpeta():
    """Lo que motiva todo esto: si apuntaran al mismo lado, no se separa nada."""
    assert os.path.normpath(paths.RESULTS_T1_DIR) != os.path.normpath(paths.RESULTS_T3_DIR)


@pytest.mark.parametrize("nombre", ["RESULTS_T1_DIR", "RESULTS_T3_DIR"])
def test_las_dos_cuelgan_de_resultados(nombre):
    """No vuelven a colgar de la raíz: ya pasó con las cuatro carpetas de LambdaTest."""
    ruta = os.path.normpath(getattr(paths, nombre))
    assert ruta.startswith(os.path.normpath(paths.RESULTS_DIR) + os.sep)


@pytest.mark.parametrize("es_t3, esperada", [
    (False, "RESULTS_T1_DIR"),
    (True, "RESULTS_T3_DIR"),
])
def test_results_dir_para_elige_bien(es_t3, esperada):
    assert paths.results_dir_para(es_t3) == getattr(paths, esperada)


def test_acepta_cualquier_cosa_que_se_evalue_como_booleano():
    """El motor le pasa `bool(config.get('excel_suffix'))`, que puede ser "" o "_T3"."""
    assert paths.results_dir_para("") == paths.RESULTS_T1_DIR
    assert paths.results_dir_para("_T3") == paths.RESULTS_T3_DIR
    assert paths.results_dir_para(None) == paths.RESULTS_T1_DIR


# --- el motor escribe donde corresponde ---------------------------------------------

def _filler(excel_suffix):
    """Un BaseFormFiller armado sin navegador, solo para ver a dónde apunta."""
    from osocio.core.base_form_filler import BaseFormFiller

    config = {
        "pais": "Argentina",
        "browser": "chrome",
        "excel_file": f"Lead_information_Formulario_Argentina_Chrome{excel_suffix}.xlsx",
        "excel_suffix": excel_suffix,
        "wait_kits_after_model": 0,
        "auto_step_max_iterations": 1,
        "dependency_dropdown_timeout": 1,
        "dependency_dropdown_poll_interval": 1,
        "dependency_selection_retries": 1,
    }
    f = object.__new__(BaseFormFiller)
    BaseFormFiller.__init__(f, config)
    return f


def test_una_corrida_t1_escribe_en_t1():
    filler = _filler("")
    assert filler.ES_T3 is False
    assert os.path.normpath(filler.RESULTADOS_DIR) == os.path.normpath(paths.RESULTS_T1_DIR)


def test_una_corrida_t3_escribe_en_t3():
    filler = _filler("_T3")
    assert filler.ES_T3 is True
    assert os.path.normpath(filler.RESULTADOS_DIR) == os.path.normpath(paths.RESULTS_T3_DIR)


def test_el_pais_arma_el_excel_de_t3_y_marca_la_config():
    """`GenericCountryBase` es quien traduce el sufijo en nombre de Excel y en marca."""
    from osocio.core.country_configs import COUNTRY_CONFIGS
    from osocio.core.generic_country_base import GenericCountryBase

    if not COUNTRY_CONFIGS:
        pytest.skip("no hay países configurados en este entorno")
    pais = sorted(COUNTRY_CONFIGS)[0]

    normal = GenericCountryBase(pais)
    t3 = GenericCountryBase(pais, excel_suffix="_T3")

    assert not normal.config["excel_suffix"]
    assert t3.config["excel_suffix"] == "_T3"
    assert t3.config["excel_file"].endswith("_T3.xlsx")
    assert not normal.config["excel_file"].endswith("_T3.xlsx")
    assert os.path.normpath(t3.RESULTADOS_DIR) != os.path.normpath(normal.RESULTADOS_DIR)


# --- el ejecutor autónomo tiene que BUSCAR donde se escribió -------------------------

def test_el_autonomo_cuenta_las_corridas_en_la_carpeta_correcta(tmp_path, monkeypatch):
    """Si contara en la carpeta equivocada, numeraría mal y pisaría resultados.

    Peor todavía: el mail final adjunta el archivo que encuentra por ese número, así que
    un conteo cruzado mandaría el Excel de la corrida que no es.
    """
    from osocio import autonomous_runner

    t1 = tmp_path / "t1"
    t3 = tmp_path / "t3"
    t1.mkdir()
    t3.mkdir()
    (t1 / "resultados_Argentina_Chrome1.xlsx").write_text("x", encoding="utf-8")
    (t1 / "resultados_Argentina_Chrome2.xlsx").write_text("x", encoding="utf-8")
    (t3 / "resultados_Argentina_Chrome1.xlsx").write_text("x", encoding="utf-8")

    monkeypatch.setattr(
        autonomous_runner, "results_dir_para",
        lambda es_t3: str(t3) if es_t3 else str(t1),
    )

    n_t1 = autonomous_runner.obtener_numero_mayor_existente(
        "Argentina", "excel", "resultados_Argentina_Chrome", es_t3=False)
    n_t3 = autonomous_runner.obtener_numero_mayor_existente(
        "Argentina", "excel", "resultados_Argentina_Chrome", es_t3=True)

    assert n_t1 == 2, "en t1 hay dos corridas"
    assert n_t3 == 1, "en t3 hay una sola, no tiene que ver las de t1"
