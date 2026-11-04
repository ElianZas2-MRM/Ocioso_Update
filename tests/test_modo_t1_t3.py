"""Modo T1 / T3: el interruptor de arriba decide con qué Excels corre todo.

Los Excels de T1 pasan a llamarse `…_T1.xlsx` para no confundirse con los `…_T3.xlsx` en
`data/`. Los viejos sin sufijo se siguen leyendo: nadie tiene que renombrar nada.
"""

import pytest

from osocio.core import variantes
from osocio.core.variantes import (CDC, T1, T3, basenames_de_corrida, etiqueta_de_corrida,
                                   nombre_excel_leads, ruta_excel_leads, variante_del_modo)


@pytest.mark.parametrize("modo, esperado", [("T1", ""), ("T3", T3), ("t3", T3), ("", ""), (None, "")])
def test_variante_del_modo(modo, esperado):
    assert variante_del_modo(modo) == esperado


@pytest.mark.parametrize("sufijo, esperado", [
    ("", "Lead_information_Formulario_Brasil_Chrome_T1.xlsx"),
    (T1, "Lead_information_Formulario_Brasil_Chrome_T1.xlsx"),
    (T3, "Lead_information_Formulario_Brasil_Chrome_T3.xlsx"),
    (CDC, "Lead_information_Formulario_Brasil_Chrome_CDC.xlsx"),
])
def test_nombre_del_excel_a_crear(sufijo, esperado):
    assert nombre_excel_leads("Brasil", "Chrome", sufijo) == esperado


def test_t1_prefiere_el_excel_nuevo(tmp_path):
    (tmp_path / "Lead_information_Formulario_Peru_Edge.xlsx").write_text("")
    (tmp_path / "Lead_information_Formulario_Peru_Edge_T1.xlsx").write_text("")
    assert ruta_excel_leads(str(tmp_path), "Peru", "Edge").endswith("_Edge_T1.xlsx")


def test_t1_usa_el_viejo_sin_sufijo_si_es_el_unico(tmp_path):
    (tmp_path / "Lead_information_Formulario_Peru_Edge.xlsx").write_text("")
    assert ruta_excel_leads(str(tmp_path), "Peru", "Edge").endswith("Peru_Edge.xlsx")


def test_t1_sin_ningun_excel_apunta_al_nuevo(tmp_path):
    assert ruta_excel_leads(str(tmp_path), "Peru", "Edge").endswith("_Edge_T1.xlsx")


def test_t3_nunca_cae_al_excel_de_t1(tmp_path):
    (tmp_path / "Lead_information_Formulario_Peru_Edge.xlsx").write_text("")
    assert ruta_excel_leads(str(tmp_path), "Peru", "Edge", T3).endswith("_Edge_T3.xlsx")


def test_t1_no_cambia_resultados_ni_mail():
    """El `_T1` es solo para los Excels de entrada: resultados y mail quedan como siempre."""
    assert basenames_de_corrida("Brasil", "chrome", sufijo=T1) == basenames_de_corrida("Brasil", "chrome")
    assert etiqueta_de_corrida("Brasil", T1) == "Brasil"
    assert not variantes.es_t3(T1)


def test_el_motor_lee_el_excel_t1(tmp_path, monkeypatch):
    from osocio.core import generic_country_base
    import osocio.paths as paths
    monkeypatch.setattr(paths, "DATA_DIR", str(tmp_path))
    (tmp_path / "Lead_information_Formulario_Argentina_Chrome.xlsx").write_text("")
    base = generic_country_base.GenericCountryBase("Argentina")
    assert base.config["excel_file"] == "Lead_information_Formulario_Argentina_Chrome.xlsx"
    (tmp_path / "Lead_information_Formulario_Argentina_Chrome_T1.xlsx").write_text("")
    base = generic_country_base.GenericCountryBase("Argentina")
    assert base.config["excel_file"] == "Lead_information_Formulario_Argentina_Chrome_T1.xlsx"


# --- Cadillac en los dos modos -------------------------------------------------------------

def test_cadillac_sigue_el_modo():
    from osocio.core.variantes import CDC_T3, variante_cadillac
    assert variante_cadillac("T1") == CDC
    assert variante_cadillac("T3") == CDC_T3


def test_cadillac_t3_es_t3_y_tiene_su_excel():
    from osocio.core.variantes import CDC_T3
    assert variantes.es_t3(CDC_T3)
    assert not variantes.es_t3(CDC)
    assert nombre_excel_leads("Brasil", "Chrome", CDC_T3) == "Lead_information_Formulario_Brasil_Chrome_CDC_T3.xlsx"
    assert etiqueta_de_corrida("Brasil", CDC_T3) == "CDC T3"
    assert basenames_de_corrida("Brasil", "chrome", sufijo=CDC_T3)[0] == "resultados_Brasil_CDC_T3_Chrome"


def test_la_etiqueta_de_sesion_de_cadillac_t3_no_se_confunde_con_t3():
    from osocio.core.variantes import CDC_T3
    from osocio.interface.envio_leads import _variante_del_dispositivo
    assert _variante_del_dispositivo("Chrome·CDC_T3") == ("Chrome", CDC_T3)
    assert _variante_del_dispositivo("Chrome·T3") == ("Chrome", T3)
    assert _variante_del_dispositivo("Chrome·CDC") == ("Chrome", CDC)
