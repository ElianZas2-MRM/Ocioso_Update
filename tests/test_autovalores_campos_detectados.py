"""Tests de utils.autovalores_campos_detectados.AutovaloresCamposDetectados.

Siembra valores de prueba en json/ids_dinamicos.json para campos detectados en
vivo sin mapeo ni dato en el Excel. El Excel siempre gana: si el campo tiene
valor en el Excel no se siembra nada.
"""
import json
import os
import re

import pytest

from utils.autovalores_campos_detectados import AutovaloresCamposDetectados


REGLAS_SINTETICAS = {
    "url": "ar",
    "fields": {
        "NOMBRE": {
            "campo": "firstname",
            "element_id": "firstname",
            "regex_full": r"^[a-zA-Z ]{2,50}$",
            "dropdown": False,
        },
        "PLACA": {
            "campo": "plate",
            "element_id": "vehicle-plate",
            "regex_full": r"^[A-Za-z0-9]{5}$",
            "dropdown": False,
        },
        "PROVINCIA": {
            "campo": "region",
            "element_id": "region",
            "regex_full": "",
            "dropdown": True,
        },
    },
}


@pytest.fixture
def json_dir(tmp_path):
    ruta = tmp_path / "json"
    ruta.mkdir()
    with open(ruta / "field_validation_rules_argentina.json", "w", encoding="utf-8") as fh:
        json.dump(REGLAS_SINTETICAS, fh, ensure_ascii=False)
    return str(ruta)


@pytest.fixture
def auto(json_dir):
    return AutovaloresCamposDetectados("Argentina", json_dir=json_dir)


def _leer_store(json_dir):
    with open(os.path.join(json_dir, "ids_dinamicos.json"), encoding="utf-8") as fh:
        return json.load(fh)


# --- lookup de reglas ---

def test_regex_para_encuentra_por_element_id_campo_y_key(auto):
    assert auto.regex_para("vehicle-plate") == r"^[A-Za-z0-9]{5}$"
    assert auto.regex_para("firstname") == r"^[a-zA-Z ]{2,50}$"
    assert auto.regex_para("NOMBRE") == r"^[a-zA-Z ]{2,50}$"
    assert auto.regex_para("PlAcA") == r"^[A-Za-z0-9]{5}$"


def test_regex_para_campo_desconocido_devuelve_vacio(auto):
    assert auto.regex_para("no-existe") == ""


def test_regex_para_sin_archivo_de_reglas_no_revienta(tmp_path):
    a = AutovaloresCamposDetectados("Argentina", json_dir=str(tmp_path))
    assert a.regex_para("firstname") == ""
    assert a.valores_para_campo("firstname", "text") == []


# --- valores_para_campo ---

def test_valores_texto_con_regex_conocido(auto):
    valores = auto.valores_para_campo("vehicle-plate", "text")
    assert 1 <= len(valores) <= 2
    for v in valores:
        assert re.fullmatch(r"^[A-Za-z0-9]{5}$", v)


def test_valores_texto_campo_desconocido_vacio(auto):
    assert auto.valores_para_campo("campo-raro", "text") == []


def test_excel_siempre_gana(auto):
    assert auto.valores_para_campo("vehicle-plate", "text", valor_excel="AB123") == []
    assert auto.valores_para_campo("region", "select",
                                   opciones_dropdown=["Buenos Aires", "Cordoba"],
                                   valor_excel="Cordoba") == []


def test_valores_dropdown_elige_dos_opciones(auto):
    opts = ["Buenos Aires", "Cordoba", "Santa Fe", "Mendoza"]
    valores = auto.valores_para_campo("region", "select", opciones_dropdown=opts)
    assert len(valores) == 2
    assert set(valores).issubset(set(opts))


def test_valores_dropdown_sin_opciones_vivas_vacio(auto):
    assert auto.valores_para_campo("region", "select", opciones_dropdown=[]) == []


# --- sembrar / persistencia ---

def test_sembrar_crea_entries_scopeadas_al_pais(auto, json_dir):
    sembrado = auto.sembrar([
        {"id": "vehicle-plate", "type": "text", "label": "Placa"},
        {"id": "region", "type": "select", "opciones": ["Buenos Aires", "Cordoba", "Santa Fe"]},
    ])
    assert set(sembrado) == {"vehicle-plate", "region"}

    store = _leer_store(json_dir)
    assert store["version"] == 2
    entries = {e["id"]: e for e in store["entries"]}
    assert entries["vehicle-plate"]["paises"] == ["Argentina"]
    assert entries["region"]["paises"] == ["Argentina"]
    # 2 variantes -> lista; 1 variante -> escalar
    assert isinstance(entries["region"]["valor"], list) and len(entries["region"]["valor"]) == 2


def test_sembrar_no_pisa_entry_existente_del_mismo_pais(auto, json_dir):
    ruta = os.path.join(json_dir, "ids_dinamicos.json")
    with open(ruta, "w", encoding="utf-8") as fh:
        json.dump({"version": 2, "entries": [
            {"id": "vehicle-plate", "valor": "MICHO", "paises": ["Argentina"]},
        ]}, fh)

    sembrado = auto.sembrar([{"id": "vehicle-plate", "type": "text"}])
    assert "vehicle-plate" not in sembrado

    store = _leer_store(json_dir)
    plate = [e for e in store["entries"] if e["id"] == "vehicle-plate"]
    assert len(plate) == 1 and plate[0]["valor"] == "MICHO"


def test_sembrar_preserva_dependencies_y_entries_ajenas(auto, json_dir):
    ruta = os.path.join(json_dir, "ids_dinamicos.json")
    with open(ruta, "w", encoding="utf-8") as fh:
        json.dump({
            "version": 2,
            "entries": [{"id": "otro-campo", "valor": "X", "paises": []}],
            "dependencies": [{"padre": "region", "hijo": "city"}],
        }, fh)

    auto.sembrar([{"id": "vehicle-plate", "type": "text"}])

    store = _leer_store(json_dir)
    ids = {e["id"] for e in store["entries"]}
    assert {"otro-campo", "vehicle-plate"}.issubset(ids)
    assert store["dependencies"] == [{"padre": "region", "hijo": "city"}]


def test_sembrar_sin_candidatos_es_noop(auto, json_dir):
    assert auto.sembrar([{"id": "campo-inexistente", "type": "text"}]) == {}
    assert not os.path.exists(os.path.join(json_dir, "ids_dinamicos.json"))


def test_sembrar_label_en_blanco_cae_al_id(auto, json_dir):
    auto.sembrar([{"id": "vehicle-plate", "type": "text", "label": "   "}])
    store = _leer_store(json_dir)
    entry = next(e for e in store["entries"] if e["id"] == "vehicle-plate")
    assert entry["nombre_campo"] == "vehicle-plate"


def test_sembrar_ignora_campos_con_valor_excel(auto, json_dir):
    sembrado = auto.sembrar([
        {"id": "vehicle-plate", "type": "text", "valor_excel": "AB123"},
    ])
    assert sembrado == {}
