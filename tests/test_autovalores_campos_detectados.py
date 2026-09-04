"""Tests de utils.autovalores_campos_detectados.AutovaloresCamposDetectados.

Siembra valores de prueba en json/ids_dinamicos.json para campos detectados en
vivo sin mapeo ni dato en el Excel. El Excel siempre gana: si el campo tiene
valor en el Excel no se siembra nada.
"""
import json
import os
import re

import pytest

from utils.autovalores_campos_detectados import (
    PAISES_CONOCIDOS,
    AutovaloresCamposDetectados,
)


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


def _como_lista(valor):
    """`valor` es escalar con 1 variante y lista con 2 — normalizar para poder afirmar
    sobre cada valor sin depender de cuántas variantes salieron."""
    return valor if isinstance(valor, list) else [valor]


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
    # Un id sin semántica reconocible tampoco tiene de dónde sacar un valor.
    assert a.valores_para_campo("zzz-desconocido", "text") == []


def test_sin_regla_cae_al_fallback_semantico_por_nombre(tmp_path):
    """Último recurso cuando el campo no tiene regla: inferir la forma del nombre."""
    a = AutovaloresCamposDetectados("Argentina", json_dir=str(tmp_path))
    valores = a.valores_para_campo("contract", "text", label="Número de contrato")
    assert valores
    for v in valores:
        assert v.isdigit() and not v.startswith("0")


def test_fallback_semantico_no_aplica_a_un_id_generico(tmp_path):
    a = AutovaloresCamposDetectados("Argentina", json_dir=str(tmp_path))
    assert a.valores_para_campo("field-42", "text", label="  ") == []


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
    sembrado, corregido = auto.sembrar([
        {"id": "vehicle-plate", "type": "text", "label": "Placa"},
        {"id": "region", "type": "select", "opciones": ["Buenos Aires", "Cordoba", "Santa Fe"]},
    ])
    assert set(sembrado) == {"vehicle-plate", "region"}
    assert corregido == {}

    store = _leer_store(json_dir)
    assert store["version"] == 2
    entries = {e["id"]: e for e in store["entries"]}
    assert entries["vehicle-plate"]["paises"] == ["Argentina"]
    assert entries["region"]["paises"] == ["Argentina"]
    # 2 variantes -> lista; 1 variante -> escalar
    assert isinstance(entries["region"]["valor"], list) and len(entries["region"]["valor"]) == 2


def test_sembrar_no_pisa_entry_valida_del_mismo_pais(auto, json_dir):
    """"MICHO" cumple ^[A-Za-z0-9]{5}$: es un valor cargado a mano que sigue sirviendo."""
    ruta = os.path.join(json_dir, "ids_dinamicos.json")
    with open(ruta, "w", encoding="utf-8") as fh:
        json.dump({"version": 2, "entries": [
            {"id": "vehicle-plate", "valor": "MICHO", "paises": ["Argentina"]},
        ]}, fh)

    sembrado, corregido = auto.sembrar([{"id": "vehicle-plate", "type": "text"}])
    assert "vehicle-plate" not in sembrado
    assert "vehicle-plate" not in corregido

    store = _leer_store(json_dir)
    plate = [e for e in store["entries"] if e["id"] == "vehicle-plate"]
    assert len(plate) == 1 and plate[0]["valor"] == "MICHO"


def test_sembrar_reemplaza_entry_que_no_cumple_su_regex(auto, json_dir):
    """El caso `contract` = "324": un valor cargado a mano que el form rechaza no protege
    nada, así que se reemplaza y el anterior queda en valor_previo."""
    ruta = os.path.join(json_dir, "ids_dinamicos.json")
    with open(ruta, "w", encoding="utf-8") as fh:
        json.dump({"version": 2, "entries": [
            {"id": "vehicle-plate", "valor": "AB", "paises": ["Argentina"],
             "nombre_campo": "Placa cargada a mano"},
        ]}, fh)

    sembrado, corregido = auto.sembrar([{"id": "vehicle-plate", "type": "text"}])
    assert "vehicle-plate" in corregido
    assert "vehicle-plate" not in sembrado

    store = _leer_store(json_dir)
    plate = [e for e in store["entries"] if e["id"] == "vehicle-plate"]
    assert len(plate) == 1, "se corrige in-place, no se duplica la entry"
    assert plate[0]["valor_previo"] == "AB"
    assert plate[0]["origen"] == "autovalor_corregido"
    assert plate[0]["nombre_campo"] == "Placa cargada a mano"
    for v in _como_lista(plate[0]["valor"]):
        assert re.fullmatch(r"^[A-Za-z0-9]{5}$", v)


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
    assert auto.sembrar([{"id": "campo-inexistente", "type": "text"}]) == ({}, {})
    assert not os.path.exists(os.path.join(json_dir, "ids_dinamicos.json"))


def test_sembrar_label_en_blanco_cae_al_id(auto, json_dir):
    auto.sembrar([{"id": "vehicle-plate", "type": "text", "label": "   "}])
    store = _leer_store(json_dir)
    entry = next(e for e in store["entries"] if e["id"] == "vehicle-plate")
    assert entry["nombre_campo"] == "vehicle-plate"


def test_sembrar_ignora_campos_con_valor_excel(auto, json_dir):
    sembrado, corregido = auto.sembrar([
        {"id": "vehicle-plate", "type": "text", "valor_excel": "AB123"},
    ])
    assert sembrado == {} and corregido == {}


# --- scoping por mercado (el mismo id valida distinto en cada país) ---

def _escribir_reglas(ruta, pais, regex):
    with open(ruta / f"field_validation_rules_{pais}.json", "w", encoding="utf-8") as fh:
        json.dump({"fields": {"PLACA": {"element_id": "vehicle-plate",
                                        "regex_full": regex, "dropdown": False}}}, fh)


def _json_dir_dos_paises(tmp_path):
    """Argentina exige 5 alfanuméricos para la patente; Uruguay, 8 dígitos. Los otros
    siete mercados quedan sin archivo de reglas a propósito."""
    ruta = tmp_path / "json"
    ruta.mkdir()
    _escribir_reglas(ruta, "argentina", r"^[A-Za-z0-9]{5}$")
    _escribir_reglas(ruta, "uruguay", r"^[0-9]{8}$")
    return str(ruta)


def _json_dir_todos_los_mercados(tmp_path, regex=r"^[0-9]{8}$"):
    """La misma regla estricta en los nueve, para poder afirmar sobre el caso en que un
    valor no sirve en ningún mercado."""
    ruta = tmp_path / "json"
    ruta.mkdir()
    for pais in PAISES_CONOCIDOS:
        _escribir_reglas(ruta, pais.lower(), regex)
    return str(ruta)


def test_valida_contra_la_regla_del_mercado_en_curso(tmp_path):
    json_dir = _json_dir_dos_paises(tmp_path)
    assert AutovaloresCamposDetectados("Argentina", json_dir=json_dir).regex_para(
        "vehicle-plate") == r"^[A-Za-z0-9]{5}$"
    assert AutovaloresCamposDetectados("Uruguay", json_dir=json_dir).regex_para(
        "vehicle-plate") == r"^[0-9]{8}$"


def test_entry_global_invalida_se_desdobla_y_no_rompe_los_otros_mercados(tmp_path):
    """Al llenar, los valores de todas las entries de un id se ACUMULAN (no hay precedencia
    de país sobre global), así que agregar una entry del país no alcanzaría: hay que sacar
    a la global del mercado donde su valor no valida, dejándola en los que sí."""
    json_dir = _json_dir_dos_paises(tmp_path)
    with open(os.path.join(json_dir, "ids_dinamicos.json"), "w", encoding="utf-8") as fh:
        # "MICHO" vale en Argentina (5 alfanuméricos) pero no en Uruguay (8 dígitos).
        json.dump({"version": 2, "entries": [
            {"id": "vehicle-plate", "valor": "MICHO", "paises": []},
        ]}, fh)

    auto = AutovaloresCamposDetectados("Uruguay", json_dir=json_dir)
    _, corregido = auto.sembrar([{"id": "vehicle-plate", "type": "text"}])
    assert "vehicle-plate" in corregido

    entries = _leer_store(json_dir)["entries"]
    global_ = next(e for e in entries if e["valor"] == "MICHO")
    nueva = next(e for e in entries if e["valor"] != "MICHO")
    # La global deja de serlo, pero conserva Argentina, donde su valor sí sirve.
    assert "Argentina" in global_["paises"] and "Uruguay" not in global_["paises"]
    assert nueva["paises"] == ["Uruguay"]
    for v in _como_lista(nueva["valor"]):
        assert re.fullmatch(r"^[0-9]{8}$", v)


def test_mercado_sin_reglas_conocidas_conserva_el_valor_global(tmp_path):
    """Un mercado sin archivo de reglas no aporta evidencia de que el valor esté mal, así
    que la global se mantiene ahí: la duda no alcanza para sacarla."""
    json_dir = _json_dir_dos_paises(tmp_path)
    with open(os.path.join(json_dir, "ids_dinamicos.json"), "w", encoding="utf-8") as fh:
        json.dump({"version": 2, "entries": [
            {"id": "vehicle-plate", "valor": "MICHO", "paises": []},
        ]}, fh)

    AutovaloresCamposDetectados("Uruguay", json_dir=json_dir).sembrar(
        [{"id": "vehicle-plate", "type": "text"}])

    global_ = next(e for e in _leer_store(json_dir)["entries"] if e["valor"] == "MICHO")
    assert "Brasil" in global_["paises"] and "Peru" in global_["paises"]


def test_no_se_contamina_la_regla_de_un_mercado_con_la_de_otro(tmp_path):
    """`vehicle-plate` valida distinto en cada mercado. Un valor bueno en Argentina no debe
    darse por bueno en Uruguay solo porque comparte element_id."""
    json_dir = _json_dir_dos_paises(tmp_path)
    ar = AutovaloresCamposDetectados("Argentina", json_dir=json_dir)
    uy = AutovaloresCamposDetectados("Uruguay", json_dir=json_dir)

    assert ar._valor_valido("MICHO", "vehicle-plate")
    assert not uy._valor_valido("MICHO", "vehicle-plate")
    # Y cada uno genera contra SU regla.
    for v in _como_lista(ar.valores_para_campo("vehicle-plate", "text")):
        assert re.fullmatch(r"^[A-Za-z0-9]{5}$", v)
    for v in _como_lista(uy.valores_para_campo("vehicle-plate", "text")):
        assert re.fullmatch(r"^[0-9]{8}$", v)


# --- varias reglas condicionadas al mismo element_id ---

def _json_dir_documento_condicional(tmp_path):
    """Peru real: `ci` vale 8 dígitos con DNI, 11 con RUC y 12 con Pasaporte, según lo que
    se haya elegido en el dropdown `document-type`."""
    ruta = tmp_path / "json"
    ruta.mkdir()
    with open(ruta / "field_validation_rules_peru.json", "w", encoding="utf-8") as fh:
        json.dump({"fields": {
            "DNI__2": {"element_id": "ci", "regex_full": r"^[0-9]{8}$",
                       "dependencies": [{"element_id": "document-type", "value": "DNI"}]},
            "RUT": {"element_id": "ci", "regex_full": r"^[0-9]{11}$",
                    "dependencies": [{"element_id": "document-type", "value": "RUC"}]},
            "Pasaporte": {"element_id": "ci", "regex_full": r"^[0-9]{12}$",
                          "dependencies": [{"element_id": "document-type", "value": "Pasaporte"}]},
        }}, fh, ensure_ascii=False)
    return str(ruta)


def test_element_id_disputado_no_resuelve_a_una_regla_arbitraria(tmp_path):
    """Quedarse con la primera daba por inválido un DNI correcto de 8 dígitos por medirlo
    contra la regla del RUC, y podía "corregirlo" a un valor que sólo sirve para RUC."""
    auto = AutovaloresCamposDetectados(
        "Peru", json_dir=_json_dir_documento_condicional(tmp_path))

    assert auto.regex_para("ci") == ""
    assert auto._valor_valido("17182701", "ci"), "un DNI válido no puede darse por inválido"
    assert auto.valores_para_campo("ci", "text") == [], "sin saber el tipo, no se inventa"


def test_la_key_del_campo_sigue_resolviendo_aunque_el_element_id_este_disputado(tmp_path):
    auto = AutovaloresCamposDetectados(
        "Peru", json_dir=_json_dir_documento_condicional(tmp_path))

    assert auto.regex_para("DNI__2") == r"^[0-9]{8}$"
    assert auto.regex_para("Pasaporte") == r"^[0-9]{12}$"


def test_un_element_id_con_una_sola_regla_no_se_marca_ambiguo(tmp_path):
    auto = AutovaloresCamposDetectados("Uruguay", json_dir=_json_dir_dos_paises(tmp_path))
    assert auto.regex_para("vehicle-plate") == r"^[0-9]{8}$"


# Ids que hoy tienen regex distinto segun el mercado en json/field_validation_rules_*.json.
IDS_DIVERGENTES = ["ci", "telephone", "patent", "vin", "firstname", "lastname"]


@pytest.mark.parametrize("field_id", IDS_DIVERGENTES)
def test_cada_mercado_genera_contra_su_propia_regla_real(field_id):
    """Barrido sobre los JSON reales del repo: para cada país que define el campo, el valor
    generado tiene que cumplir el regex DE ESE PAÍS. Atajo del riesgo más caro de esta
    feature — corregir un mercado con la regla de otro rompe uno que hoy anda."""
    json_dir = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "json")

    evaluados = 0
    for pais in PAISES_CONOCIDOS:
        auto = AutovaloresCamposDetectados(pais, json_dir=json_dir)
        regex = auto.regex_para(field_id)
        if not regex:
            continue
        evaluados += 1
        for v in auto.valores_para_campo(field_id, "text"):
            assert re.fullmatch(regex, v), (
                f"{v!r} no cumple la regla de {pais} para {field_id}: {regex}")

    assert evaluados >= 2, f"{field_id} deberia estar definido en varios mercados"


def test_entry_global_invalida_en_todos_los_mercados_se_reemplaza_sin_desdoblar(tmp_path):
    json_dir = _json_dir_todos_los_mercados(tmp_path)
    with open(os.path.join(json_dir, "ids_dinamicos.json"), "w", encoding="utf-8") as fh:
        # "!!" no cumple en ningún mercado: desdoblar no tendría sentido.
        json.dump({"version": 2, "entries": [
            {"id": "vehicle-plate", "valor": "!!", "paises": []},
        ]}, fh)

    auto = AutovaloresCamposDetectados("Uruguay", json_dir=json_dir)
    auto.sembrar([{"id": "vehicle-plate", "type": "text"}])

    entries = [e for e in _leer_store(json_dir)["entries"] if e["id"] == "vehicle-plate"]
    assert len(entries) == 1, "no se desdobla: no hay mercado que preservar"
    assert entries[0]["valor_previo"] == "!!"
    assert entries[0]["paises"] == []
