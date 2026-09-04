"""Tests de utils.data_generator.valor_plausible_por_nombre.

Deduce un valor con forma razonable a partir del id/label de un campo que no está en el
mapping de ningún mercado. Es el último recurso del motor: sólo corre cuando el campo no
tiene dato en el Excel ni en IDs dinámicos, y su alternativa era el probe generico
("Carlos" / "12345678"), que dejaba edades de 12345678 en el lead.

Contrato: si el nombre no dice nada reconocible devuelve "" y el motor sigue con su probe.
"""
import re

import pytest

from utils.data_generator import valor_plausible_por_nombre as valor


@pytest.mark.parametrize("field_id,patron", [
    ("age", r"^(2[1-9]|[3-5][0-9]|6[0-5])$"),          # 21-65
    ("edad", r"^(2[1-9]|[3-5][0-9]|6[0-5])$"),
    ("idade", r"^(2[1-9]|[3-5][0-9]|6[0-5])$"),
    ("shoes", r"^(3[6-9]|4[0-5])$"),                   # 36-45
    ("model_year", r"^20(1[5-9]|2[0-5])$"),            # 2015-2025
    ("street-number", r"^[1-9][0-9]{2,3}$"),
    ("group", r"^[1-5]$"),
    ("txtMilage", r"^[0-9]+000$"),
    ("kilometraje", r"^[0-9]+000$"),
])
def test_devuelve_un_valor_con_la_forma_esperada(field_id, patron):
    v = valor(field_id)
    assert re.fullmatch(patron, v), f"{field_id} devolvió {v!r}"


def test_fecha_de_nacimiento_es_de_alguien_mayor_de_edad():
    for _ in range(20):
        v = valor("birthday")
        assert re.fullmatch(r"\d{2}/\d{2}/\d{4}", v), v
        anio = int(v.split("/")[-1])
        assert 1970 <= anio <= 2004, f"año fuera de rango: {anio}"


@pytest.mark.parametrize("field_id", ["street", "rua", "direccion", "endereco"])
def test_direccion_devuelve_un_nombre_de_calle_no_una_cadena_al_azar(field_id):
    v = valor(field_id)
    assert v and " " in v, f"{field_id} devolvió {v!r}"
    assert not v.isdigit()


@pytest.mark.parametrize("field_id,esperado", [
    ("comment", "texto con espacios"),
    ("mensaje", "texto con espacios"),
    ("institution", "texto con espacios"),
])
def test_campos_de_texto_libre_devuelven_algo_legible(field_id, esperado):
    v = valor(field_id)
    assert v and " " in v, f"{field_id} devolvió {v!r} ({esperado})"


def test_instagram_tiene_forma_de_usuario():
    assert valor("instagram").startswith("@")


# --- el bug que motivó el matcheo por palabra entera ---

def test_txt_milage_es_kilometraje_y_no_edad():
    """"txtMilage" contiene "age": como substring devolvía una edad (39) en un campo de
    kilometraje. Las claves cortas se exigen como palabra entera."""
    for _ in range(15):
        v = int(valor("txtMilage"))
        assert v >= 5000, f"parece una edad, no kilometraje: {v}"


@pytest.mark.parametrize("field_id", ["package", "message", "language", "manage", "storage"])
def test_una_palabra_que_contiene_age_no_se_lee_como_edad(field_id):
    v = valor(field_id)
    if not v:
        return
    assert not re.fullmatch(r"\d{2}", v), f"{field_id} devolvió una edad: {v!r}"


# --- contrato: no inventar ---

@pytest.mark.parametrize("field_id", ["xyz123", "campo-raro", "firstname", "lastname", ""])
def test_sin_semantica_reconocible_devuelve_vacio(field_id):
    assert valor(field_id) == ""


def test_usa_tambien_el_label_no_solo_el_id():
    assert valor("f_12", "Fecha de nacimiento")
    assert valor("f_13", "Kilometraje del vehículo")


def test_no_revienta_con_entradas_raras():
    for basura in (None, "", "   ", "?" * 200, 12345):
        assert isinstance(valor(basura), str)
