"""Tests de BaseFormFiller._campos_llenados_quedaron_vacios (lógica pura, sin driver).

Detecta que el form se re-renderizó/recargó después de llenarlo y borró todo (caso
evento-blazer-rs-ev Chile en Edge: tras el error de event_id el form se recargaba DESPUÉS
del relleno y el reintento se enviaba vacío).
"""
from osocio.core.base_form_filler import BaseFormFiller

vacios = BaseFormFiller._campos_llenados_quedaron_vacios


def test_todos_los_llenados_vacios_es_reset():
    dom = {"firstname": "", "lastname": "", "email": " ", "comment": "hola"}
    assert vacios(["firstname", "lastname", "email"], dom) is True


def test_alguno_con_valor_no_es_reset():
    dom = {"firstname": "Ana", "lastname": "", "email": ""}
    assert vacios(["firstname", "lastname", "email"], dom) is False


def test_un_solo_campo_presente_no_alcanza_para_opinar():
    # Puede ser un opcional o un campo que el form limpia a propósito.
    assert vacios(["firstname", "region"], {"firstname": ""}) is False


def test_campos_no_presentes_no_cuentan():
    # Recargando: el documento todavía no tiene los campos → no es "reset" aún.
    assert vacios(["firstname", "lastname"], {}) is False


def test_selects_y_campos_ajenos_se_ignoran():
    # 'region' es un select (no aparece en el snapshot de texto) y 'otro' no se llenó.
    dom = {"firstname": "", "lastname": "", "otro": ""}
    assert vacios(["firstname", "lastname", "region"], dom) is True
