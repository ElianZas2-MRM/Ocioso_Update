"""Tests de BaseFormFiller._decidir_marca_checkbox (lógica pura, sin driver).

Prioridad acordada con el usuario para checkboxes:
1) Excel / IDs únicos (por id o name — uno de los dos siempre existe) manda siempre.
2) Sin preferencia explícita: sólo se marca si es requerido (HTML required/aria-required)
   o es un checkbox de términos/privacidad conocido. Los opcionales quedan como están
   (antes se marcaba cualquier checkbox visible sin distinguir).
"""
from osocio.core.base_form_filler import BaseFormFiller

decidir = BaseFormFiller._decidir_marca_checkbox


def test_excel_no_desmarca_sin_importar_lo_demas():
    assert decidir(is_known=True, is_required=True, pref=False, tiene_identificador=True) == "uncheck"


def test_excel_si_marca_sin_importar_lo_demas():
    assert decidir(is_known=False, is_required=False, pref=True, tiene_identificador=True) == "mark"


def test_sin_preferencia_requerido_se_marca():
    assert decidir(is_known=False, is_required=True, pref=None, tiene_identificador=True) == "mark"


def test_sin_preferencia_termino_conocido_se_marca():
    assert decidir(is_known=True, is_required=False, pref=None, tiene_identificador=True) == "mark"


def test_sin_preferencia_opcional_no_se_toca():
    # El caso que cambia respecto del comportamiento viejo: antes esto se marcaba igual.
    assert decidir(is_known=False, is_required=False, pref=None, tiene_identificador=True) == "skip"


def test_sin_identificador_se_saltea_aunque_sea_requerido():
    assert decidir(is_known=False, is_required=True, pref=None, tiene_identificador=False) == "skip"


# --- reconocer el checkbox de terminos por id, no solo por name ---------------------
#
# `terms-and-conditions` ya estaba en la lista de conocidos, pero la comparacion se hacia
# solo contra el atributo `name`. Un form que renderiza el checkbox con `id` y sin `name`
# (React lo hace seguido, porque el id es lo que engancha el <label for>) no se reconocia,
# y entonces solo se marcaba si ademas traia el atributo `required`. Si no, el form no
# dejaba enviar y no quedaba claro por que.

conocido = BaseFormFiller._es_checkbox_conocido


def test_se_reconoce_por_name():
    assert conocido("terms-and-conditions", "")


def test_se_reconoce_por_id():
    """La regresion: mismo valor, pero en el id."""
    assert conocido("", "terms-and-conditions")


def test_alcanza_con_que_uno_de_los_dos_sea_conocido():
    assert conocido("acepta-esto", "terms")
    assert conocido("privacy", "campo-42")


def test_no_importan_mayusculas_ni_espacios():
    assert conocido(" Terms-And-Conditions ", "")
    assert conocido("", "TERMS")


def test_un_checkbox_cualquiera_no_es_conocido():
    """No puede volver el comportamiento viejo de marcar cualquier casilla visible."""
    assert not conocido("test-drive", "newsletter")
    assert not conocido("", "")
    assert not conocido(None, None)


# --- la prioridad tambien sale del id ------------------------------------------------

prioridad = BaseFormFiller._prioridad_checkbox


def test_la_prioridad_sale_del_name():
    assert prioridad("terms", "") == 3


def test_la_prioridad_tambien_sale_del_id():
    """Si no, un checkbox identificado por id quedaba siempre en prioridad 0."""
    assert prioridad("", "terms-and-conditions") == 3


def test_terminos_va_antes_que_contacto():
    """El orden importa: marcar el de contacto primero puede re-renderizar el form."""
    assert prioridad("terms", "") > prioridad("terms-contact", "")


def test_lo_desconocido_va_ultimo():
    assert prioridad("cualquier-cosa", "otra-cosa") == 0
