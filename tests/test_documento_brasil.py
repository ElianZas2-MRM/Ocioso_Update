"""Los documentos de Brasil salen del Excel, no se inventan en medio de la corrida.

El caso que destapo esto: el formulario de Cadillac Brasil T1 llama `document` al campo de
CPF y `zip_code` al de CEP. El mapping de Brasil busca `cpf` y `cep`, y el match es exacto
por id, asi que ninguno de los dos matcheaba. El campo quedaba "no mapeado" y lo rellenaba
el auto-descubrimiento con un valor inventado: en la corrida real viajo `304593` como CPF
—seis digitos, ni siquiera un CPF valido— mientras en el Excel habia `09446875048`.

Dos cosas se arreglan juntas porque una rompe a la otra si van separadas:

1. Los alias `cpf -> document` y `cep -> zip_code`, para que el campo se encuentre.
2. La normalizacion del documento, que preguntaba `"cpf" in field_id`. Al resolverse el
   alias, field_id pasa a ser `document`, donde no esta la palabra "cpf": arreglar solo el
   alias habria dejado de rellenar el cero inicial que Excel se come.

Y la regla de fondo: **el CPF se genera al crear el Excel, nunca durante una corrida.** Si
se genera en el momento, el lead viaja con un numero que no esta en ninguna parte del
Excel y despues no hay contra que comparar el resultado.
"""
import inspect

import pytest

from osocio.core.base_form_filler import BaseFormFiller
from osocio.utils.field_id_aliases import alias_ids_for

tipo = BaseFormFiller._tipo_documento_brasil
normalizar = BaseFormFiller._normalizar_documento_brasil


# --- encontrar el campo -------------------------------------------------------------

def test_el_cpf_se_encuentra_como_document():
    """El id que usa el estandar gm_frontend."""
    assert "document" in alias_ids_for("cpf")


def test_el_cep_se_encuentra_como_zip_code():
    assert "zip_code" in alias_ids_for("cep")


def test_el_alias_funciona_tambien_al_reves():
    """Una regla escrita con el id nuevo tiene que correr contra un form viejo."""
    assert "cpf" in alias_ids_for("document")
    assert "cep" in alias_ids_for("zip_code")


def test_document_sigue_sirviendo_para_el_ci_de_peru():
    """Ya existia `ci -> document`. Agregar el CPF no lo puede pisar."""
    assert "document" in alias_ids_for("ci")
    assert "ci" in alias_ids_for("document")


# --- reconocer de que documento se trata --------------------------------------------

@pytest.mark.parametrize("ids, esperado", [
    (("cpf",), "cpf"),
    (("cnpj",), "cnpj"),
    (("cep",), "cep"),
    (("customer-cep",), "cep"),
    (("zip",), "cep"),
    (("postal",), "cep"),
    (("zip_code",), "cep"),          # el id del form de Cadillac
])
def test_se_reconoce_el_tipo_por_el_id(ids, esperado):
    assert tipo(*ids) == esperado


def test_se_reconoce_por_el_id_del_mapping_aunque_el_del_dom_no_lo_diga():
    """La regresion: resuelto el alias, el id del DOM es `document`, que no dice "cpf".

    Si solo se mirara ese, el CPF dejaria de normalizarse justo en el form que motivo
    todo esto.
    """
    assert tipo("document", "cpf") == "cpf"


@pytest.mark.parametrize("ids", [
    ("email",),
    ("firstname",),
    ("document",),      # sin el id del mapping al lado no se puede saber cual es
    (),
    (None,),
])
def test_lo_que_no_es_documento_no_se_toca(ids):
    assert tipo(*ids) is None


def test_cnpj_no_se_confunde_con_cpf():
    assert tipo("cnpj") == "cnpj"
    assert tipo("cpf") == "cpf"


# --- normalizar lo que vino del Excel -----------------------------------------------

def test_un_cpf_bien_cargado_pasa_intacto():
    assert normalizar("cpf", "09446875048") == "09446875048"


def test_le_saca_el_formato():
    """Alguien puede pegar el CPF con puntos y guion; el form quiere solo digitos."""
    assert normalizar("cpf", "094.468.750-48") == "09446875048"
    assert normalizar("cep", "81230-031") == "81230031"


@pytest.mark.parametrize("tipo_doc, valor, esperado", [
    ("cpf", 9446875048, "09446875048"),       # 10 digitos -> 11
    ("cnpj", 5679080000177, "05679080000177"),  # 13 -> 14
    ("cep", 4538133, "04538133"),             # 7 -> 8
])
def test_devuelve_el_cero_que_excel_se_come(tipo_doc, valor, esperado):
    """Una celda numerica pierde el cero inicial. Es la causa mas comun de un CPF invalido."""
    assert normalizar(tipo_doc, valor) == esperado


def test_no_agrega_ceros_si_no_falta_exactamente_uno():
    """Rellenar un numero al que le faltan tres digitos seria inventar datos."""
    assert normalizar("cpf", "12345") == "12345"


def test_una_celda_vacia_queda_vacia():
    """Lo importante: NO se genera un documento. Antes, aca se inventaba uno."""
    assert normalizar("cpf", "") == ""
    assert normalizar("cpf", None) == ""
    assert normalizar("cpf", "   ") == ""


def test_un_valor_sin_digitos_se_respeta():
    """Si alguien escribio algo raro, que llegue al form y que el form lo rechace.

    Es un resultado de test valido; taparlo con un documento generado no lo es.
    """
    assert normalizar("cpf", "-") == "-"


# --- la regla de fondo ---------------------------------------------------------------

def test_el_llenado_no_genera_documentos():
    """El unico generador legitimo es el del Excel (pestana "Generar Excels con Datos").

    Si el llenado genera uno, el lead viaja con un numero que no esta en el Excel y el
    resultado no se puede comparar contra lo que se pidio.
    """
    fuente = inspect.getsource(BaseFormFiller._fill_visible_fields_from_mapping)
    assert "_generate_brazil_document" not in fuente, (
        "el llenado volvio a generar documentos en vez de usar los del Excel"
    )
