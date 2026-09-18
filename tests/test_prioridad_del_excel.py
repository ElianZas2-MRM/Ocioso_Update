"""Lo que cargaste en el Excel le gana a cualquier valor fijo guardado.

El caso que destapo esto: el CPF del Excel (`09446875048`) nunca llegaba al formulario de
Cadillac Brasil. En `json/ids_dinamicos.json` habia una entrada:

    {"id": "document", "valor": ["304593", ...], "paises": ["Brasil"], "origen": "autovalor"}

La habia sembrado la app sola, cuando `document` todavia era un campo "no mapeado" (antes
de que existiera el alias `cpf -> document`). El llenado consultaba los IDs dinamicos
ANTES del valor del Excel y hacia `return`, asi que ese 304593 ganaba en cada corrida.

El invariante ya estaba escrito en el codigo, arriba del sembrador de autovalores:
"El Excel siempre gana; nunca pisa lo que el usuario ya cargo a mano". El llenado no lo
respetaba.

Que un valor fijo se aplique cuando la celda esta vacia sigue siendo correcto: para eso
existe la pestana "IDs Dinamicos", para campos que el Excel no cubre (el textarea unico de
los crm-validacion de Argentina, por ejemplo).
"""
import ast
import inspect

import pytest

from osocio.core.base_form_filler import BaseFormFiller

usar_fijo = BaseFormFiller._usar_valor_fijo

GUARDADOS = {"document": ["304593"], "crm-validacion": ["texto de prueba"]}


# --- la regla ------------------------------------------------------------------------

def test_el_excel_le_gana_al_valor_guardado():
    """La regresion: el CPF del Excel tiene que ganarle al autovalor de `document`."""
    assert usar_fijo("document", "09446875048", GUARDADOS) is False


def test_si_la_celda_esta_vacia_se_usa_el_valor_guardado():
    """Para esto existe la pestana de IDs Dinamicos: campos que el Excel no cubre."""
    assert usar_fijo("crm-validacion", "", GUARDADOS) is True
    assert usar_fijo("crm-validacion", None, GUARDADOS) is True


def test_una_celda_con_espacios_cuenta_como_vacia():
    assert usar_fijo("document", "   ", GUARDADOS) is True


def test_sin_entrada_guardada_no_hay_nada_que_usar():
    assert usar_fijo("email", "", GUARDADOS) is False
    assert usar_fijo("email", "alguien@ejemplo.com", GUARDADOS) is False


def test_no_se_rompe_sin_configuracion():
    assert usar_fijo("document", "", None) is False
    assert usar_fijo("document", "", {}) is False


@pytest.mark.parametrize("valor", ["09446875048", 9446875048, "0", 0, "-"])
def test_cualquier_dato_real_del_excel_gana(valor):
    """Incluye el 0 y la marca de omitir: si hay algo en la celda, lo puso una persona."""
    assert usar_fijo("document", valor, GUARDADOS) is False


# --- que el llenado la consulte ------------------------------------------------------

def test_el_llenado_consulta_la_regla():
    """Si el llenado vuelve a mirar los IDs dinamicos por su cuenta, el Excel pierde."""
    fuente = inspect.getsource(BaseFormFiller._fill_visible_fields_from_mapping)
    arbol = ast.parse(fuente.lstrip())
    llamadas = {
        n.func.attr for n in ast.walk(arbol)
        if isinstance(n, ast.Call) and isinstance(n.func, ast.Attribute)
    }
    assert "_usar_valor_fijo" in llamadas, (
        "el llenado no consulta si corresponde usar el valor fijo"
    )
