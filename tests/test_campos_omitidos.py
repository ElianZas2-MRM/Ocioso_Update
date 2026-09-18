"""El Excel puede pedir que un campo NO se complete.

Dos casos reales, los dos del formulario de Cadillac Brasil T1:

- CPF y CEP son opcionales y se querian poder dejar vacios a proposito. (Cuando se
  escribio esto la app ademas los generaba sola si la celda venia vacia; eso se saco
  despues: los documentos se generan al crear el Excel, nunca durante una corrida.)
- El dropdown de concesionario viene bloqueado (`disabled`) a proposito. Igual quedaba
  anotado como "quedo sin elegir" y marcaba como sucia una corrida que estuvo bien.

Lo que estos tests fijan sobre todo es el DEFAULT: una celda vacia se sigue comportando
como siempre. Los Excels que ya existen estan llenos de celdas vacias y ninguna queria
decir "omitir"; si lo vacio cambiara de significado, cada corrida vieja cambiaria de
resultado en silencio.
"""
import ast
import inspect

import pytest

from osocio.utils.omitir_campo import MARCAS, pide_omitir


# --- la regla, sola --------------------------------------------------------------------

@pytest.mark.parametrize("marca", sorted(MARCAS))
def test_las_marcas_piden_omitir(marca):
    assert pide_omitir(marca)


@pytest.mark.parametrize("escrito", [
    " - ",          # un espacio de mas al tipear
    "OMITIR",
    "Omitir",
    "VACIO",
    "N/A",
    "  skip",
])
def test_no_importan_mayusculas_ni_espacios(escrito):
    """Lo escribe una persona a mano en una celda; tiene que perdonar el tipeo."""
    assert pide_omitir(escrito)


@pytest.mark.parametrize("celda", [
    None,
    "",
    "   ",
])
def test_una_celda_vacia_no_pide_nada(celda):
    """EL default, y la mitad del pedido: "si no pongo nada en excel entonces por
    defecto que llene el de siempre".

    Vacio no es una instruccion: el campo se completa con lo que corresponda. Omitir hay
    que pedirlo.
    """
    assert not pide_omitir(celda)


@pytest.mark.parametrize("valor", [
    "41863799025",          # un CPF
    41863799025,            # el mismo, como lo devuelve openpyxl de una celda numerica
    "04538-133",            # un CEP con guion adentro
    4538133,
    "Cadillac Sao Paulo",   # un concesionario
    0,
    -1,
    "0",
])
def test_un_valor_real_no_se_confunde_con_una_marca(valor):
    """Un CEP tiene un guion adentro y no por eso pide omitirse."""
    assert not pide_omitir(valor)


def test_no_se_acepta_no_como_marca():
    """En las columnas de checkbox "NO" ya significa destildar la casilla.

    Dos significados para la misma palabra en el mismo Excel es pedirle a alguien que se
    equivoque.
    """
    assert not pide_omitir("NO")
    assert not pide_omitir("no")


# --- el motor la usa -------------------------------------------------------------------

def _filler():
    """Un BaseFormFiller sin navegador, solo para el estado por fila."""
    from osocio.core.base_form_filler import BaseFormFiller

    f = object.__new__(BaseFormFiller)
    f.current_row_field_values = {}
    f._dropdowns_sin_elegir = []
    f._campos_omitidos = []
    return f


def test_se_anota_el_campo_omitido():
    f = _filler()
    assert f._pide_omitir_campo("cpf", "-") is True
    assert "cpf" in f._campos_omitidos


def test_no_se_anota_lo_que_no_se_omite():
    f = _filler()
    assert f._pide_omitir_campo("cpf", "41863799025") is False
    assert f._campos_omitidos == []


def test_no_se_anota_dos_veces():
    """Un form multi-paso pasa por el mismo campo en varias pasadas."""
    f = _filler()
    f._pide_omitir_campo("cep", "omitir")
    f._pide_omitir_campo("cep", "omitir")
    assert f._campos_omitidos == ["cep"]


def test_arranca_limpio_en_cada_fila():
    """Cada fila del Excel es un formulario distinto.

    Si la lista sobreviviera de una fila a la siguiente, omitir el concesionario en una URL
    lo omitiria en todas las que vengan despues. Ya paso algo asi con los dropdowns sin
    elegir, por eso hay un reset por fila.
    """
    from osocio.core.base_form_filler import BaseFormFiller

    f = _filler()
    f._campos_omitidos = ["dealer", "cpf"]
    f._ids_propios_cache = object()
    BaseFormFiller.begin_row_tracking(f)
    assert f._campos_omitidos == []


# --- el caso del concesionario bloqueado ----------------------------------------------

class DriverConUnSelect:
    """Devuelve un solo <select> visible, parado en su placeholder."""

    def __init__(self, fid="dealer", texto="Selecione"):
        self.fid = fid
        self.texto = texto

    def execute_script(self, script, *args):
        if "querySelectorAll('select')" in script:
            return [{"id": self.fid, "texto": self.texto}]
        return []


def test_un_dropdown_omitido_no_se_reporta_como_sin_elegir():
    """El caso del concesionario: bloqueado a proposito, no es una falla de la corrida."""
    f = _filler()
    f.driver = DriverConUnSelect("dealer")
    f._campos_omitidos = ["dealer"]

    anotados = f._registrar_dropdowns_en_placeholder()

    assert anotados == []
    assert f._dropdowns_sin_elegir == [], (
        "un campo que el Excel pidio omitir no puede contar como dato que falto"
    )


def test_un_dropdown_no_omitido_sigue_reportandose():
    """Lo de arriba no puede tapar un dropdown que de verdad quedo sin elegir."""
    f = _filler()
    f.driver = DriverConUnSelect("models")
    f._campos_omitidos = ["dealer"]

    anotados = f._registrar_dropdowns_en_placeholder()

    assert anotados == ["models"]
    assert f._dropdowns_sin_elegir == ["models"]


# --- que la regla se consulte en los dos caminos de llenado ---------------------------

@pytest.mark.parametrize("metodo", [
    "_fill_visible_fields_from_mapping",   # el llenado normal
    "fill_fields_present",                 # el de los forms multi-paso
])
def test_los_dos_caminos_de_llenado_consultan_la_regla(metodo):
    """Hay dos funciones que llenan campos. Si una sola consulta, la otra ignora el Excel.

    Es el mismo problema que tenia el pre-scroll copiado cuatro veces: arreglado en un
    lugar y roto en el resto.
    """
    from osocio.core.base_form_filler import BaseFormFiller

    fuente = inspect.getsource(getattr(BaseFormFiller, metodo))
    arbol = ast.parse(fuente.lstrip())
    llamadas = {
        n.func.attr for n in ast.walk(arbol)
        if isinstance(n, ast.Call) and isinstance(n.func, ast.Attribute)
    }
    assert "_pide_omitir_campo" in llamadas, (
        f"{metodo} no consulta si el Excel pidio omitir el campo"
    )


def test_la_marca_se_consulta_antes_de_tocar_el_documento():
    """El bloque que normaliza CPF/CEP no puede correr para un campo omitido.

    Si se consultara despues, un "-" en la celda de CPF pasaria primero por la
    normalizacion, que le saca todo lo que no sea digito y lo dejaria en "": el campo se
    escribiria vacio en vez de no tocarse, y la diferencia se perderia.

    (Antes este test miraba el generador de documentos, que se saco del llenado: los CPF
    se generan al crear el Excel, nunca durante una corrida.)
    """
    from osocio.core.base_form_filler import BaseFormFiller

    fuente = inspect.getsource(BaseFormFiller._fill_visible_fields_from_mapping)
    assert "_pide_omitir_campo" in fuente
    assert fuente.index("_pide_omitir_campo") < fuente.index("_normalizar_documento_brasil"), (
        "la consulta tiene que ir antes de normalizar el documento"
    )
