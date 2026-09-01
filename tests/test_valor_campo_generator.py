"""Tests de utils.valor_campo_generator.GeneradorValorCampo.

El generador arma valores de prueba para campos detectados en vivo que no tienen
mapeo ni valor en el Excel. Contrato central: TODO string que devuelve pasa
re.fullmatch(regex_full, x); si no puede satisfacer el patron devuelve [] (el
caller cae al comportamiento previo). Nunca levanta excepcion, nunca devuelve
un valor invalido.
"""
import glob
import json
import os
import re

import pytest

from utils.valor_campo_generator import GeneradorValorCampo


# Patrones representativos tomados 1:1 de json/field_validation_rules_*.json.
# Se listan aca (ademas del barrido dinamico de mas abajo) para que un fallo
# apunte al shape exacto que se rompio.
PATRONES_LIMPIOS = [
    # nombre / apellido: sin triple repetido, con vocal y consonante, letras y espacio
    r"^(?!.*(.)(\1){2})(?=.*[aeiouAEIOUáéíóúÁÉÍÓÚ])(?=.*[bcdfghjklmnpqrstvwxyzBCDFGHJKLMNPQRSTVWXYZñÑ])[a-zA-ZáéíóúÁÉÍÓÚñÑ ]{1,50}$",
    r"^(?!.*(.)(\1){2})(?=.*[aeiouAEIOU])(?=.*[bcdfghjklmnpqrstvwxyz])[a-zA-Z ]{2,50}$",
    # documentos / telefonos numericos de largo fijo o rango
    r"^[0-9]{11}$",
    r"^[0-9]{9}$",
    r"^[0-9]{8}$",
    r"^[0-9]{1,3}$",
    r"^[0-9]{5,12}$",
    r"^[0-9]{7,10}$",
    r"^[0-9]{7,9}$",
    r"^[0-9]{17}$",
    r"^[0-9]{12}$",
    # alfanumericos
    r"^[a-zA-Z0-9]{17}$",
    r"^[a-zA-Z0-9]{15}$",
    r"^[a-zA-Z0-9]{6}$",
    r"^[a-zA-Z0-9]{6,7}$",
    r"^[a-zA-Z0-9]{2,36}$",
    r"^[a-zA-Z0-9]{1,40}$",
    r"^[a-zA-Z0-9]{17,}$",
    # "no todos iguales" + digitos
    r"^(?!(.)(\1+$))[0-9]{10,15}$",
    r"^(?!(.)(\1+$))[0-9]{7,9}$",
    r"^(?!(.)(\1+$))[0-9]{10}$",
    r"^(?!(.)(\1+$))[0-9]{12}$",
    r"^(?!(.)(\1+$))[0-9]{8}$",
    r"^(?!(.)(\1+$))[a-zA-Z0-9]{2,50}$",
    # prefijo forzado por lookahead
    r"^(?=(?:6|7))[0-9]{8}$",
    r"^(?=(?:3))[0-9]{10}$",
    r"^(?=(?:09))[0-9]{10}$",
    r"^(?=(?:9))[A-Z0-9]{9}$",
    r"^(?=(?:1|2|3|4|5|6|7|8|9))[0-9]{1,10}$",
    r"^(?!(.)(\1+$))(?=(?:1|2|3|4|5|6|7|8|9))[0-9]{7}$",
    r"^(?!(.)(\1+$))(?=(?:09))[0-9]{9}$",
    # RUT chileno: cuerpo - digito verificador
    r"^[0-9]{7,8}-[0-9kK]{1}$",
    # alternaciones finitas
    r"^(00[1-9]|0[1-9][0-9]|100)$",
    r"^((0[0-9]|1[0-9]|2[0-4])|30)[0-9]{5,8}$",
    # patente tipo ABC / ABC1234
    r"^[A-Za-z]{3}[A-Za-z0-9]{0,4}$",
    # texto largo (mensaje / comentario)
    r"^[a-zA-ZñÑ .,;:!?\"'()_\-+/@#%&*0-9]{5,500}$",
    # email
    r"^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)+$",
]


@pytest.fixture
def gen():
    # Semilla fija: los tests que comparan cantidad/distincion de variantes no
    # deben depender del azar del proceso.
    return GeneradorValorCampo(semilla=1234)


@pytest.mark.parametrize("patron", PATRONES_LIMPIOS)
def test_genera_valor_valido_para_patrones_limpios(gen, patron):
    valores = gen.generar_desde_regex(patron, max_variantes=2)
    assert valores, f"no genero ningun valor para {patron!r}"
    for v in valores:
        assert re.fullmatch(patron, v), f"{v!r} no matchea {patron!r}"


def test_dos_variantes_son_distintas_cuando_el_patron_lo_permite(gen):
    valores = gen.generar_desde_regex(r"^[0-9]{8}$", max_variantes=2)
    assert len(valores) == 2
    assert valores[0] != valores[1]


def test_respeta_max_variantes(gen):
    assert len(gen.generar_desde_regex(r"^[0-9]{8}$", max_variantes=1)) == 1
    assert len(gen.generar_desde_regex(r"^[0-9]{8}$", max_variantes=3)) <= 3


def test_patron_de_una_sola_solucion_devuelve_una_variante(gen):
    # Solo "5" cumple: 1 digito, entre 1 y 9 (lookahead), longitud exacta 1.
    valores = gen.generar_desde_regex(r"^(?=(?:5))[0-9]{1}$", max_variantes=2)
    assert valores == ["5"]


def test_email_usa_dominio_mrm(gen):
    patron = r"^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)+$"
    valores = gen.generar_desde_regex(patron, max_variantes=2)
    assert valores
    for v in valores:
        assert re.fullmatch(patron, v)
        assert v.endswith("@mrm.com")


def test_campo_texto_largo_con_arroba_en_clase_no_se_trata_como_email(gen):
    # La clase admite '@' pero el campo es un comentario, no un email: el valor
    # generado debe matchear el patron y NO ser una direccion de correo.
    patron = r"^[a-zA-ZñÑ .,;:!?\"'()_\-+/@#%&*0-9]{5,500}$"
    valores = gen.generar_desde_regex(patron, max_variantes=2)
    assert valores
    for v in valores:
        assert re.fullmatch(patron, v)
        assert "@mrm.com" not in v


@pytest.mark.parametrize("vacio", ["", "   ", None])
def test_regex_vacio_devuelve_lista_vacia(gen, vacio):
    assert gen.generar_desde_regex(vacio) == []


def test_patron_imposible_devuelve_lista_vacia_sin_reventar(gen):
    # Contradiccion: el lookahead exige empezar con 9 pero la clase solo admite 0-1.
    assert gen.generar_desde_regex(r"^(?=(?:9))[01]{5}$") == []


def test_patron_no_soportado_nunca_devuelve_valor_invalido(gen):
    # Backreference de grupo capturado en el cuerpo: fuera de alcance del generador.
    # Debe devolver [] o valores validos, jamas algo que no matchee.
    for v in gen.generar_desde_regex(r"^([a-z]{3})\1$"):
        assert re.fullmatch(r"^([a-z]{3})\1$", v)


def test_todos_los_patrones_reales_del_repo_son_seguros(gen):
    """Barrido sobre los regex_full vivos: o genera un valor valido, o [].
    Nunca un valor invalido, nunca una excepcion."""
    raiz = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "json")
    patrones = set()
    for ruta in glob.glob(os.path.join(raiz, "field_validation_rules_*.json")):
        with open(ruta, encoding="utf-8") as fh:
            data = json.load(fh)
        for regla in (data.get("fields") or {}).values():
            rf = (regla.get("regex_full") or "").strip()
            if rf:
                patrones.add(rf)

    assert patrones, "no se encontraron patrones reales para el barrido"
    sin_solucion = []
    for rf in patrones:
        try:
            valores = gen.generar_desde_regex(rf, max_variantes=2)
        except Exception as exc:  # noqa: BLE001 - el contrato es que NUNCA reviente
            pytest.fail(f"generar_desde_regex revento con {rf!r}: {exc}")
        for v in valores:
            assert re.fullmatch(rf, v), f"{v!r} no matchea patron real {rf!r}"
        if not valores:
            sin_solucion.append(rf)

    # Cota blanda: la mayoria de los patrones reales deberian resolverse. Si esto
    # se dispara es que el generador perdio cobertura, no un fallo duro.
    assert len(sin_solucion) <= len(patrones) // 3, (
        f"demasiados patrones sin solucion ({len(sin_solucion)}/{len(patrones)}): {sin_solucion}"
    )


# --- elegir_opciones_dropdown ---

def test_elige_k_opciones_distintas(gen):
    opciones = ["Rojo", "Verde", "Azul", "Negro", "Blanco"]
    elegidas = gen.elegir_opciones_dropdown(opciones, k=2)
    assert len(elegidas) == 2
    assert len(set(elegidas)) == 2
    assert set(elegidas).issubset(set(opciones))


def test_dropdown_con_menos_opciones_que_k(gen):
    assert sorted(gen.elegir_opciones_dropdown(["Unica"], k=2)) == ["Unica"]


def test_dropdown_descarta_vacios_y_espacios(gen):
    elegidas = gen.elegir_opciones_dropdown(["  ", "", "Valida", None], k=2)
    assert elegidas == ["Valida"]


def test_dropdown_lista_vacia(gen):
    assert gen.elegir_opciones_dropdown([], k=2) == []
