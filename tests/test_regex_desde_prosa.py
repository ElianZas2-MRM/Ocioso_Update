"""Tests de utils.regex_desde_prosa.

Traduce la prosa del excel del CRM (columnas DESCRIPCION y MENSAJE DE ERROR) a un
regex_full. Contrato: o devuelve un patron que el generador sabe satisfacer, o "" —
nunca un patron inventado ni uno que nadie pueda cumplir.

Los textos de prosa de estos tests estan tomados literales de json/crm_import_report_*.json,
no inventados: si el excel cambia de redaccion, estos tests son los que avisan.
"""
import glob
import json
import os
import re

import pytest

from utils.regex_desde_prosa import (
    ajustar_largo_desde_prosa,
    derivar_regex,
    regex_por_semantica,
)
from utils.valor_campo_generator import GeneradorValorCampo


# (nombre, descripcion, mensajes_error, regex esperado)
CASOS = [
    (
        "cantidad exacta en digitos",
        "17 caracteres alfanumericos",
        ["Ingresa 17 dígitos."],
        r"^[a-zA-Z0-9]{17}$",
    ),
    (
        "rango explicito",
        "De 6 a 7 caracteres alfanuméricos.",
        ["La patente debe tener entre 6 y 7 caracteres."],
        r"^[a-zA-Z0-9]{6,7}$",
    ),
    (
        "minimo acotado por el techo de la descripcion",
        "De 2 a 50 caracteres alfabeticos.",
        ["Ingresa al menos 2 letras."],
        r"^[a-zA-ZáéíóúÁÉÍÓÚñÑ ]{2,50}$",
    ),
    (
        "prefijo forzado de dos digitos",
        "Campo numérico de 9 dígitos. Debe comenzar con 09.",
        ["El celular debe comenzar con 09."],
        r"^(?=(?:09))[0-9]{9}$",
    ),
    (
        "prefijo forzado de un digito",
        "Acepta 10 caracteres numéricos. El primer número del celular debe ser 3.",
        ["Completa los 10 dígitos de tu celular que empieza con 3."],
        r"^(?=(?:3))[0-9]{10}$",
    ),
    (
        "dos largos alternativos",
        "14 caracteres numéricos (Usar el generador de CPF)",
        ["Insira 11/14 dígitos."],
        r"^(?:[0-9]{11}|[0-9]{14})$",
    ),
    (
        "rango corto solo en la descripcion",
        "De 1 a 3 caracteres numéricos.",
        ["Ingresá el orden."],
        r"^[0-9]{1,3}$",
    ),
    (
        # Sin el "de" opcional esto se leia como 120 EXACTOS: todo comentario de Peru
        # habria salido con 120 caracteres clavados.
        "maximo con 'de' intermedio",
        "Form ACDelco. Tiene un maximo de 120 caracteres alfanuméricos.",
        [],
        r"^[a-zA-Z0-9]{1,120}$",
    ),
    (
        "maximo directo",
        "Máximo 30 caracteres alfanuméricos.",
        [],
        r"^[a-zA-Z0-9]{1,30}$",
    ),
]


@pytest.mark.parametrize("nombre,desc,msgs,esperado",
                         CASOS, ids=[c[0] for c in CASOS])
def test_deriva_el_regex_esperado(nombre, desc, msgs, esperado):
    regex, motivo = derivar_regex(desc, msgs)
    assert regex == esperado
    assert motivo, "todo regex derivado tiene que citar la prosa que lo justifico"


# --- el caso que motivo la feature ---

DESC_CONTRATO = "De 1 a 10 caracteres numéricos. No permite iniciar con 0."
MSGS_CONTRATO = ["Ingresá el número de contrato.",
                 "Ingresá 10 caracteres que no comiencen con 0."]


def test_numero_de_contrato_gana_el_mensaje_de_error_sobre_la_descripcion():
    """La descripcion dice "de 1 a 10" y el mensaje dice "10 caracteres". El form rechaza
    todo lo que no tenga 10, asi que manda el mensaje: con la descripcion, un valor de 3
    digitos como "324" pasaba la validacion local y fallaba en produccion."""
    regex, motivo = derivar_regex(DESC_CONTRATO, MSGS_CONTRATO)

    assert regex == r"^[1-9][0-9]{9}$"
    assert "10 caracteres" in motivo
    assert re.fullmatch(regex, "5372819044")
    assert not re.fullmatch(regex, "324"), "el valor que fallaba en el form real"
    assert not re.fullmatch(regex, "0372819044"), "no puede empezar con 0"


def test_no_iniciar_con_cero_no_se_confunde_con_un_prefijo_forzado():
    """"no comiencen con 0" matchea tambien el patron de prefijo forzado; si se leyera al
    reves, el regex exigiria empezar con 0 — exactamente lo contrario."""
    regex, _ = derivar_regex(DESC_CONTRATO, MSGS_CONTRATO)
    assert not regex.startswith("^(?=")


# --- prosa que no alcanza: el modulo no inventa ---

@pytest.mark.parametrize("desc,msgs", [
    ("", ["Ingrese mensaje de error 1."]),        # placeholder del propio excel
    ("", ["Selecciona una opción."]),             # required de dropdown, no describe forma
    ("", ["Deben ser mayores de 18 años."]),      # regla de negocio, no forma
    ("", ["Competa con números entre el 001 y 100."]),  # rango de VALORES, no de largo
    ("Posee un algoritmo de validación.", ["Ingresá tu RUT."]),  # sin clase ni cantidad
    ("De 2 a 50 caracteres.", ["Ingresá tu apellido."]),  # cantidad sin clase
    ("Campo alfanumérico.", ["Ingresa tu VIN."]),  # clase sin cantidad
    ("", []),
    (None, None),
])
def test_prosa_insuficiente_devuelve_vacio(desc, msgs):
    assert derivar_regex(desc, msgs) == ("", "")


# --- contrato con el generador ---

def test_todo_regex_derivado_es_satisfacible_por_el_generador():
    """Barrido sobre la prosa real de los nueve mercados: cada patron derivado tiene que
    compilar y producir un valor que lo cumpla. Es el guardarrail del modulo."""
    raiz = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "json")
    gen = GeneradorValorCampo(semilla=99)

    derivados = 0
    for ruta in glob.glob(os.path.join(raiz, "crm_import_report_*.json")):
        with open(ruta, encoding="utf-8") as fh:
            reporte = json.load(fh)
        for seccion in ("faltantes", "incompletos", "conflictos", "ok"):
            for campo in reporte.get(seccion, []):
                regex, _ = derivar_regex(campo.get("descripcion"), campo.get("mensajes_error"))
                if not regex:
                    continue
                derivados += 1
                valores = gen.generar_desde_regex(regex, max_variantes=2)
                assert valores, f"{regex!r} no produjo ningun valor ({campo['nombre_campo']})"
                for v in valores:
                    assert re.fullmatch(regex, v), f"{v!r} no matchea {regex!r}"

    assert derivados > 50, f"cobertura sospechosamente baja: solo {derivados} patrones"


def test_nunca_revienta_con_prosa_arbitraria():
    for basura in ["((((", "[0-9", "\\", "…" * 50, "1" * 500, "de a y entre 3"]:
        assert isinstance(derivar_regex(basura, [basura]), tuple)


# --- ajustar_largo_desde_prosa: corrige el largo sin tocar el resto de la regla ---

def test_ajusta_el_largo_conservando_el_lookahead():
    """El caso testigo. `(?=(?:1|...|9))` impide empezar con 0 y la prosa no lo menciona:
    si se regenerara el regex entero desde la prosa, esa regla se perderia."""
    nuevo, motivo = ajustar_largo_desde_prosa(
        r"^(?=(?:1|2|3|4|5|6|7|8|9))[0-9]{1,10}$", DESC_CONTRATO, MSGS_CONTRATO)

    assert nuevo == r"^(?=(?:1|2|3|4|5|6|7|8|9))[0-9]{10}$"
    assert "10 caracteres" in motivo
    assert not re.fullmatch(nuevo, "324")
    assert not re.fullmatch(nuevo, "0372819044")
    assert re.fullmatch(nuevo, "5372819044")


def test_cierra_un_cuantificador_abierto():
    nuevo, _ = ajustar_largo_desde_prosa(
        r"^[a-zA-Z0-9]{17,}$", "17 caracteres alfanumericos", ["Ingresá 17 dígitos."])
    assert nuevo == r"^[a-zA-Z0-9]{17}$"


def test_la_prosa_gana_cuando_el_largo_cargado_es_imposible():
    """El JSON pide 15 fijos y el form acepta hasta 10: no hay intersección, así que manda
    el mensaje de error, que es lo que el form muestra al rechazar."""
    nuevo, _ = ajustar_largo_desde_prosa(
        r"^[a-zA-Z0-9]{15}$", "", ["Permite hasta 10 caracteres alfanuméricos."])
    assert nuevo == r"^[a-zA-Z0-9]{1,10}$"


def test_un_minimo_de_la_prosa_no_borra_el_techo_ya_cargado():
    """"mínimo 10" no dice nada del techo: conservar el 15 es más informativo que abrirlo."""
    assert ajustar_largo_desde_prosa(
        r"^(?!(.)(\1+$))[0-9]{10,15}$", "", ["Ingresá mínimo 10 números."]) == ("", "")


def test_descuenta_los_caracteres_fijos_del_patron():
    """En la cédula ecuatoriana el prefijo de provincia aporta 2 caracteres al largo total,
    así que el cuantificador no puede tomar el número de la prosa tal cual."""
    actual = r"^((0[0-9]|1[0-9]|2[0-4])|30)[0-9]{5,8}$"
    # Largo total 7..10; la prosa pide exactamente 9 -> el cuantificador baja a 7.
    nuevo, _ = ajustar_largo_desde_prosa(actual, "", ["Ingresa 9 dígitos."])
    assert nuevo == r"^((0[0-9]|1[0-9]|2[0-4])|30)[0-9]{7}$"
    for v in GeneradorValorCampo(semilla=1).generar_desde_regex(nuevo, max_variantes=2):
        assert len(v) == 9


@pytest.mark.parametrize("actual,desc,msgs,motivo", [
    # Ya coincide: no hay nada que corregir.
    (r"^[0-9]{8}$", "8 caracteres numéricos", ["Ingresa 8 dígitos."], "largo ya correcto"),
    # Sin cuantificador no hay qué ajustar (alternación de valores literales).
    (r"^(00[1-9]|0[1-9][0-9]|100)$", "", ["Ingresa 3 dígitos."], "sin cuantificador"),
    # Dos cuantificadores: no se sabe a cuál se refiere la prosa.
    (r"^[A-Za-z]{3}[A-Za-z0-9]{0,4}$", "", ["Ingresa 7 caracteres."], "ambiguo"),
    # La prosa no habla de largo.
    (r"^[0-9]{8}$", "Posee un algoritmo de validación.", ["Ingresá tu RUT."], "sin largo"),
    # Dos largos alternativos no se expresan con un solo cuantificador.
    (r"^[0-9]{11}$", "", ["Insira 11/14 dígitos."], "alternativa"),
])
def test_no_toca_la_regla_cuando_no_corresponde(actual, desc, msgs, motivo):
    assert ajustar_largo_desde_prosa(actual, desc, msgs) == ("", ""), motivo


def test_regex_vacio_o_invalido_no_revienta():
    assert ajustar_largo_desde_prosa("", "x", ["Ingresa 8 dígitos."]) == ("", "")
    assert ajustar_largo_desde_prosa("^[0-9{8}$", "x", ["Ingresa 8 dígitos."]) == ("", "")
    assert ajustar_largo_desde_prosa(None, None, None) == ("", "")


def test_el_ajuste_nunca_deja_un_regex_insatisfacible():
    """Barrido sobre las reglas vivas del repo: si se ajusta, el resultado tiene que seguir
    produciendo valores válidos."""
    raiz = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "json")
    gen = GeneradorValorCampo(semilla=11)

    for ruta in glob.glob(os.path.join(raiz, "crm_import_report_*.json")):
        with open(ruta, encoding="utf-8") as fh:
            reporte = json.load(fh)
        for campo in reporte.get("ok", []) + reporte.get("conflictos", []):
            actual, _ = derivar_regex(campo.get("descripcion"), campo.get("mensajes_error"))
            if not actual:
                continue
            nuevo, _ = ajustar_largo_desde_prosa(
                actual, campo.get("descripcion"), campo.get("mensajes_error"))
            if not nuevo:
                continue
            valores = gen.generar_desde_regex(nuevo, max_variantes=2)
            assert valores, f"{nuevo!r} quedó sin solución tras el ajuste"
            for v in valores:
                assert re.fullmatch(nuevo, v)


# --- fallback semantico ---

@pytest.mark.parametrize("field_id,esperado_digitos", [
    ("contract", True),
    ("telephone", True),
    ("order", True),
    ("ci", True),
])
def test_semantica_numerica(field_id, esperado_digitos):
    regex = regex_por_semantica(field_id)
    assert regex
    for v in GeneradorValorCampo(semilla=5).generar_desde_regex(regex, max_variantes=2):
        assert v.isdigit() is esperado_digitos
        assert not v.startswith("0")


@pytest.mark.parametrize("field_id,largo", [
    # El VIN son 17 por norma ISO 3779; la patente comparte vocabulario pero es más corta.
    ("vin", 17), ("cc_vin", 17), ("chasis", 17),
    ("patent", None), ("placa", None),
])
def test_semantica_distingue_vin_de_patente(field_id, largo):
    regex = regex_por_semantica(field_id)
    assert regex
    for v in GeneradorValorCampo(semilla=5).generar_desde_regex(regex, max_variantes=2):
        if largo:
            assert len(v) == largo, f"{field_id} debería dar {largo} caracteres, dio {len(v)}"
        else:
            assert 6 <= len(v) <= 8


def test_semantica_email_usa_el_dominio_de_prueba():
    regex = regex_por_semantica("email", "Correo electrónico")
    valores = GeneradorValorCampo(semilla=5).generar_desde_regex(regex, max_variantes=1)
    assert valores and valores[0].endswith("@mrm.com")


@pytest.mark.parametrize("field_id", ["ciudad", "city", "nacimiento", "vinculo", "direccion"])
def test_claves_cortas_no_matchean_como_substring(field_id):
    """"ci" esta dentro de "ciudad"/"nacimiento" y "vin" dentro de "vinculo": tomarlas como
    substring convertiria un campo de texto en uno numerico."""
    assert regex_por_semantica(field_id) == ""


@pytest.mark.parametrize("field_id,label", [
    ("field-42", "  "),
    ("", ""),
    ("regret-reason", "Motivo"),
])
def test_sin_semantica_reconocible_devuelve_vacio(field_id, label):
    assert regex_por_semantica(field_id, label) == ""


def test_id_en_kebab_o_snake_se_trocea_para_el_match():
    assert regex_por_semantica("first_name") == regex_por_semantica("firstname")
    assert regex_por_semantica("numero-de-contrato")


@pytest.mark.parametrize("field_id", [
    "telephone_prefix", "Celular_prefix", "codigo_de_area", "prefijo-celular",
])
def test_no_infiere_nada_para_un_campo_de_prefijo(field_id):
    """Un "telephone_prefix" matchea la raiz "telephone" pero no es un telefono: es un
    prefijo de 1-3 digitos. Inferirlo como numero largo llenaba el campo con 8 digitos y
    el form rechazaba el lead — peor que dejarlo vacio, que es lo que pasaba antes."""
    assert regex_por_semantica(field_id) == ""
