"""Las opciones del envío de leads se leen bien de la pantalla.

Hasta hace poco esto no se podía probar: la corrida leía las once opciones directamente de
los widgets de Tkinter, desde adentro de una función anidada de 1.054 líneas. Para
verificar que `url_max` se acotara entre 1 y 20 había que abrir la app.

Ahora hay un objeto de opciones y una función que lo arma, así que se puede probar con
dobles de los widgets y sin levantar una ventana.
"""
import dataclasses

import pytest

from osocio.interface.envio_leads import OpcionesEnvioLeads, _leer_opciones


class VarFalsa:
    """Un doble de `tkinter.BooleanVar` / `StringVar`: solo necesita `.get()`."""

    def __init__(self, valor):
        self._valor = valor

    def get(self):
        return self._valor


class ContextoFalso:
    """Lo mínimo que `_leer_opciones` necesita del contexto."""

    def __init__(self, **cambios):
        por_defecto = {
            "var_t3": False,
            "var_t3_also": False,
            "var_sched_t3": False,
            "var_url_parallel": False,
            "url_max_var": 3,
            "var_enviar_email": True,
            "var_modo_email": "por_pais",
            "email_entry": "alguien@ejemplo.com",
            "var_adjuntar_res": True,
            "var_adjuntar_ss": False,
            "var_ver_navegador": False,
        }
        por_defecto.update(cambios)
        for nombre, valor in por_defecto.items():
            setattr(self, nombre, VarFalsa(valor))


# --- la forma del objeto -----------------------------------------------------------

def test_es_inmutable():
    """Si alguien necesita mutarla, lo que necesita es una lectura en vivo, no una opción.

    Tres valores quedaron deliberadamente fuera de este objeto porque se leen mientras la
    corrida avanza. `frozen=True` hace que esa distinción se note al intentar romperla.
    """
    opciones = _leer_opciones(ContextoFalso(), scheduled=False)
    with pytest.raises(dataclasses.FrozenInstanceError):
        opciones.t3 = True


def test_no_entraron_las_que_se_leen_en_vivo():
    """`pausar_autenticacion`, `preview_navegador` y `selected_disp` NO van acá.

    Se leen adentro de funciones que corren después, en hilos: si el usuario destilda
    "pausar autenticación" con la corrida ya empezada, las sesiones siguientes tienen que
    respetar el cambio. Congelarlas rompería eso en silencio.
    """
    campos = {f.name for f in dataclasses.fields(OpcionesEnvioLeads)}
    for prohibido in ("pausar_autenticacion", "preview_navegador", "selected_disp"):
        assert prohibido not in campos, (
            f"'{prohibido}' se lee en vivo durante la corrida: no puede ser una opción congelada"
        )


# --- lo que lee de la pantalla ------------------------------------------------------

def test_toma_lo_que_el_usuario_dejo_elegido():
    opciones = _leer_opciones(
        ContextoFalso(var_enviar_email=True, var_modo_email="consolidado",
                      email_entry="qa@osocio.com", var_adjuntar_res=True,
                      var_adjuntar_ss=True),
        scheduled=False,
    )
    assert opciones.enviar_mail is True
    assert opciones.modo_email == "consolidado"
    assert opciones.destinatario == "qa@osocio.com"
    assert opciones.adjuntar_resultados is True
    assert opciones.adjuntar_capturas is True


def test_le_saca_los_espacios_al_destinatario():
    """Un espacio pegado al mail lo hace rebotar, y es facilísimo de pegar sin querer."""
    opciones = _leer_opciones(ContextoFalso(email_entry="  qa@osocio.com  "), scheduled=False)
    assert opciones.destinatario == "qa@osocio.com"


@pytest.mark.parametrize("ingresado, esperado", [
    (0, 1),      # nadie puede correr 0 URLs
    (1, 1),
    (7, 7),
    (20, 20),
    (99, 20),    # el tope existe para no reventar la maquina
    (-5, 1),
])
def test_la_cantidad_de_urls_en_paralelo_queda_entre_1_y_20(ingresado, esperado):
    opciones = _leer_opciones(ContextoFalso(url_max_var=ingresado), scheduled=False)
    assert opciones.url_max == esperado


def test_ver_navegador_se_lee_tal_cual():
    assert _leer_opciones(ContextoFalso(var_ver_navegador=True), scheduled=False).ver_navegador
    assert not _leer_opciones(ContextoFalso(var_ver_navegador=False), scheduled=False).ver_navegador


# --- la regla de los formularios 2.0 (T3) -------------------------------------------

def test_si_la_corrida_es_solo_t3_no_se_agregan_los_normales():
    """`t3_also` significa 'además de los normales'. Si ya es una corrida T3, no aplica."""
    opciones = _leer_opciones(ContextoFalso(var_t3=True, var_t3_also=True), scheduled=False)
    assert opciones.t3 is True
    assert opciones.t3_also is False


def test_en_una_corrida_manual_manda_la_casilla_de_la_pantalla():
    opciones = _leer_opciones(
        ContextoFalso(var_t3=False, var_t3_also=True, var_sched_t3=False), scheduled=False
    )
    assert opciones.t3_also is True


def test_en_una_corrida_programada_manda_la_casilla_del_scheduler():
    """Acá está el detalle: programada y manual leen variables DISTINTAS.

    Si se mezclaran, una corrida automática usaría lo que quedó tildado en pantalla en vez
    de lo que se configuró en la programación.
    """
    opciones = _leer_opciones(
        ContextoFalso(var_t3=False, var_t3_also=False, var_sched_t3=True), scheduled=True
    )
    assert opciones.t3_also is True

    opciones = _leer_opciones(
        ContextoFalso(var_t3=False, var_t3_also=True, var_sched_t3=False), scheduled=True
    )
    assert opciones.t3_also is False, "programada tiene que ignorar la casilla de pantalla"
