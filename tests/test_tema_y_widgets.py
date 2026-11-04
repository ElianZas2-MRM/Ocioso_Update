"""Tema oscuro/clásico y los controles redondeados.

La interfaz está escrita con los colores clásicos y el tema los traduce al pintar. Lo que
importa probar: que la traducción no rompa al código que lee colores de vuelta (`cget`),
que dos colores distintos no terminen iguales, y que los botones y marcos nuevos se usen
igual que los de Tk.
"""
import tkinter as tk

import pytest

from osocio.interface import tema


@pytest.fixture
def raiz():
    try:
        r = tk.Tk()
    except tk.TclError:
        pytest.skip("sin pantalla")
    r.withdraw()
    yield r
    tema.instalar(tema.CLASICO)
    r.destroy()


def test_clasico_no_cambia_nada():
    tema.instalar(tema.CLASICO)
    assert tema.color("#4A2666") == "#4A2666"
    assert tema.clasico("#4A2666") == "#4A2666"


def test_oscuro_traduce_y_cget_devuelve_el_clasico(raiz):
    tema.instalar(tema.OSCURO)
    f = tk.Frame(raiz, bg="#4A2666")
    assert f.tk.call(f._w, "cget", "-bg") != "#4A2666"   # lo que pinta Tk: el oscuro
    assert f.cget("bg") == "#4A2666"                    # lo que ve el código: el clásico
    assert f["bg"] == "#4A2666"


def test_dos_colores_nunca_dan_el_mismo():
    """Si dos clásicos dieran el mismo oscuro, `cget` no sabría cuál devolver."""
    tema.instalar(tema.OSCURO)
    clasicos = ["#5D3C7A", "#4A2666", "#7D4E9F", "#35164D", "#602D8A", "#4C226E", "#2E1146",
                "#E6D6F2", "#C5B2D6", "#C5A9DF", "#3A1D52", "#9B5CD6", "#24242A", "#303037",
                "#36363E", "#3A3A42", "#0F0F12"]
    oscuros = [tema.color(c) for c in clasicos]
    assert len(set(oscuros)) == len(clasicos)
    tema.instalar(tema.CLASICO)


def test_los_colores_de_estado_no_se_tocan():
    """Verde de ejecutar, coral de eliminar: se leen igual en los dos temas."""
    tema.instalar(tema.OSCURO)
    for c in ("#27AE60", "#F1948A", "#82E0AA", "#E2B33C", "#3DC1E3"):
        assert tema.color(c).lower() == c.lower()
    tema.instalar(tema.CLASICO)


def test_en_oscuro_las_entradas_blancas_pasan_a_oscuras():
    tema.instalar(tema.OSCURO)
    assert tema.color_opcion("bg", "white") == tema.ENTRADA_OSCURA
    assert tema.color_opcion("fg", "black") == tema.TEXTO_ENTRADA_OSCURA
    assert tema.color_opcion("fg", "white") == "white"      # texto blanco: queda
    tema.instalar(tema.CLASICO)


def test_cambio_en_vivo_ida_y_vuelta(raiz):
    tema.instalar(tema.CLASICO)
    lbl = tk.Label(raiz, bg="#4A2666", fg="#E6D6F2")
    tema.cambiar(tema.OSCURO, raiz)
    assert lbl.tk.call(lbl._w, "cget", "-bg") == tema.color("#4A2666")
    assert lbl.cget("bg") == "#4A2666"
    tema.cambiar(tema.CLASICO, raiz)
    assert lbl.tk.call(lbl._w, "cget", "-bg") == "#4A2666"


def test_boton_redondo_se_usa_como_un_button(raiz):
    from osocio.interface.widgets import BotonRedondo
    tema.instalar(tema.CLASICO)
    clics = []
    b = BotonRedondo(raiz, text="Guardar", bg="#35164D", fg="white", padx=10, pady=4,
                     command=lambda: clics.append(1))
    b.pack()
    assert b.cget("text") == "Guardar"
    assert b["bg"] == "#35164D"
    b.config(bg="#602D8A", text="Guardado")
    assert b.cget("bg") == "#602D8A" and b["text"] == "Guardado"
    b.invoke()
    assert clics == [1]
    b.config(state="disabled")
    assert b["state"] == "disabled"
    b.invoke()
    assert clics == [1], "deshabilitado no ejecuta"


def test_marco_con_borde_es_tarjeta_y_conserva_sus_colores(raiz):
    from osocio.interface.widgets import Marco
    tema.instalar(tema.CLASICO)
    tarjeta = Marco(raiz, bg="#4A2666", highlightthickness=1, highlightbackground="#7D4E9F")
    assert tarjeta.cget("bg") == "#4A2666"
    assert tarjeta.cget("highlightbackground") == "#7D4E9F"
    tarjeta.config(highlightbackground="#AED6F1")
    assert tarjeta["highlightbackground"] == "#AED6F1"
    simple = Marco(raiz, bg="#5D3C7A")
    assert simple.cget("bg") == "#5D3C7A"


def test_menu_lateral_marca_la_pagina(raiz):
    from osocio.interface.menu_lateral import MenuLateral
    tema.instalar(tema.CLASICO)
    elegidas = []
    modo = tk.StringVar(value="T1")
    menu = MenuLateral(raiz, "Asset", [("Envíos", [("leads", "Envío de Leads", "play_white.png")]),
                                       ("Herramientas", [("ids", "IDs Dinámicos", "copy_lavender.png")])],
                       al_elegir=elegidas.append, var_modo=modo, al_cambiar_modo=None,
                       al_cambiar_tema=None, al_actualizar_drivers=None)
    menu.pack()
    menu.botones["ids"].invoke()
    assert elegidas == ["ids"]
    menu.marcar("ids")
    assert menu.botones["ids"].cget("bg") == tema.MENU_SEL
    menu._alternar()          # tira de íconos
    assert menu.compacto and set(menu.botones) == {"leads", "ids"}
    menu._alternar_modo()
    assert modo.get() == "T3"
