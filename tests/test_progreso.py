"""Piezas de las ventanas de "procesando": spinner, barra redondeada y cronómetro.

Reemplazan a lo que ya usaban esas ventanas (el Label con "↻" y el Canvas de la barra)
sin tocar su lógica, así que lo que importa es que acepten las mismas llamadas.
"""
import tkinter as tk

import pytest

from osocio.interface import tema
from osocio.interface.progreso import BarraProgreso, Cronometro, IconoEstado


@pytest.fixture
def raiz():
    try:
        r = tk.Tk()
    except tk.TclError:
        pytest.skip("sin pantalla")
    tema.instalar(tema.CLASICO)
    r.geometry("400x200")
    yield r
    r.destroy()


def test_el_icono_gira_hasta_que_le_ponen_la_marca(raiz):
    icono = IconoEstado(raiz)
    icono.pack()
    assert icono.girando
    icono.config(text="➔")              # lo que hacían las ventanas para animarlo
    assert icono.girando
    icono.config(text="✓", fg="#82E0AA", font=("Segoe UI", 18, "bold"))
    assert not icono.girando and icono._marca == ("✓", "#82E0AA")


def test_la_barra_acepta_lo_que_se_hacia_sobre_el_rectangulo(raiz):
    barra = BarraProgreso(raiz, height=10, bg="#35164D")
    barra.pack(fill="x")
    raiz.update()
    item = barra.create_rectangle(0, 0, 0, 8, fill="#F8C471", width=0)
    barra.coords(item, 0, 0, 120, 8)
    barra.itemconfig(item, fill="#82E0AA")
    assert barra._ancho_lleno == 120 and barra._color == "#82E0AA"
    assert len(barra._imgs) == 2, "pista + parte llena"


def test_el_cronometro_cuenta_mientras_gira_y_despues_da_el_total(raiz):
    icono = IconoEstado(raiz)
    crono = Cronometro(raiz, icono)
    assert "transcurrido" in crono.cget("text")
    icono.config(text="✓", fg="#82E0AA")
    raiz.after(1200, raiz.quit)
    raiz.mainloop()
    assert "en total" in crono.cget("text")
