"""Aviso flotante (toast): "Generando…" mientras corre y el resultado al terminar.

Generar Excels corría en el hilo de la pantalla: la ventana quedaba congelada sin decir
nada hasta el cartel final. Ahora corre en segundo plano y el aviso cuenta qué pasa.
"""
import tkinter as tk

import pytest

from osocio.interface import tema
from osocio.interface.widgets import AvisoFlotante


@pytest.fixture
def raiz():
    try:
        r = tk.Tk()
    except tk.TclError:
        pytest.skip("sin pantalla")
    tema.instalar(tema.CLASICO)
    yield r
    r.destroy()


def test_cargando_muestra_el_aviso_con_el_spinner(raiz):
    aviso = AvisoFlotante(raiz)
    aviso.cargando("Generando Excels…", "Argentina · 2 archivo(s)")
    raiz.update()
    assert aviso.caja.winfo_ismapped()
    assert aviso.titulo.cget("text") == "Generando Excels…"
    assert aviso._giro is not None, "el spinner tiene que estar girando"


def test_listo_frena_el_spinner_y_se_va_solo(raiz):
    aviso = AvisoFlotante(raiz)
    aviso.cargando("Generando Excels…")
    aviso.listo("Generar Excels: listo", "• Lead_information_Formulario_Argentina_Chrome_T1.xlsx",
                segundos=0.05)
    assert aviso._giro is None
    raiz.after(150, raiz.quit)
    raiz.mainloop()
    assert not aviso.caja.winfo_ismapped()


def test_el_error_queda_hasta_que_se_cierra(raiz):
    aviso = AvisoFlotante(raiz)
    aviso.error("Generar Excels: no se pudo", "Cerrá los Excel abiertos y volvé a intentar.")
    raiz.update()
    assert aviso._cierre is None and aviso.caja.winfo_ismapped()
    aviso.ocultar()
    raiz.update()
    assert not aviso.caja.winfo_ismapped()
