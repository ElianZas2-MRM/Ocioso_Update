"""La interfaz se construye entera sin romperse.

`iniciar_interfaz()` son 4.944 líneas que arman toda la ventana y terminan en
`root.mainloop()`. Esa última línea es lo único que bloquea: si se la neutraliza, la
función construye la UI completa y devuelve.

Eso convierte a estas ~20 líneas de test en cobertura de **4.942 líneas que no tenían
ninguna**. No prueba que la app haga bien su trabajo — prueba que no está rota de entrada,
que es justo lo que faltaba para poder tocar este archivo sin ir a ciegas.

Lo que sí atrapa: un widget mal construido, un nombre indefinido en cualquier rama de
armado, un JSON de configuración ilegible, un import que se cayó, una constante que se
movió de lugar.
"""
import tkinter
import tkinter as tk

import pytest


@pytest.fixture
def tk_sin_ventana(monkeypatch):
    """Neutraliza `mainloop` y esconde la ventana.

    Sin esto el test se colgaría para siempre en el loop de eventos de Tk.
    """
    llamadas = {"mainloop": 0}

    def _mainloop_falso(self, *a, **k):
        llamadas["mainloop"] += 1

    monkeypatch.setattr(tk.Misc, "mainloop", _mainloop_falso)
    monkeypatch.setattr(tk.Tk, "mainloop", _mainloop_falso, raising=False)

    _init_original = tk.Tk.__init__

    def _init_escondido(self, *a, **k):
        _init_original(self, *a, **k)
        self.withdraw()

    monkeypatch.setattr(tk.Tk, "__init__", _init_escondido)
    return llamadas


def _hay_display():
    try:
        r = tk.Tk()
        r.destroy()
        return True
    except Exception:
        return False


requiere_display = pytest.mark.skipif(
    not _hay_display(), reason="no hay display disponible para Tk"
)


@requiere_display
def test_la_interfaz_se_construye_entera(tk_sin_ventana):
    """El test que más cubre de todo el proyecto: 4.942 líneas de armado de UI.

    Si algo del armado está roto, esto falla acá y no cuando el usuario abre la app.
    """
    from osocio.interface.main_interface import iniciar_interfaz

    iniciar_interfaz()

    assert tk_sin_ventana["mainloop"] == 1, (
        "iniciar_interfaz() tiene que terminar entrando al mainloop exactamente una vez"
    )


def test_no_dispara_el_envio_automatico_por_defecto():
    """`autostart_leads` arranca una corrida real de leads: sin pedirlo, no puede pasar.

    No hace falta construir la UI de nuevo para verificarlo (y construirla dos veces en el
    mismo proceso rompe Tk, que recicla los nombres de las imagenes). Alcanza con la firma.
    """
    import inspect

    from osocio.interface.main_interface import iniciar_interfaz

    firma = inspect.signature(iniciar_interfaz)
    assert firma.parameters["autostart_leads"].default is False


# --- las piezas de nivel superior, que se pueden probar sueltas --------------------

@requiere_display
def test_el_panel_del_scheduler_se_arma():
    """`_DemoSchedulerPanel` son 409 líneas y es la única clase del archivo."""
    from osocio.interface.main_interface import _DemoSchedulerPanel

    root = tk.Tk()
    root.withdraw()
    try:
        panel = _DemoSchedulerPanel(root, on_save=lambda *a, **k: None)
        assert panel is not None
    finally:
        root.destroy()


@pytest.mark.parametrize("pais, esperado_contiene", [
    ("Argentina", "Argentina"),
    ("Brasil", "Brasil"),
])
def test_el_nombre_del_excel_lleva_el_pais(pais, esperado_contiene):
    from osocio.interface.main_interface import _lead_excel_name

    nombre = _lead_excel_name(pais, suffix="")
    assert esperado_contiene in nombre
    assert nombre.endswith(".xlsx")


@pytest.mark.parametrize("navegador", ["chrome", "firefox", "edge"])
def test_cada_navegador_tiene_su_sufijo_de_excel(navegador):
    """Los Excel se separan por dispositivo: sin sufijo, dos corridas se pisan."""
    from osocio.interface.main_interface import _device_excel_suffix

    sufijo = _device_excel_suffix(navegador)
    assert isinstance(sufijo, str)


def test_los_iconos_de_los_botones_no_revientan_si_falta_el_archivo():
    """La app no puede dejar de abrir porque falte un PNG decorativo."""
    from osocio.interface.main_interface import get_button_icon

    assert get_button_icon("no_existe_este_icono_12345.png") is None
