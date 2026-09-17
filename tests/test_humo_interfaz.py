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
import threading
import tkinter as tk

import pytest


@pytest.fixture
def tk_sin_ventana(monkeypatch):
    """Neutraliza `mainloop`, esconde la ventana y deja Tk limpio al terminar.

    Sin lo primero el test se colgaría para siempre en el loop de eventos de Tk.

    Lo de limpiar al final no es cosmético: Tk guarda las imágenes en un registro GLOBAL
    del proceso. Si otro test dejó un root vivo, este falla con "pyimage2 doesn't exist" —
    y falla *según el orden en que corran*, que es la peor clase de test: el que se ignora
    porque a veces pasa.
    """
    llamadas = {"mainloop": 0}

    # Si quedó un root de otro test, sacarlo de en medio antes de empezar.
    if tk._default_root is not None:
        try:
            tk._default_root.destroy()
        except Exception:
            pass
        tk._default_root = None

    def _mainloop_falso(self, *a, **k):
        llamadas["mainloop"] += 1

    monkeypatch.setattr(tk.Misc, "mainloop", _mainloop_falso)
    monkeypatch.setattr(tk.Tk, "mainloop", _mainloop_falso, raising=False)

    _init_original = tk.Tk.__init__

    def _init_escondido(self, *a, **k):
        _init_original(self, *a, **k)
        self.withdraw()

    monkeypatch.setattr(tk.Tk, "__init__", _init_escondido)

    # No dejar arrancar los hilos de fondo. iniciar_interfaz() lanza 3 daemons, y uno
    # de ellos (_sched_monitor) es un `while True` que le pega a la ventana con
    # root.after(). Cuando el test destruye el root, ese hilo sigue vivo y le escribe a
    # una ventana muerta: el test fallaba 1 de cada 5 corridas, sin patron.
    #
    # Neutralizarlos ademas enfoca mejor lo que se quiere probar: que la UI se ARME.
    # Los loops de fondo son otra cosa y merecen sus propios tests.
    monkeypatch.setattr(threading.Thread, "start", lambda self: None)

    yield llamadas

    # Dejar el proceso como estaba, para no arruinarle el Tk al test que venga despues.
    if tk._default_root is not None:
        try:
            tk._default_root.destroy()
        except Exception:
            pass
        tk._default_root = None


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



def _saltear_si_es_culpa_del_entorno(exc):
    """Distingue un Tk roto por el entorno de un bug nuestro.

    Este equipo corre el Python de la Microsoft Store, que vive bajo `WindowsApps` con el
    sistema de archivos virtualizado. Cada tanto no puede leer su propio `init.tcl` y tira
    un TclError cuyo mensaje de error dice, literalmente, "No error".

    Eso hacia fallar la suite 1 de cada 5 corridas sin ningun patron. Un test que falla a
    veces se termina ignorando, asi que se saltea SOLO ante ese error puntual: cualquier
    otro TclError sigue siendo un fallo de verdad.

    La solucion de fondo es instalar un Python normal (python.org) en vez del de la Store.
    """
    texto = str(exc)
    # La firma del problema: un TclError sobre un archivo .tcl que vive bajo WindowsApps.
    # Se vio con init.tcl y con auto.tcl, asi que se matchea el patron y no el nombre.
    es_del_entorno = (
        ("WindowsApps" in texto and ".tcl" in texto)
        or "Can't find a usable" in texto
    )
    if es_del_entorno:
        pytest.skip(
            "Tk no pudo inicializarse: el Python de la Microsoft Store no siempre puede "
            "leer su init.tcl. Es del entorno, no del codigo."
        )
    raise exc


@requiere_display
def test_la_interfaz_se_construye_entera(tk_sin_ventana):
    """El test que más cubre de todo el proyecto: 4.942 líneas de armado de UI.

    Si algo del armado está roto, esto falla acá y no cuando el usuario abre la app.
    """
    from osocio.interface.main_interface import iniciar_interfaz

    try:
        iniciar_interfaz()
    except tk.TclError as exc:
        _saltear_si_es_culpa_del_entorno(exc)

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

    if tk._default_root is not None:
        try:
            tk._default_root.destroy()
        except Exception:
            pass
        tk._default_root = None

    try:
        root = tk.Tk()
    except tk.TclError as exc:
        _saltear_si_es_culpa_del_entorno(exc)

    root.withdraw()
    try:
        panel = _DemoSchedulerPanel(root, on_save=lambda *a, **k: None)
        assert panel is not None
    finally:
        root.destroy()
        tk._default_root = None


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
