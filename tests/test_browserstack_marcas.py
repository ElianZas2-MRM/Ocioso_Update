"""Marcas por lead en el video de BrowserStack y motivo del estado de la sesión.
Sin red: el driver es falso y solo guarda los comandos que recibe."""
import json

from osocio.core.base_form_filler import BaseFormFiller
from osocio.providers.browserstack import runner
from osocio.providers.browserstack.bs_driver import annotate_script


def _payload(script):
    prefijo = "browserstack_executor: "
    assert script.startswith(prefijo)
    return json.loads(script[len(prefijo):])


def test_annotate_script_con_nivel_y_texto_recortado():
    p = _payload(annotate_script("Lead 3/10 · fila 4 · FALLÓ: sin TYP", "error"))
    assert p == {"action": "annotate",
                 "arguments": {"data": "Lead 3/10 · fila 4 · FALLÓ: sin TYP", "level": "error"}}
    assert len(_payload(annotate_script("x" * 400))["arguments"]["data"]) == 255


def test_texto_de_las_marcas():
    url = "https://www.chevrolet.com.co/vans/n400-cargo?utm=1"
    assert runner.texto_marca_lead("inicio", 3, 10, 4, url) == "Lead 3/10 · fila 4 · empieza · vans/n400-cargo"
    assert runner.texto_marca_lead("fin", 3, 10, 4, ok=True) == "Lead 3/10 · fila 4 · enviado OK"
    assert runner.texto_marca_lead("fin", 5, 10, 6, ok=False, motivo="sin TYP") == \
        "Lead 5/10 · fila 6 · FALLÓ: sin TYP"
    # Sin envío (chequeo de campos) no es un fallo.
    assert runner.texto_marca_lead("fin", 5, 10, 6, ok=False, enviar=False) == \
        "Lead 5/10 · fila 6 · revisado sin enviar"


class _Driver:
    def __init__(self):
        self.scripts = []

    def execute_script(self, script):
        self.scripts.append(_payload(script)["arguments"])


class _Sesion:
    def __init__(self):
        self.driver = _Driver()


class _Fabrica:
    sesion = None


def test_marcas_van_al_driver_de_la_sesion():
    fab = _Fabrica()
    marcas = runner.MarcasDelVideo(fab)
    marcas("inicio", fila=2, numero=1, total=2, url="https://a.com/x")   # sin sesión: no hace nada
    fab.sesion = _Sesion()
    marcas("fin", fila=2, numero=1, total=2, ok=True)
    marcas("fin", fila=3, numero=2, total=2, ok=False, motivo="landing 404")
    assert fab.sesion.driver.scripts == [
        {"data": "Lead 1/2 · fila 2 · enviado OK", "level": "info"},
        {"data": "Lead 2/2 · fila 3 · FALLÓ: landing 404", "level": "error"},
    ]
    sin_envio = runner.MarcasDelVideo(fab, enviar=False)
    sin_envio("fin", fila=3, numero=2, total=2, ok=False)
    assert fab.sesion.driver.scripts[-1]["level"] == "info"


def test_un_fallo_de_la_marca_no_corta_la_corrida():
    class _Motor:
        _avisar_lead = BaseFormFiller._avisar_lead

        def on_lead(self, evento, **datos):
            raise RuntimeError("BrowserStack no respondió")
    _Motor()._avisar_lead("fin", fila=2)            # no levanta
    sin_hook = type("M", (), {"on_lead": None, "_avisar_lead": BaseFormFiller._avisar_lead})()
    sin_hook._avisar_lead("inicio", fila=2)


def test_motivo_de_la_sesion():
    assert runner.motivo_de_sesion(10, 0) == "10 de 10 leads enviados OK"
    assert runner.motivo_de_sesion(0, 0) == "No se procesó ningún lead"
    assert runner.motivo_de_sesion(1, 0, error="Timeout") == "Error de la corrida: Timeout"
    filas = [{"row": 4, "reason": "sin TYP"}, {"row": 7, "reason": "sin TYP"},
             {"row": 9, "reason": "dropdown no encontrado"}]
    assert runner.motivo_de_sesion(7, 3, filas) == (
        "3 de 10 leads fallaron: sin TYP (filas 4, 7); dropdown no encontrado (fila 9)")
    assert runner.motivo_de_sesion(9, 1) == "1 de 10 leads fallaron"


def test_el_motivo_del_estado_va_sin_tildes():
    from osocio.providers.browserstack.bs_driver import status_script
    p = _payload(status_script(False, "No se procesó ningún lead · envío"))
    assert p["arguments"]["reason"] == "No se proceso ningun lead  envio"
