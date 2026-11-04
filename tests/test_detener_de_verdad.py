"""'Detener' (o cerrar la corrida y confirmar) frena el envío de verdad.

Antes solo evitaba que arrancaran sesiones nuevas: la sesión en curso seguía con todas sus
filas y mandaba los leads igual. Ahora el motor mira el pedido justo antes de tocar Enviar.
"""
import threading

import pytest

from osocio.core.base_form_filler import (MENSAJE_DETENIDO, BaseFormFiller, DetenidoPorUsuario)


class _Boton:
    def __init__(self):
        self.clics = 0

    def click(self):
        self.clics += 1


class _Motor(BaseFormFiller):
    """Solo lo necesario para llegar al clic de Enviar, sin navegador."""

    def __init__(self, stop_event):
        self.config = {}
        self.stop_event = stop_event
        self.boton = _Boton()

    def _sanitize_url(self, u):
        return u

    def _resolve_submit_button(self, wait_seconds=3):
        return self.boton, "#enviar"

    def _sync_tracked_with_dom_before_submit(self, _datos):
        pass

    def _scroll_element_into_view(self, _el):
        pass


def test_con_detener_pedido_no_se_toca_enviar():
    ev = threading.Event()
    ev.set()
    motor = _Motor(ev)
    with pytest.raises(DetenidoPorUsuario):
        motor.submit_and_verify_form(1, "https://x/form")
    assert motor.boton.clics == 0, "el lead no puede salir"


def test_sin_detener_se_envia_como_siempre():
    motor = _Motor(threading.Event())
    try:
        motor.submit_and_verify_form(1, "https://x/form")
    except Exception:
        pass      # después del clic no hay navegador para verificar la TYP: no importa acá
    assert motor.boton.clics == 1


def test_sin_ventana_de_corrida_nunca_frena():
    motor = _Motor(None)
    assert not motor.pidieron_detener()


def test_la_fila_detenida_queda_como_aviso_para_reintentar():
    motivo = BaseFormFiller._short_fail_reason(MENSAJE_DETENIDO)
    estado, errores, avisos = BaseFormFiller.clasificar_fila(envio_ok=False, motivo_envio=motivo)
    assert estado == "aviso" and not errores
    assert any("detenido" in a for a in avisos)
