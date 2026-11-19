"""Regresión: si la TYP ya se confirmó, una caída de conexión o error de captura
post-envío no puede convertir el lead en FAIL ni disparar reintentos."""

from unittest.mock import MagicMock

from selenium.webdriver.common.by import By

from osocio.core.base_form_filler import BaseFormFiller


def test_ty_confirmada_no_se_pierde_si_cae_conexion_al_capturar(monkeypatch):
    filler = BaseFormFiller.__new__(BaseFormFiller)
    filler.config = {"capturas": True}
    filler.SCREENSHOT_DIR = "."

    btn = MagicMock()
    filler._resolve_submit_button = MagicMock(return_value=(btn, "button[type='submit']"))
    filler._sync_tracked_with_dom_before_submit = MagicMock()
    filler._scroll_element_into_view = MagicMock()
    filler.frenar_si_pidieron_detener = MagicMock()
    filler._detect_event_id_error = MagicMock(return_value=False)

    ty_el = MagicMock()
    ty_el.get_attribute.return_value = "display: block;"

    driver = MagicMock()
    driver.find_element.side_effect = lambda by, sel: ty_el if sel == "div#thank-you" else None
    # Simula pérdida de conexión justo después de detectar la TYP, al salir del iframe.
    driver.switch_to.default_content.side_effect = RuntimeError("Connection reset by peer")
    filler.driver = driver

    filler.reposition_to_form = MagicMock()
    filler.screenshot_manager = MagicMock()

    result_text, ty_page_name = filler.submit_and_verify_form(1, "https://example.com/form")

    assert result_text == "Lead enviado correctamente"
    assert ty_page_name is not None
    assert ty_page_name == ""
