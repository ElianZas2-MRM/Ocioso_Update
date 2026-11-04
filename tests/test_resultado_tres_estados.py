"""Resultado en tres estados: OK (verde), AVISO (naranja) y ERROR (rojo).

Pasó en Perú (evento-agrofest) y en Brasil (eléctricos, en Edge): el form esperado no
aparecía a tiempo, el motor tomaba "cualquier iframe GM" de la landing, el lead se llenaba
en OTRO form y la corrida salía PASS aunque la columna "Form coincide" dijera FAIL.

Ahora cada fila se clasifica una sola vez y de ahí sale todo: el Excel, el resumen del
final, el mail y "Reintentar fallidos".
  - ERROR: algo de los campos o del envío en el form correcto.
  - AVISO: problemas de URL / form / TYP (404, form distinto o no inserto, link de la TYP,
    TYP que no se ve) — que se vean a simple vista y no se pasen.
"""
import pandas as pd
import pytest
from openpyxl import Workbook
from openpyxl.styles import PatternFill

from osocio.core.base_form_filler import BaseFormFiller

clasificar = BaseFormFiller.clasificar_fila


# --- clasificación de una fila ---------------------------------------------------------------

def test_todo_bien_es_ok():
    assert clasificar(envio_ok=True) == ("ok", [], [])


def test_form_inserto_distinto_con_lead_enviado_es_aviso_no_pass():
    """El caso de Perú: el lead viajó, pero por otro form. No puede quedar verde."""
    estado, errores, avisos = clasificar(envio_ok=True, form_estado="inserto_otro",
                                         url_problema="form 404 NO ENCONTRADA (404)")
    assert estado == "aviso"
    assert errores == []
    assert any("distinto" in a for a in avisos)
    assert any("404" in a for a in avisos)


def test_campos_en_el_form_correcto_son_error():
    estado, errores, _ = clasificar(envio_ok=True,
                                    campos_con_error=["dropdown no encontrado: dealer"])
    assert estado == "error"
    assert errores == ["dropdown no encontrado: dealer"]


def test_el_form_rechazo_el_envio_es_error():
    estado, errores, _ = clasificar(envio_ok=False, motivo_envio="event_id")
    assert estado == "error" and errores == ["event_id"]


@pytest.mark.parametrize("kw", [
    dict(envio_ok=False, landing_caida=True, url_problema="landing 404 NO ENCONTRADA (404)"),
    dict(envio_ok=False, form_estado="no_inserto", motivo_envio="form no inserto"),
    dict(envio_ok=True, link_issue=True),
    dict(envio_ok=False, motivo_envio="sin TYP"),
    dict(envio_ok=True, cta_eletricos="FAIL — 404"),
    dict(envio_ok=True, dif_menor=["mayúsculas"]),
])
def test_problemas_de_url_form_o_typ_son_aviso(kw):
    assert clasificar(**kw)[0] == "aviso"


def test_con_error_de_campos_y_404_a_la_vez_manda_el_rojo():
    estado, errores, avisos = clasificar(envio_ok=True, url_problema="form 404",
                                         campos_con_error=["campos sin completar: cpf"])
    assert estado == "error"
    assert errores and avisos, "el motivo tiene que decir las dos cosas"


def test_en_otro_form_los_campos_que_faltan_son_consecuencia_no_error():
    """Si el form es otro, que no encuentre el dropdown es lógico: el problema es el form."""
    estado, errores, avisos = clasificar(envio_ok=False, form_estado="inserto_otro",
                                         campos_con_error=["dropdown no encontrado: dealer"])
    assert estado == "aviso" and errores == []
    assert any("dropdown" in a for a in avisos)


# --- "Estado URL form": esperada e inserta -----------------------------------------------------

class _Filler:
    _estado_urls_form = BaseFormFiller._estado_urls_form
    _diferencia_url_form = staticmethod(BaseFormFiller._diferencia_url_form)

    def __init__(self, estado_esperada):
        self._estado_url_form = estado_esperada


def test_estado_url_form_muestra_tambien_la_inserta(monkeypatch):
    from osocio.utils import url_status
    monkeypatch.setattr(url_status, "check_url_status", lambda u: {"label": "200 OK"})
    f = _Filler("404 NO ENCONTRADA (404)")
    txt = f._estado_urls_form("https://x.com/peru/gm_forms/agrofest",
                              "https://x.com/peru/gm_front/form/agrofest")
    assert txt == "Esperada: 404 NO ENCONTRADA (404) | Inserta: 200 OK"


def test_misma_url_un_solo_estado(monkeypatch):
    from osocio.utils import url_status
    monkeypatch.setattr(url_status, "check_url_status",
                        lambda u: pytest.fail("no hace falta consultar de nuevo"))
    f = _Filler("200 OK")
    assert f._estado_urls_form("https://x.com/f/raq", "https://x.com/f/raq/") == "200 OK"


@pytest.mark.parametrize("valor, color", [
    ("200 OK", "C6EFCE"),
    ("Esperada: 404 NO ENCONTRADA (404) | Inserta: 200 OK", "FFC7CE"),
    ("Esperada: 200 OK | Inserta: 200 OK", "C6EFCE"),
    ("REDIRIGE (301)", "FFC000"),
    ("-", None),
])
def test_color_del_estado_url_toma_el_peor(valor, color):
    assert BaseFormFiller._color_estado_url(valor) == color


# --- espera extra por el iframe esperado --------------------------------------------------------

class _Driver:
    class _SwitchTo:
        def default_content(self):
            pass

    def __init__(self, srcs):
        self.srcs = srcs
        self.switch_to = self._SwitchTo()

    def find_elements(self, *_):
        return list(self.srcs)


class _FillerEspera:
    _esperar_iframe_esperado = BaseFormFiller._esperar_iframe_esperado
    _diferencia_url_form = staticmethod(BaseFormFiller._diferencia_url_form)

    def __init__(self, srcs):
        self.driver = _Driver(srcs)

    def _src_y_visibilidad_de_iframes(self, iframes):
        return [(s, True) for s in iframes]


def test_la_espera_extra_encuentra_el_esperado_y_no_otro_gm():
    f = _FillerEspera(["https://x.com/brasil/gm_forms/raq", "https://x.com/brasil/gm_forms/raq-eletricos"])
    assert f._esperar_iframe_esperado("https://x.com/brasil/gm_forms/raq-eletricos", 0.3) \
        == "https://x.com/brasil/gm_forms/raq-eletricos"


def test_si_el_esperado_no_esta_la_espera_devuelve_nada():
    f = _FillerEspera(["https://x.com/brasil/gm_forms/raq"])
    assert f._esperar_iframe_esperado("https://x.com/brasil/gm_forms/raq-eletricos", 0.3) is None


# --- mail y reintento leen lo mismo --------------------------------------------------------------

def _excel(tmp_path, filas, con_estado=True):
    cols = (["Estado"] if con_estado else []) + ["Resultado", "Motivo", "Form coincide", "URL", "Formulario"]
    df = pd.DataFrame(filas, columns=cols)
    p = tmp_path / "resultados_Peru_Edge1.xlsx"
    df.to_excel(p, index=False)
    return str(p)


def test_el_mail_cuenta_los_avisos_aparte(tmp_path):
    from osocio.interface.helpers_interface import analizar_errores_excel
    p = _excel(tmp_path, [
        ["OK", "Lead enviado correctamente", "-", "PASS", "https://a/1", "https://f/1"],
        ["AVISO", "Lead enviado correctamente", "AVISO: form inserto distinto", "FAIL — form inserto no coincide con el esperado", "https://a/2", "https://f/2"],
        ["ERROR", "Error visual: cpf", "ERROR: validación", "PASS", "https://a/3", "https://f/3"],
    ])
    r = analizar_errores_excel(p)
    assert (r["exitosos"], r["con_avisos"], r["con_errores"]) == (1, 1, 1)
    assert r["detalles_avisos"][0]["linea"] == 3


def test_un_excel_viejo_sin_estado_se_analiza_como_antes(tmp_path):
    from osocio.interface.helpers_interface import analizar_errores_excel
    p = _excel(tmp_path, [
        ["Lead enviado correctamente", "-", "PASS", "https://a/1", "https://f/1"],
        ["Lead enviado correctamente", "-", "FAIL — form inserto no coincide con el esperado", "https://a/2", "https://f/2"],
    ], con_estado=False)
    r = analizar_errores_excel(p)
    assert (r["exitosos"], r["con_errores"], r["con_avisos"]) == (1, 1, 0)


def test_la_tabla_del_mail_muestra_aviso_en_naranja():
    from osocio.interface.helpers_interface import _build_url_table_html
    html = _build_url_table_html([
        {"linea": 2, "url": "https://a/2", "url_secure": "https://f/2", "ok": False, "aviso": True,
         "error": "form inserto distinto"},
        {"linea": 3, "url": "https://a/3", "url_secure": "https://f/3", "ok": True, "error": ""},
    ])
    assert "AVISO" in html and "#b9770e" in html and "PASS" in html


def test_reintentar_toma_el_naranja_sin_envio_pero_no_el_enviado(tmp_path):
    from osocio.utils.reintento_fallidos import (NARANJA_ENVIADO, NARANJA_SIN_ENVIO, ROJO_FALLA,
                                                 urls_fallidas)
    wb = Workbook()
    ws = wb.active
    ws.append(["Resultado", "URL", "Formulario"])
    for color, n in ((ROJO_FALLA, 1), (NARANJA_SIN_ENVIO, 2), (NARANJA_ENVIADO, 3), ("C6EFCE", 4)):
        ws.append([f"r{n}", f"https://a/{n}", f"https://f/{n}"])
        ws.cell(row=n + 1, column=1).fill = PatternFill(fill_type="solid", fgColor=color)
    p = tmp_path / "res.xlsx"
    wb.save(p)
    urls = [u for u, _f in urls_fallidas(str(p))]
    assert urls == ["https://a/1", "https://a/2"]
