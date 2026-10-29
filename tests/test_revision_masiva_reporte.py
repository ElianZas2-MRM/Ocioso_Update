"""Reporte de la Revisión Masiva: columna Motivo, filas que no son URL y un Excel por
mercado a medida que terminan. Sin navegador: el revisor de cada fila es falso."""
import os

import openpyxl

from osocio.core import massive_check_runner as mcr
from osocio.core.massive_check_runner import motivo_de_resultado

FORM = "https://secure-developments.com/commonwealth/colombia/gm_front/form/movs-visid-quote?model=n400"
OK_200 = "200 - ✅ PASS Valid link. HTTP Code: 200"


def _res(**kw):
    base = {"landing_status": OK_200, "form_coincide": "✅ PASS", "displayed": "✅ PASS",
            "secure_status": OK_200, "form_url_found": FORM, "error": ""}
    base.update(kw)
    return base


def test_motivo_ok():
    assert motivo_de_resultado(_res()) == "OK"


def test_motivo_landing():
    assert motivo_de_resultado(_res(
        landing_status="404 - ❌ FAIL Landing 404 / no encontrada (título: '404')",
        form_coincide="❌ FAIL No se pudo revisar el form: falló la landing",
        form_url_found="")) == "Landing 404 (no existe)"
    assert motivo_de_resultado(_res(landing_status="503 - ❌ FAIL HTTP Code: 503")).startswith(
        "Landing 503 (error del servidor)")
    m = motivo_de_resultado(_res(
        landing_status=("301 - 🔀 FAIL Link redirects. Code: 301, URL Original: https://a.com/x "
                        "redirects to: https://a.com/"),
        form_coincide=f"❌ FAIL Mismatch: expected {FORM} but found ningún form insertado",
        form_url_found=""))
    assert m == "Landing redirige (301) a https://a.com/"
    assert motivo_de_resultado(_res(landing_status="Timeout - ❌ FAIL La URL no respondió en 8s")
                               ).startswith("Landing no respondió (timeout)")
    # Un CTA que no lleva al form no es un 404.
    m = motivo_de_resultado(_res(landing_status="CTA - ❌ FAIL No se encontró ningún CTA (a.x) en https://a.com",
                                 form_coincide="❌ FAIL No se pudo revisar el form: falló la landing"))
    assert m.startswith("CTA de la landing: No se encontró ningún CTA") and "404" not in m


def test_motivo_form():
    otro = "https://secure-developments.com/commonwealth/colombia/gm_front/form/otro"
    assert motivo_de_resultado(_res(
        form_coincide=f"❌ FAIL Mismatch: expected {FORM} but found {otro}",
        form_url_found=otro, displayed="❌ FAIL")) == f"Form distinto al esperado: en la landing está {otro}"
    assert motivo_de_resultado(_res(
        form_coincide=f"❌ FAIL Mismatch: expected {FORM} but found ningún form insertado",
        form_url_found="", displayed="❌ FAIL")) == "La landing no tiene el form insertado"
    assert motivo_de_resultado(_res(displayed="❌ FAIL")) == "Form correcto pero no se ven sus campos"
    assert motivo_de_resultado(_res(secure_status="404 - ❌ FAIL HTTP Code: 404")) == (
        "La URL del form del Excel da 404")


def _matriz(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "COLOMBIA"
    ws.append(["SEGMENTO", "ESTADO", "URL LIVE", "URL SECURE", "Comentarios"])
    ws.append(["N400", "ON", "https://www.chevrolet.com.co/vans/n400-cargo", FORM, ""])
    ws.append(["Tabla", "ON", "LIVE URL", "FORMS PROD", ""])           # otra tabla en la hoja
    ws.append(["Dia", "ON", "Lunes - ANDROID", "", ""])
    ws.append(["Viejo", "OFF", "https://www.chevrolet.com.co/x", FORM, ""])
    ws2 = wb.create_sheet("CHILE")
    ws2.append(["SEGMENTO", "ESTADO", "URL LIVE", "URL SECURE", "Comentarios"])
    ws2.append(["Solo form", "ON", "Lo usa GM por codigo, no existe en una lading",
                "https://secure-developments.com/commonwealth/chile/gm_forms/vehicle-shopper", ""])
    wb.save(path)


class _FillerFalso:
    revisadas = []

    def __init__(self, country_name, **kw):
        self.driver = self
        self.screenshot_manager = None

    def quit(self):
        pass

    def setup_directories_and_files(self, **kw):
        pass

    def initialize_browser(self):
        pass

    def check_single_url(self, landing, form, row, tomar_capturas=True):
        _FillerFalso.revisadas.append((landing, form))
        return {"ok": True, "current_url": landing or form, "landing_status": OK_200,
                "form_url_found": form, "form_coincide": "✅ PASS", "displayed": "✅ PASS",
                "secure_status": OK_200, "forms_count": 1, "error": ""}


def test_excel_por_mercado_y_filas_sin_url(tmp_path, monkeypatch):
    import osocio.paths
    monkeypatch.setattr(osocio.paths, "BASE_DIR", str(tmp_path))
    monkeypatch.setattr(mcr, "MassiveFormFiller", _FillerFalso)
    _FillerFalso.revisadas = []
    matriz = tmp_path / "Matriz.xlsx"
    _matriz(matriz)
    cols = {"segmento": "SEGMENTO", "estado": "ESTADO", "url_live": "URL LIVE", "url_secure": "URL SECURE"}
    avisos = []

    ok, info = mcr.run_massive_check(str(matriz), cols, ["Colombia", "Chile"], tomar_capturas=False,
                                     progress_callback=lambda p, m, pct: avisos.append((p, m)),
                                     paralelo=True, max_workers=2)
    assert ok, info

    # Solo se revisan las filas con URL; la de Chile, como form suelto.
    assert sorted(_FillerFalso.revisadas) == sorted([
        ("https://www.chevrolet.com.co/vans/n400-cargo", FORM),
        ("", "https://secure-developments.com/commonwealth/chile/gm_forms/vehicle-shopper"),
    ])
    assert info["total_failed"] == 0 and info["total_skipped"] == 1

    # Un Excel por mercado, con solo su hoja, y el aviso de que está listo.
    por_mercado = sorted(os.path.basename(p) for p in info["reportes_por_mercado"])
    assert [n.rsplit("_", 1)[-1] for n in por_mercado] == ["CHILE.xlsx", "COLOMBIA.xlsx"]
    for p in info["reportes_por_mercado"]:
        assert os.path.dirname(p).endswith("por_mercado")
        assert len(openpyxl.load_workbook(p).sheetnames) == 1
    assert ("Colombia", "Listo — Excel del mercado guardado") in avisos

    # El completo tiene la columna Motivo primero.
    ws = openpyxl.load_workbook(info["reporte_path"])["COLOMBIA"]
    assert [c.value for c in ws[1]][:2] == ["Motivo", "currentUrl"]
    assert ws.cell(row=2, column=1).value == "OK"
    assert ws.cell(row=5, column=1).value == "Omitido (OFF)"
    assert ws.cell(row=3, column=1).value is None      # "LIVE URL": no se revisó


def test_reporte_viejo_sin_motivo_no_duplica_el_bloque(tmp_path, monkeypatch):
    import osocio.paths
    monkeypatch.setattr(osocio.paths, "BASE_DIR", str(tmp_path))
    monkeypatch.setattr(mcr, "MassiveFormFiller", _FillerFalso)
    matriz = tmp_path / "Viejo.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "COLOMBIA"
    ws.append(mcr.RESULT_HEADERS[1:] + ["SEGMENTO", "ESTADO", "URL LIVE", "URL SECURE", "Comentarios"])
    ws.append([""] * 7 + ["N400", "ON", "https://www.chevrolet.com.co/vans/n400-cargo", FORM, ""])
    wb.save(matriz)
    cols = {"segmento": "SEGMENTO", "estado": "ESTADO", "url_live": "URL LIVE", "url_secure": "URL SECURE"}

    ok, info = mcr.run_massive_check(str(matriz), cols, ["Colombia"], tomar_capturas=False)
    assert ok, info
    cab = [c.value for c in openpyxl.load_workbook(info["reporte_path"])["COLOMBIA"][1]]
    assert cab[:8] == mcr.RESULT_HEADERS and cab.count("currentUrl") == 1
