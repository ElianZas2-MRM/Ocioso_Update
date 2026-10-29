"""Mail de Envío de Leads con el Excel de resultados bloqueado (OneDrive / Excel abierto).
Antes la lectura fallaba en silencio, contaba "0 errores" y el mail decía PASS sin URLs."""
import openpyxl
import pandas as pd
import pytest

import osocio.interface.helpers_interface as h

LANDING = "https://www.chevrolet.com.uy/solicita-una-cotizacion"
FORM = "https://secure-developments.com/commonwealth/uruguay/gm_front/form/raq"


def _excel(path, resultados):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Resultado", "Motivo", "URL", "Formulario"])
    for r in resultados:
        ws.append([r, "-", LANDING, FORM])
    wb.save(path)
    return str(path)


@pytest.fixture
def mail(monkeypatch):
    capturado = {}
    monkeypatch.setattr(h, "_encolar_email", lambda d, asunto, cuerpo, adj=None, callback=None, html_extra="":
                        capturado.update(asunto=asunto, cuerpo=cuerpo, html=html_extra))
    monkeypatch.setattr(h, "cargar_config_global", lambda: {
        "enviar_mail": True, "adjuntar_resultados": False, "adjuntar_screenshots": False})
    monkeypatch.setattr(h, "obtener_email_destinatario", lambda: "nadie@ejemplo.com")
    monkeypatch.setattr(h.time, "sleep", lambda s: None)
    return capturado


def _consolidado(rutas):
    return h.enviar_email_resultados_consolidados(
        [{"pais": p, "navegador": "bs_iphone", "viewport": "iphone", "excel_path": r} for p, r in rutas])


def test_bloqueo_momentaneo_se_reintenta(tmp_path, mail, monkeypatch):
    ruta = _excel(tmp_path / "resultados_Uruguay_BS_iPhone1.xlsx", ["Lead enviado correctamente"])
    original, fallos = pd.read_excel, {"n": 0}

    def read_excel(p, *a, **k):
        if fallos["n"] < 3:
            fallos["n"] += 1
            raise PermissionError(13, "Permission denied")
        return original(p, *a, **k)
    monkeypatch.setattr(h.pd, "read_excel", read_excel)

    assert _consolidado([("Uruguay", ruta)])
    assert mail["asunto"].startswith("[PASS]")
    assert LANDING in mail["html"] and FORM in mail["html"]


def test_excel_ilegible_no_es_pass(tmp_path, mail, monkeypatch):
    ok = _excel(tmp_path / "resultados_Chile_BS_iPhone1.xlsx", ["Lead enviado correctamente"])
    bloqueado = _excel(tmp_path / "resultados_Uruguay_BS_iPhone1.xlsx", ["Lead enviado correctamente"])
    original = pd.read_excel

    def read_excel(p, *a, **k):
        if "Uruguay" in str(p):
            raise PermissionError(13, "Permission denied")
        return original(p, *a, **k)
    monkeypatch.setattr(h.pd, "read_excel", read_excel)

    assert _consolidado([("Chile", ok), ("Uruguay", bloqueado)])
    assert mail["asunto"].startswith("[FAILED]")
    assert "FAILED (1): Uruguay" in mail["cuerpo"] and "PASSED (1): Chile" in mail["cuerpo"]
    assert "No se pudo leer el Excel de resultados" in mail["html"]
    assert "resultados_Uruguay_BS_iPhone1.xlsx" in mail["html"]


def test_excel_sin_filas_procesadas_no_es_pass(tmp_path, mail):
    vacio = _excel(tmp_path / "resultados_Bolivia_BS_iPhone1.xlsx", [])
    assert _consolidado([("Bolivia", vacio)])
    assert mail["asunto"].startswith("[FAILED]")
    assert "ninguna fila procesada" in mail["html"]


def test_fail_real_sale_con_sus_urls(tmp_path, mail):
    ruta = _excel(tmp_path / "resultados_Paraguay_BS_iPhone1.xlsx",
                  ["Lead enviado correctamente", "TY Page no detectada"])
    assert _consolidado([("Paraguay", ruta)])
    assert mail["asunto"].startswith("[FAILED]")
    assert "FAILED (1): Paraguay" in mail["cuerpo"]
    assert mail["html"].count(LANDING) == 2 and "TY Page no detectada" in mail["html"]


def test_mail_por_pais_ilegible_no_es_pass(tmp_path, mail, monkeypatch):
    ruta = _excel(tmp_path / "resultados_Uruguay_Chrome1.xlsx", ["Lead enviado correctamente"])
    monkeypatch.setattr(h.pd, "read_excel", lambda *a, **k: (_ for _ in ()).throw(PermissionError(13, "x")))
    assert h.enviar_email_resultados("Uruguay", ruta, None)
    assert mail["asunto"].startswith("[FAILED]")
    assert "No se pudo leer el Excel de resultados" in mail["cuerpo"]
