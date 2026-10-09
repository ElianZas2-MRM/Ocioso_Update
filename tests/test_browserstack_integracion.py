"""BrowserStack integrado a la app, sin red.

El riesgo de esta integración no es que algo explote al importar: es que un dispositivo
quede mal cableado entre la UI, el despacho y los resultados (y el mail salga sin
adjunto), o que el iPhone —que llena por la capa nativa— se aparte de las reglas de datos
de la app. Estos tests vigilan eso.
"""
import json
import os

import pytest

from osocio.core.campos_inferencia import ids_para_etiqueta
from osocio.core.variantes import basenames_de_corrida
from osocio.providers.browserstack import dispositivos, runner
from osocio.providers.browserstack.ios_nativo import OMITIR


# --- tabla única de dispositivos ----------------------------------------------------

def test_los_tres_dispositivos_y_sus_caras():
    assert [d.etiqueta for d in dispositivos.DISPOSITIVOS_BS] == ["iPhone BS", "Android BS", "Mac BS"]
    iphone = dispositivos.por_cualquier_nombre("iPhone BS")
    for nombre in ("iphone bs", "iphone", "bs_iphone", "browserstack_iphone"):
        assert dispositivos.por_cualquier_nombre(nombre) is iphone
    assert iphone.sufijo_excel == "Iphone"
    assert iphone.token_email == "bsiphone"


def test_un_navegador_local_no_es_de_la_nube():
    for nombre in ("chrome", "firefox", "edge", "", None):
        assert not dispositivos.es_browserstack(nombre)


def test_la_seleccion_guardada_con_lambdatest_pasa_a_browserstack():
    viejo = {"chrome": True, "mac lt": True, "android lt": False}
    assert dispositivos.migrar_seleccion(viejo) == {"chrome": True, "mac bs": True,
                                                    "android bs": False}


@pytest.mark.parametrize("configurado,esperado", [(None, 3), (1, 1), (5, 5), (50, 10), (0, 1)])
def test_tope_de_sesiones_en_paralelo(monkeypatch, tmp_path, configurado, esperado):
    cfg = {} if configurado is None else {"browserstack": {"max_paralelo": configurado}}
    (tmp_path / "config_global.json").write_text(json.dumps(cfg), encoding="utf-8")
    monkeypatch.setattr("osocio.paths.JSON_DIR", str(tmp_path))
    assert dispositivos.max_paralelo() == esperado


def test_el_excel_por_dispositivo_es_el_que_usa_la_ui():
    from osocio.interface.main_interface import _device_excel_suffix
    assert _device_excel_suffix("iPhone BS") == "Iphone"
    assert _device_excel_suffix("Android BS") == "Android"
    assert _device_excel_suffix("Mac BS") == "Mac"
    assert _device_excel_suffix("Chrome") == "Chrome"
    assert runner.excel_por_defecto("Brasil", "mac", "_T3").endswith(
        "Lead_information_Formulario_Brasil_Mac_T3.xlsx")


# --- resultados: mismas carpetas y nombres que los navegadores ------------------------

@pytest.mark.parametrize("navegador,sufijo,programada,esperado", [
    ("browserstack_mac", "", False, "resultados_Brasil_BS_Mac"),
    ("browserstack_iphone", "_CDC", False, "resultados_Brasil_CDC_BS_iPhone"),
    ("browserstack_android", "_T3", True, "Automatizacion_Brasil_T3_BS_Android"),
])
def test_nombre_de_los_resultados(navegador, sufijo, programada, esperado):
    excel, _capturas = basenames_de_corrida("Brasil", navegador, programada=programada, sufijo=sufijo)
    assert excel == esperado


def _solo_caracteres_que_acepta_browserstack(nombre):
    # Letras sin acento, dígitos, espacio y . : - [ ] / @ & ' _ (lo demás lo descarta).
    import re
    return re.fullmatch(r"[A-Za-z0-9 .:\-\[\]/@&'_]+", nombre) is not None


def test_un_build_por_mercado_y_por_corrida():
    from datetime import datetime
    inicio = datetime(2026, 10, 9, 10, 32)
    nombre = runner.nombre_build_mercado("Colombia", False, inicio)
    assert nombre == "Osocio BS - Envio Manual - Colombia - 2026-10-09 10:32"
    assert _solo_caracteres_que_acepta_browserstack(nombre)
    # Mismo mercado y misma corrida -> mismo build, para los tres dispositivos.
    assert runner.nombre_build_mercado("Colombia", False, inicio) == nombre
    # Otra corrida (otra hora) u otro mercado -> otro build.
    assert runner.nombre_build_mercado("Colombia", False, datetime(2026, 10, 9, 11, 0)) != nombre
    assert runner.nombre_build_mercado("Brasil", False, inicio) != nombre
    assert runner.nombre_build_mercado("Colombia", True, inicio).startswith("Osocio BS - Automatizado")


def test_dentro_del_build_cada_sesion_dice_su_dispositivo():
    sesion = runner.nombre_sesion("Colombia", "iphone", "data/Lead_information_Formulario_Colombia_Iphone.xlsx")
    assert sesion == "iPhone - Colombia - Lead_information_Formulario_Colombia_Iphone.xlsx"
    assert _solo_caracteres_que_acepta_browserstack(runner.nombre_build("Colombia", "mac", False))


# --- sin capturas en la nube ----------------------------------------------------------

class _DriverFalso:
    name = "safari"


def test_sin_capturas_pero_con_la_senal_de_evidencia(tmp_path):
    from osocio.core.screenshot_manager import ScreenshotManagerNulo
    gestor = ScreenshotManagerNulo(_DriverFalso(), str(tmp_path))
    assert gestor.fname("form", "typage", 1) == ScreenshotManagerNulo.TEXTO_EVIDENCIA
    assert gestor.take_form_screenshot(1, "typage", full_page=True) is None
    assert gestor.take_landing_screenshot(1, "inicial") is None
    assert list(tmp_path.iterdir()) == []


# --- contrato del runner ----------------------------------------------------------------

def test_runner_plataforma_desconocida():
    res = runner.run("Brasil", "windows")
    assert res["error"] and res["ok"] == 0 and res["fail_rows"] == []


def test_runner_sin_excel(tmp_path):
    res = runner.run("Brasil", "mac", excel_path=str(tmp_path / "no_existe.xlsx"))
    assert "Excel no encontrado" in res["error"]


def test_runner_sin_credenciales_no_abre_nada(monkeypatch, tmp_path):
    excel = tmp_path / "leads.xlsx"
    excel.write_bytes(b"x")

    def _sin_credenciales():
        raise FileNotFoundError("No se encontraron credenciales de BrowserStack.")
    monkeypatch.setattr("osocio.providers.browserstack.bs_driver.load_credentials", _sin_credenciales)
    res = runner.run("Brasil", "iphone", excel_path=str(excel))
    assert "credenciales" in res["error"]
    assert set(res) >= {"pais", "ok", "failed", "total", "results_excel", "video_url",
                        "session_id", "error", "fail_rows"}


def test_cada_dispositivo_corre_el_motor_de_la_app():
    from osocio.core.generic_country_base import GenericCountryBase
    from osocio.providers.browserstack.iphone_runner import FormFillerIphone
    assert runner._clase_motor("mac") is GenericCountryBase
    assert runner._clase_motor("android") is GenericCountryBase
    assert runner._clase_motor("iphone") is FormFillerIphone
    assert issubclass(FormFillerIphone, GenericCountryBase)


def test_la_columna_video_se_escribe_en_todas_las_filas(tmp_path):
    from openpyxl import Workbook, load_workbook
    path = tmp_path / "resultados.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.append(["Resultado", "URL"])
    ws.append(["Lead enviado correctamente", "https://a"])
    ws.append(["Error", "https://b"])
    wb.save(path)
    runner.escribir_video(str(path), "https://video", log=lambda *_: None)
    ws = load_workbook(path).active
    col = [c.value for c in ws[1]].index("Video") + 1
    assert [ws.cell(row=r, column=col).value for r in (2, 3)] == ["https://video"] * 2


# --- iPhone: etiqueta visible -> id del mapping ------------------------------------------

@pytest.mark.parametrize("etiqueta,id_esperado", [
    ("Nombre*", "firstname"), ("Nome", "firstname"), ("Sobrenome", "lastname"),
    ("Cédula", "document"), ("CPF", "document"), ("CEP", "cep"),
    ("Tipo de documento", "document-type"), ("Concessionária", "dealer"),
    ("Departamento", "region"), ("Cidade", "city"), ("E-mail", "email"),
    ("Data estimada de compra", "estimated-day"), ("Modelos", "models"),
])
def test_etiqueta_a_id(etiqueta, id_esperado):
    assert id_esperado in ids_para_etiqueta(etiqueta)


@pytest.mark.parametrize("etiqueta", ["Acepto los términos y condiciones", "Username", ""])
def test_etiquetas_que_no_son_campos_del_mapping(etiqueta):
    # "Acepto" contiene "cep": las claves cortas van por palabra entera.
    assert ids_para_etiqueta(etiqueta) == ()


# --- iPhone: las reglas de datos son las de la app ---------------------------------------

@pytest.fixture
def iphone_brasil(monkeypatch):
    from osocio.providers.browserstack.iphone_runner import FormFillerIphone
    form = FormFillerIphone("Brasil", browser="browserstack_iphone", capturas=False)
    form.begin_row_tracking()
    form._tipo_doc_ios = ""
    monkeypatch.setattr(form, "_cargar_ids_dinamicos", lambda: {"lastname": ["Fijo"]})
    return form


def _form_data(form, valores_por_nombre):
    """form_data como lo arma extract_form_data: valor por id del mapping."""
    by_id = {}
    for fc in form.field_mapping:
        valor = valores_por_nombre.get(fc.get("name"))
        if valor is None:
            continue
        ids = fc.get("id") if isinstance(fc.get("id"), list) else [fc.get("id")]
        for i in ids:
            by_id[str(i)] = valor
    return {"__by_id": by_id}


def test_iphone_el_guion_omite_el_campo(iphone_brasil):
    datos = _form_data(iphone_brasil, {"Nombre": "-"})
    assert iphone_brasil._valor_para_etiqueta(datos, "texto", "Nome") is OMITIR


def test_iphone_el_excel_le_gana_al_id_dinamico(iphone_brasil):
    datos = _form_data(iphone_brasil, {"Apellido": "Pérez"})
    assert iphone_brasil._valor_para_etiqueta(datos, "texto", "Sobrenome") == "Pérez"


def test_iphone_celda_vacia_usa_el_id_dinamico(iphone_brasil):
    datos = _form_data(iphone_brasil, {"Apellido": ""})
    assert iphone_brasil._valor_para_etiqueta(datos, "texto", "Sobrenome") == "Fijo"


def test_iphone_select_vacio_va_al_azar(iphone_brasil):
    datos = _form_data(iphone_brasil, {"Modelo": ""})
    assert iphone_brasil._valor_para_etiqueta(datos, "select", "Modelos") is None


def test_iphone_select_con_dato_del_excel(iphone_brasil):
    datos = _form_data(iphone_brasil, {"Modelo": "S10"})
    assert iphone_brasil._valor_para_etiqueta(datos, "select", "Modelos") == "S10"


def test_iphone_cpf_recupera_el_cero_que_excel_se_come(iphone_brasil):
    cpf_col = next(fc["name"] for fc in iphone_brasil.field_mapping
                   if "cpf" in (fc.get("id") if isinstance(fc.get("id"), list) else [fc.get("id")]))
    datos = _form_data(iphone_brasil, {cpf_col: "9446875048"})
    assert iphone_brasil._valor_para_etiqueta(datos, "texto", "CPF") == "09446875048"


def test_iphone_campo_opcional_no_mapeado_se_decide_una_vez(iphone_brasil):
    datos = _form_data(iphone_brasil, {})
    primero = iphone_brasil._valor_para_etiqueta(datos, "texto", "Profesión")
    for _ in range(5):
        assert iphone_brasil._valor_para_etiqueta(datos, "texto", "Profesión") == primero


def test_iphone_campo_obligatorio_no_mapeado_siempre_se_completa(iphone_brasil):
    datos = _form_data(iphone_brasil, {})
    assert iphone_brasil._valor_para_etiqueta(datos, "texto", "Endereço*")


# --- los dispositivos de la nube de un mercado arrancan juntos ---------------------------

def test_la_nube_arranca_junta_aunque_los_excels_esten_en_consecutivo():
    """iPhone, Android y Mac del mismo mercado tienen que estar corriendo a la vez aunque
    "Excels por dispositivo" esté en Consecutivo: en la nube no compiten con la PC."""
    import ast
    import inspect
    import textwrap
    from osocio.interface import envio_leads
    fuente = textwrap.dedent(inspect.getsource(envio_leads))
    arbol = ast.parse(fuente)
    run_market = next(n for n in ast.walk(arbol)
                      if isinstance(n, ast.FunctionDef) and n.name == "_run_market")
    codigo = ast.get_source_segment(fuente, run_market)
    # Las sesiones de la nube se separan y se lanzan en hilos ANTES de mirar el modo.
    assert "bs_dispositivo" in codigo
    assert codigo.index("for t in hilos_nube:") < codigo.index("_parallel = (")
