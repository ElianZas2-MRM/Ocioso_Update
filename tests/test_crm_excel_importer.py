"""Tests de utils.crm_excel_importer.CrmValidacionesImporter.

No usan el excel real del usuario (dato externo sensible, nunca se versiona): cada test
arma un workbook sintético con openpyxl en un archivo temporal.
"""
import datetime
import json
import os

import openpyxl
import pytest

from utils.crm_excel_importer import CrmValidacionesImporter, construir_reporte, guardar_reporte


def _crear_excel_sintetico(tmp_path):
    ruta = os.path.join(tmp_path, "excel_sintetico.xlsx")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Hoja no-país: debe ser ignorada por completo.
    hoja_reportes = wb.create_sheet("GMSA | Reportes")
    hoja_reportes.append(["DASH", "https://example.com", "algo"])

    # Argentina: usa "MENSAJE DE ERROR2" (sin espacio) como variante de header.
    ar = wb.create_sheet("Argentina")
    ar.append(["CAMPO", "TIPO", "OBLIGATORIO", "VALIDACIONES", "MENSAJE DE ERROR",
               "MENSAJE DE ERROR2", "ÚLTIMO AJUSTE"])
    # Ya existe en el JSON con regex_full armado -> no se debe tocar (va a "ok").
    ar.append(["NOMBRE", "Rellenable", True, "De 2 a 50 caracteres.",
               "Ingresá tu nombre.", "Ingresá mínimo 2 letras.", "19 de Junio"])
    # Existe en el JSON pero regex_full vacío y obligatorio=True -> "incompletos".
    ar.append(["APELLIDO", "Rellenable", True, "De 2 a 50 caracteres.",
               "Ingresá tu apellido.", None, None])
    # No existe en el JSON -> "faltantes", se debe crear en el merge.
    ar.append(["DNI", "Rellenable", True, "De 7 a 9 caracteres numéricos.",
               "Ingresá tu DNI.", "Ingresá mínimo 7 dígitos.", None])

    # Chile: usa "MENSAJE DE ERROR 2" (con espacio) como variante de header, y
    # "DESCRIPCIÓN" en vez de "VALIDACIONES" para la columna de prosa.
    cl = wb.create_sheet("Chile")
    cl.append(["CAMPO", "TIPO", "OBLIGATORIO", "DESCRIPCIÓN", "MENSAJE DE ERROR",
               "MENSAJE DE ERROR 2", "ÚLTIMO AJUSTE"])
    # ÚLTIMO AJUSTE con datetime real de Excel (varias hojas del CRM real lo traen así,
    # no como texto) — tiene que quedar como string parseado, nunca crudo.
    cl.append(["RUT", "Rellenable", True, "Posee un algoritmo de validación.",
               "Ingresá tu RUT.", None, datetime.datetime(2026, 6, 30, 0, 0)])

    wb.save(ruta)
    return ruta


def _reglas_json_argentina():
    return {
        "fields": {
            "NOMBRE": {
                "regex_full": "^[a-zA-Z ]{2,50}$",
                "descripcion": "NOMBRE",
            },
            "APELLIDO": {
                "regex_full": "",
                "descripcion": "APELLIDO",
            },
        }
    }


@pytest.fixture
def excel_sintetico(tmp_path):
    return _crear_excel_sintetico(str(tmp_path))


class TestDisponible:
    def test_archivo_inexistente_devuelve_false_con_mensaje_claro(self, tmp_path):
        importer = CrmValidacionesImporter(ruta_excel=str(tmp_path / "no_existe.xlsx"))
        disponible, mensaje = importer.disponible()
        assert disponible is False
        assert "no_existe.xlsx" in mensaje

    def test_archivo_existente_devuelve_true_con_ruta(self, excel_sintetico):
        importer = CrmValidacionesImporter(ruta_excel=excel_sintetico)
        disponible, ruta = importer.disponible()
        assert disponible is True
        assert ruta == excel_sintetico


class TestParsear:
    def test_ignora_hojas_no_pais_conocidas_sin_marcarlas_como_no_reconocidas(self, excel_sintetico):
        importer = CrmValidacionesImporter(ruta_excel=excel_sintetico)
        campos_por_pais = importer.parsear()
        assert "GMSA | Reportes" not in campos_por_pais
        # Es una hoja conocida-y-esperada (NON_COUNTRY_SHEETS), no una desconocida.
        assert "GMSA | Reportes" not in importer.hojas_no_reconocidas

    def test_hoja_con_nombre_de_pais_no_soportado_va_a_no_reconocidas(self, tmp_path):
        ruta = os.path.join(str(tmp_path), "excel_con_hoja_rara.xlsx")
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        rara = wb.create_sheet("Venezuela")
        rara.append(["CAMPO", "TIPO", "OBLIGATORIO"])
        rara.append(["NOMBRE", "Rellenable", True])
        wb.save(ruta)

        importer = CrmValidacionesImporter(ruta_excel=ruta)
        campos_por_pais = importer.parsear()

        assert "Venezuela" not in campos_por_pais
        assert "Venezuela" in importer.hojas_no_reconocidas

    def test_parsea_ambas_hojas_de_pais(self, excel_sintetico):
        importer = CrmValidacionesImporter(ruta_excel=excel_sintetico)
        campos_por_pais = importer.parsear()
        assert set(campos_por_pais.keys()) == {"Argentina", "Chile"}
        assert len(campos_por_pais["Argentina"]) == 3
        assert len(campos_por_pais["Chile"]) == 1

    def test_detecta_columna_mensaje_error_con_headers_no_uniformes(self, excel_sintetico):
        importer = CrmValidacionesImporter(ruta_excel=excel_sintetico)
        campos_por_pais = importer.parsear()

        nombre = next(c for c in campos_por_pais["Argentina"] if c["nombre_campo"] == "NOMBRE")
        assert nombre["mensajes_error"] == ["Ingresá tu nombre.", "Ingresá mínimo 2 letras."]

        rut = next(c for c in campos_por_pais["Chile"] if c["nombre_campo"] == "RUT")
        assert rut["mensajes_error"] == ["Ingresá tu RUT."]

    def test_ultimo_ajuste_datetime_se_convierte_a_string(self, excel_sintetico):
        """El excel real del CRM trae ÚLTIMO AJUSTE como datetime nativo de Excel en
        varias hojas (Chile, Colombia, Ecuador, Paraguay, Peru, Uruguay, Bolivia) — si se
        guarda crudo, json.dump revienta más adelante en construir_reporte/guardar_reporte
        ("Object of type datetime is not JSON serializable")."""
        importer = CrmValidacionesImporter(ruta_excel=excel_sintetico)
        campos_por_pais = importer.parsear()

        rut = next(c for c in campos_por_pais["Chile"] if c["nombre_campo"] == "RUT")
        assert isinstance(rut["ultimo_ajuste"], str)
        assert rut["ultimo_ajuste"] == "2026-06-30"

    def test_cachea_el_parseo_no_relee_el_archivo(self, excel_sintetico):
        importer = CrmValidacionesImporter(ruta_excel=excel_sintetico)
        primera = importer.parsear()
        os.remove(excel_sintetico)
        segunda = importer.parsear()
        assert segunda is primera


class TestComparar:
    def test_campo_con_regex_ya_armada_va_a_ok_y_no_a_incompletos(self, excel_sintetico):
        importer = CrmValidacionesImporter(ruta_excel=excel_sintetico)
        comparacion = importer.comparar("Argentina", _reglas_json_argentina())

        nombres_ok = {c["nombre_campo"] for c in comparacion["ok"]}
        assert "NOMBRE" in nombres_ok
        assert "NOMBRE" not in {c["nombre_campo"] for c in comparacion["incompletos"]}
        assert "NOMBRE" not in {c["nombre_campo"] for c in comparacion["faltantes"]}

    def test_campo_existente_sin_regex_y_obligatorio_va_a_incompletos(self, excel_sintetico):
        importer = CrmValidacionesImporter(ruta_excel=excel_sintetico)
        comparacion = importer.comparar("Argentina", _reglas_json_argentina())

        nombres_incompletos = {c["nombre_campo"] for c in comparacion["incompletos"]}
        assert "APELLIDO" in nombres_incompletos

    def test_campo_inexistente_va_a_faltantes(self, excel_sintetico):
        importer = CrmValidacionesImporter(ruta_excel=excel_sintetico)
        comparacion = importer.comparar("Argentina", _reglas_json_argentina())

        nombres_faltantes = {c["nombre_campo"] for c in comparacion["faltantes"]}
        assert "DNI" in nombres_faltantes


class TestAplicarMerge:
    def test_nunca_pisa_una_entrada_existente(self, excel_sintetico):
        """Invariante más importante del plan: ni 'ok' ni 'incompletos' se tocan."""
        importer = CrmValidacionesImporter(ruta_excel=excel_sintetico)
        reglas = _reglas_json_argentina()
        nombre_original = dict(reglas["fields"]["NOMBRE"])
        apellido_original = dict(reglas["fields"]["APELLIDO"])

        comparacion = importer.comparar("Argentina", reglas)
        actualizado, _ = importer.aplicar_merge("Argentina", comparacion, reglas)

        assert actualizado["fields"]["NOMBRE"] == nombre_original
        assert actualizado["fields"]["APELLIDO"] == apellido_original
        # La copia de entrada no se muta tampoco.
        assert reglas["fields"]["NOMBRE"] == nombre_original

    def test_agrega_solo_los_campos_faltantes(self, excel_sintetico):
        importer = CrmValidacionesImporter(ruta_excel=excel_sintetico)
        reglas = _reglas_json_argentina()
        comparacion = importer.comparar("Argentina", reglas)

        actualizado, cantidad = importer.aplicar_merge("Argentina", comparacion, reglas)

        assert cantidad == 1
        assert "DNI" in actualizado["fields"]
        nuevo = actualizado["fields"]["DNI"]
        assert nuevo["pendiente_regex"] is True
        assert nuevo["origen"] == "crm_excel_import"
        assert nuevo["regex_full"] == ""
        assert nuevo["obligatorio_excel"] is True
        assert nuevo["mensajes_error_excel"] == ["Ingresá tu DNI.", "Ingresá mínimo 7 dígitos."]

    def test_no_agrega_nada_si_no_hay_faltantes(self, excel_sintetico):
        importer = CrmValidacionesImporter(ruta_excel=excel_sintetico)
        reglas_chile = {"fields": {"RUT": {"regex_full": "^[0-9kK.-]+$"}}}
        comparacion = importer.comparar("Chile", reglas_chile)

        actualizado, cantidad = importer.aplicar_merge("Chile", comparacion, reglas_chile)

        assert cantidad == 0
        assert actualizado["fields"]["RUT"]["regex_full"] == "^[0-9kK.-]+$"


class TestReporte:
    def test_construir_reporte_tiene_el_shape_documentado(self, excel_sintetico):
        importer = CrmValidacionesImporter(ruta_excel=excel_sintetico)
        comparacion = importer.comparar("Argentina", _reglas_json_argentina())

        reporte = construir_reporte("Argentina", "Argentina", comparacion)

        assert reporte["pais"] == "Argentina"
        assert reporte["hoja_excel"] == "Argentina"
        assert "fecha_importacion" in reporte
        assert reporte["resumen"] == {"faltantes": 1, "incompletos": 1, "ok": 1}
        assert reporte["faltantes"] == comparacion["faltantes"]
        assert reporte["incompletos"] == comparacion["incompletos"]
        assert reporte["ok"] == comparacion["ok"]

    def test_guardar_reporte_no_revienta_con_ultimo_ajuste_datetime(self, excel_sintetico, tmp_path, monkeypatch):
        json_dir = str(tmp_path / "json")
        monkeypatch.setattr("utils.crm_excel_importer.JSON_DIR", json_dir)

        importer = CrmValidacionesImporter(ruta_excel=excel_sintetico)
        comparacion = importer.comparar("Chile", {"fields": {}})
        reporte = construir_reporte("Chile", "Chile", comparacion)

        ruta_guardada = guardar_reporte("Chile", reporte)  # no debe lanzar TypeError

        with open(ruta_guardada, "r", encoding="utf-8") as fh:
            contenido = json.load(fh)
        assert contenido["faltantes"][0]["ultimo_ajuste"] == "2026-06-30"

    def test_guardar_reporte_escribe_json_dir_con_nombre_por_pais(self, excel_sintetico, tmp_path, monkeypatch):
        json_dir = str(tmp_path / "json")
        monkeypatch.setattr("utils.crm_excel_importer.JSON_DIR", json_dir)

        importer = CrmValidacionesImporter(ruta_excel=excel_sintetico)
        comparacion = importer.comparar("Argentina", _reglas_json_argentina())
        reporte = construir_reporte("Argentina", "Argentina", comparacion)

        ruta_guardada = guardar_reporte("Argentina", reporte)

        assert ruta_guardada == os.path.join(json_dir, "crm_import_report_argentina.json")
        with open(ruta_guardada, "r", encoding="utf-8") as fh:
            contenido = json.load(fh)
        assert contenido["pais"] == "Argentina"
        assert contenido["resumen"]["faltantes"] == 1
