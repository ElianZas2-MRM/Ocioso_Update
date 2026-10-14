"""Reintentar fallidos lee las filas en rojo del último Excel de resultados (no de memoria)."""
import os
import time

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill

from osocio.utils.reintento_fallidos import fallidas_desde_resultados, ultimo_resultado

VERDE, ROJO, AMBAR = "C6EFCE", "FFC7CE", "FFC000"


def _resultado(path, filas):
    """filas: (URL, Formulario, color de 'Resultado')."""
    wb = Workbook()
    ws = wb.active
    ws.append(["Resultado", "Motivo", "URL", "Formulario"])
    for url, form, color in filas:
        ws.append(["x", "-", url, form])
        ws.cell(row=ws.max_row, column=1).fill = PatternFill(fill_type="solid", fgColor=color)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)


def _entrada(path, filas):
    pd.DataFrame(filas, columns=["URL", "Formulario", "Nombre"]).to_excel(path, index=False)


def test_toma_las_filas_en_rojo_del_ultimo_resultado(tmp_path):
    entrada = str(tmp_path / "Lead_information_Formulario_Argentina_Mac.xlsx")
    _entrada(entrada, [("https://a.com/1", "https://f.com/1", "Ana"),
                       ("https://a.com/2", "https://f.com/2", "Bea"),
                       ("https://a.com/3", "https://f.com/3", "Ceci")])
    res = str(tmp_path / "resultados" / "t1" / "resultados_Argentina_BS_Mac1.xlsx")
    # Ámbar = se envió por otro form: no se reintenta. Sólo el rojo.
    _resultado(res, [("https://a.com/1", "https://f.com/1", VERDE),
                     ("https://a.com/2", "https://f.com/2", ROJO),
                     ("https://a.com/3", "https://f.com/3", AMBAR)])
    usado, filas = fallidas_desde_resultados(str(tmp_path / "resultados"), "Argentina",
                                             "browserstack_mac", entrada)
    assert os.path.basename(usado) == "resultados_Argentina_BS_Mac1.xlsx"
    assert filas == [3]


def test_un_reintento_anterior_se_ubica_por_url_y_no_por_numero(tmp_path):
    entrada = str(tmp_path / "entrada.xlsx")
    _entrada(entrada, [("https://a.com/1", "https://f.com/1", "Ana"),
                       ("https://a.com/2", "https://f.com/2", "Bea"),
                       ("https://a.com/2", "https://f.com/2", "Bea 2")])
    carpeta = str(tmp_path / "resultados")
    _resultado(os.path.join(carpeta, "t1", "resultados_Argentina_Chrome1.xlsx"),
               [("https://a.com/1", "https://f.com/1", ROJO)])
    time.sleep(0.05)
    # El último resultado es de un reintento: sólo 2 filas, con otra numeración.
    _resultado(os.path.join(carpeta, "t1", "resultados_Argentina_Chrome2.xlsx"),
               [("https://a.com/2", "https://f.com/2/", ROJO),
                ("https://a.com/2", "https://f.com/2", ROJO)])
    usado, filas = fallidas_desde_resultados(carpeta, "Argentina", "chrome", entrada)
    assert os.path.basename(usado) == "resultados_Argentina_Chrome2.xlsx"
    assert filas == [3, 4]   # la URL repetida toma una fila distinta cada vez


def test_no_mezcla_otro_dispositivo_ni_otra_variante(tmp_path):
    carpeta = str(tmp_path / "resultados")
    for nombre in ("resultados_Argentina_BS_Mac1.xlsx", "resultados_Argentina_T3_Chrome1.xlsx",
                   "resultados_Argentina_Chrome10.xlsx"):
        _resultado(os.path.join(carpeta, "t1", nombre), [])
    assert os.path.basename(ultimo_resultado(carpeta, "Argentina", "chrome")) == \
        "resultados_Argentina_Chrome10.xlsx"
    assert ultimo_resultado(carpeta, "Argentina", "browserstack_iphone") is None


def test_si_se_corrigio_la_url_reintenta_con_la_del_excel_de_hoy(tmp_path):
    entrada = str(tmp_path / "entrada.xlsx")
    carpeta = str(tmp_path / "resultados")
    _resultado(os.path.join(carpeta, "t1", "resultados_Argentina_Chrome1.xlsx"),
               [("https://a.com/1", "https://f.com/1", VERDE),
                ("https://a.com/vieja-404", "https://f.com/2", ROJO)])
    # Después de la corrida, la persona corrige la URL de la fila que falló.
    _entrada(entrada, [("https://a.com/1", "https://f.com/1", "Ana"),
                       ("https://a.com/nueva", "https://f.com/2", "Bea")])
    _, filas = fallidas_desde_resultados(carpeta, "Argentina", "chrome", entrada)
    assert filas == [3]   # la fila 3, que hoy dice https://a.com/nueva
    assert pd.read_excel(entrada).iloc[filas[0] - 2]["URL"] == "https://a.com/nueva"
