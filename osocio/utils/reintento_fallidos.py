"""
reintento_fallidos.py
=====================
Qué filas fallaron la última vez, leído del Excel de resultados (no de memoria).

"Reintentar fallidos" usaba sólo lo que la app recordaba de la última corrida de ESTA
sesión: si la app se cerró (o se reabrió para tomar una actualización) el botón decía que
no había nada para reintentar. Con esto se toma el último Excel de resultados del país y
dispositivo, sus filas en rojo, y se ubican en el Excel de entrada por URL + Formulario
(así sirve aunque el último resultado sea de un reintento, con otra numeración).
"""

import glob
import os
import re
from typing import List, Optional, Tuple

from osocio.core.variantes import basenames_de_corrida

ROJO_FALLA = "FFC7CE"   # el color con el que el motor pinta "Resultado" si el lead falló


def ultimo_resultado(carpeta_resultados: str, pais: str, navegador: str,
                     sufijo: str = "") -> Optional[str]:
    """El Excel de resultados más reciente de esa corrida (resultados_<Pais>_<Disp>N.xlsx)."""
    base = basenames_de_corrida(pais, navegador, False, sufijo)[0]
    patron = re.compile(re.escape(base) + r"\d+\.xlsx$", re.IGNORECASE)
    candidatos = [p for p in glob.glob(os.path.join(carpeta_resultados, "**", f"{base}*.xlsx"),
                                       recursive=True)
                  if patron.match(os.path.basename(p)) and not os.path.basename(p).startswith("~$")]
    return max(candidatos, key=os.path.getmtime) if candidatos else None


def _norm(url) -> str:
    return str(url or "").strip().rstrip("/").lower()


def urls_fallidas(resultados_path: str) -> List[Tuple[str, str]]:
    """(URL, Formulario) de cada fila con "Resultado" en rojo."""
    from openpyxl import load_workbook
    wb = load_workbook(resultados_path, read_only=False)
    ws = wb.active
    hdr = [str(c.value or "").strip() for c in ws[1]]
    if "Resultado" not in hdr or "URL" not in hdr:
        return []
    c_res, c_url = hdr.index("Resultado") + 1, hdr.index("URL") + 1
    c_form = hdr.index("Formulario") + 1 if "Formulario" in hdr else None
    out = []
    for fila in range(2, ws.max_row + 1):
        celda = ws.cell(row=fila, column=c_res)
        color = str(getattr(celda.fill.fgColor, "rgb", "") or "")
        if not str(celda.value or "").strip() or not color.upper().endswith(ROJO_FALLA):
            continue
        out.append((str(ws.cell(row=fila, column=c_url).value or ""),
                    str(ws.cell(row=fila, column=c_form).value or "") if c_form else ""))
    return out


def filas_en_entrada(excel_entrada: str, fallidas: List[Tuple[str, str]]) -> List[int]:
    """Filas físicas (2 = primera de datos) del Excel de entrada con esas URL + Formulario.

    Si una misma URL aparece varias veces, cada falla toma una fila distinta, en orden.
    """
    import pandas as pd
    if not fallidas or not os.path.exists(excel_entrada):
        return []
    df = pd.read_excel(excel_entrada, dtype=str).fillna("")
    cols = list(df.columns)
    c_url = cols[0] if "URL" not in cols else "URL"
    c_form = "Formulario" if "Formulario" in cols else (cols[1] if len(cols) > 1 else None)
    libres = [(i + 2, _norm(r[c_url]), _norm(r[c_form]) if c_form else "") for i, r in df.iterrows()]
    filas = []
    for url, form in fallidas:
        u, f = _norm(url), _norm(form)
        hit = next((x for x in libres if x[1] == u and (not f or x[2] == f)), None) \
            or next((x for x in libres if x[1] == u), None)
        if hit is not None:
            libres.remove(hit)
            filas.append(hit[0])
    return sorted(filas)


def fallidas_desde_resultados(carpeta_resultados: str, pais: str, navegador: str,
                              excel_entrada: str, sufijo: str = "") -> Tuple[Optional[str], List[int]]:
    """(Excel de resultados usado, filas del Excel de entrada a reintentar)."""
    res = ultimo_resultado(carpeta_resultados, pais, navegador, sufijo)
    if not res:
        return None, []
    try:
        return res, filas_en_entrada(excel_entrada, urls_fallidas(res))
    except Exception:
        return res, []
