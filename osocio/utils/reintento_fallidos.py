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
    return [(u, f) for _, u, f in _fallidas_con_fila(resultados_path)[0]]


def _fallidas_con_fila(resultados_path: str):
    """([(fila, URL, Formulario) en rojo], cantidad de filas de datos del resultado)."""
    from openpyxl import load_workbook
    wb = load_workbook(resultados_path, read_only=False)
    ws = wb.active
    hdr = [str(c.value or "").strip() for c in ws[1]]
    if "Resultado" not in hdr or "URL" not in hdr:
        return [], 0
    c_res, c_url = hdr.index("Resultado") + 1, hdr.index("URL") + 1
    c_form = hdr.index("Formulario") + 1 if "Formulario" in hdr else None
    out, filas_datos = [], 0
    for fila in range(2, ws.max_row + 1):
        if any(ws.cell(row=fila, column=c).value not in (None, "") for c in range(1, len(hdr) + 1)):
            filas_datos = fila - 1
        celda = ws.cell(row=fila, column=c_res)
        color = str(getattr(celda.fill.fgColor, "rgb", "") or "")
        if not str(celda.value or "").strip() or not color.upper().endswith(ROJO_FALLA):
            continue
        out.append((fila, str(ws.cell(row=fila, column=c_url).value or ""),
                    str(ws.cell(row=fila, column=c_form).value or "") if c_form else ""))
    return out, filas_datos


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
                              excel_entrada: str, sufijo: str = "",
                              log=None) -> Tuple[Optional[str], List[int]]:
    """(Excel de resultados usado, filas del Excel de entrada a reintentar).

    Se reintenta con lo que dice el Excel de entrada AHORA (si se corrigió la URL de una
    fila que falló, va la corregida):
      - Resultado de una corrida completa (mismas filas que la entrada): por número de
        fila, como el reintento en memoria. La fila 3 que falló es la fila 3 de hoy.
      - Resultado de un reintento anterior (otra numeración): por URL + Formulario. Una
        fila que ya no se encuentra (se cambió la URL) se avisa en el log.
    """
    res = ultimo_resultado(carpeta_resultados, pais, navegador, sufijo)
    if not res or not os.path.exists(excel_entrada):
        return res, []
    try:
        import pandas as pd
        fallidas, filas_res = _fallidas_con_fila(res)
        filas_entrada = len(pd.read_excel(excel_entrada, dtype=str))
        if filas_res == filas_entrada:
            return res, sorted(f for f, _, _ in fallidas)
        filas = filas_en_entrada(excel_entrada, [(u, f) for _, u, f in fallidas])
        if len(filas) < len(fallidas) and log is not None:
            log(f"[WARN] Reintentar fallidos: {len(fallidas) - len(filas)} fila(s) de "
                f"{os.path.basename(res)} ya no están en el Excel de entrada (¿se cambió la URL?). "
                "Corré el envío normal para esas filas.")
        return res, filas
    except Exception:
        return res, []
