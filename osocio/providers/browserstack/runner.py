"""
runner.py
=========
Punto de entrada de BrowserStack para la app: lo llaman Envío de Leads, el CLI
(run.py --run-browserstack) y el runner autónomo. Corre un Excel de leads de un mercado
en un dispositivo de la nube y devuelve el resumen.

- Android y Mac: el motor de los navegadores (GenericCountryBase / BaseFormFiller) con el
  driver de BrowserStack. Heredan todas las reglas de la app (mapping, IDs dinámicos,
  '-' para omitir, documentos de Brasil, variantes CDC/T3, carpetas t1/t3, reintentos,
  columnas de resultado) sin duplicar nada.
- iPhone: el mismo motor (iphone_runner.FormFillerIphone), con el llenado por la capa
  nativa porque Safari en iPhone no deja entrar al iframe del form.

Mismo contrato de resumen que tenía LambdaTest, así la UI no cambió de forma:
  pais, ok, failed, total, results_excel, video_url, session_id, error, fail_rows
"""

import os
import threading
from datetime import datetime
from typing import Callable, Dict, Optional

from osocio.paths import DATA_DIR
from osocio.core.variantes import ruta_excel_leads

PLATAFORMAS = ("iphone", "android", "mac")

# Sufijo del Excel de entrada por dispositivo (los de Mac/Android son los que ya
# generaba la app para LambdaTest, así siguen sirviendo).
SUFIJO_EXCEL = {"iphone": "Iphone", "android": "Android", "mac": "Mac"}

COLUMNA_VIDEO = "Video"


def excel_por_defecto(pais: str, plataforma: str, variante: str = "") -> str:
    return ruta_excel_leads(DATA_DIR, pais, SUFIJO_EXCEL[plataforma], variante)


_NOMBRE_DISP = {"iphone": "iPhone", "android": "Android", "mac": "Mac"}


_DIAS = ("Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado", "Domingo")


def nombre_build_del_dia(programada: bool, dia: Optional[datetime] = None) -> str:
    """Un build por día: todos los mercados y dispositivos de ese día van juntos.

    Ej.: "Osocio BS - Envio Manual - Miercoles 2026-10-14". BrowserStack junta en el
    mismo build las sesiones con el mismo nombre aunque lleguen en corridas distintas
    (probado con horas de diferencia), y cada sesión se distingue por su nombre
    (dispositivo - mercado - Excel). Manual y automatizado van en builds separados.
    Sin acentos, guiones largos ni paréntesis: BrowserStack los descarta del nombre.
    """
    dia = dia or datetime.now()
    modo = "Automatizado" if programada else "Envio Manual"
    return f"Osocio BS - {modo} - {_DIAS[dia.weekday()]} {dia:%Y-%m-%d}"


def nombre_build(pais: str, plataforma: str, programada: bool) -> str:
    """Build de un dispositivo suelto (CLI sin --build-name)."""
    modo = "Automatizado" if programada else "Envio Manual"
    return f"Osocio BS {_NOMBRE_DISP[plataforma]} - {modo} - {pais}"


def nombre_sesion(pais: str, plataforma: str, excel_path: str) -> str:
    """Dentro del build del mercado, cada sesión se distingue por el dispositivo."""
    return f"{_NOMBRE_DISP[plataforma]} - {pais} - {os.path.basename(excel_path)}"


def _ruta_corta(url: str) -> str:
    """La parte de la URL que distingue un lead de otro (sin dominio ni parámetros)."""
    from urllib.parse import urlparse
    ruta = urlparse(url or "").path.strip("/")
    return ruta[-60:] if ruta else (url or "")[:60]


def texto_marca_lead(evento: str, numero: int, total: int, fila: int, url: str = "",
                     ok: bool = True, motivo: str = "", enviar: bool = True) -> str:
    """Texto de la marca en el video. Ej.: "Lead 3/10 · fila 4 · FALLÓ: sin TYP"."""
    base = f"Lead {numero}/{total} · fila {fila}"
    if evento == "inicio":
        return f"{base} · empieza · {_ruta_corta(url)}"
    if not enviar:
        # Sin envío el motor no llega a la TYP: no es un fallo del lead.
        return f"{base} · revisado sin enviar"
    return f"{base} · enviado OK" if ok else f"{base} · FALLÓ: {motivo or 'ver Excel'}"


class MarcasDelVideo:
    """on_lead del motor: deja una marca en el video de BrowserStack al empezar y al
    terminar cada lead, así en un video largo se va directo al lead que falló."""

    def __init__(self, fabrica, enviar: bool = True):
        self.fabrica = fabrica
        self.enviar = enviar

    def __call__(self, evento, fila=0, numero=0, total=0, url="", ok=True, motivo=""):
        sesion = getattr(self.fabrica, "sesion", None)
        driver = getattr(sesion, "driver", None)
        if driver is None:
            return
        from osocio.providers.browserstack.bs_driver import annotate_script
        texto = texto_marca_lead(evento, numero, total, fila, url, ok, motivo, self.enviar)
        error = evento == "fin" and self.enviar and not ok
        driver.execute_script(annotate_script(texto, "error" if error else "info"))


def motivo_de_sesion(ok: int, fallidos: int, fail_rows=None, error: str = "") -> str:
    """Lo que se ve en el dashboard junto al passed/failed, sin abrir la sesión.

    Ej.: "2 de 10 leads fallaron: sin TYP (filas 4, 7)"."""
    total = ok + fallidos
    if error:
        return f"Error de la corrida: {error}"
    if not total:
        return "No se procesó ningún lead"
    if not fallidos:
        return f"{ok} de {total} leads enviados OK"
    por_motivo: Dict[str, list] = {}
    for f in fail_rows or []:
        por_motivo.setdefault(f.get("reason") or "sin detalle", []).append(str(f.get("row", "?")))
    detalle = "; ".join(f"{m} (fila{'s' if len(filas) > 1 else ''} {', '.join(filas)})"
                        for m, filas in por_motivo.items())
    return f"{fallidos} de {total} leads fallaron" + (f": {detalle}" if detalle else "")


def _resumen_vacio(pais: str, error: str = "") -> Dict:
    return {"pais": pais, "ok": 0, "failed": 0, "total": 0, "results_excel": None,
            "video_url": "", "session_id": "", "error": error or None, "fail_rows": [],
            "warn": 0, "warn_rows": []}


def escribir_video(results_path: str, video_url: str, log: Callable = print) -> None:
    """Agrega/llena la columna Video en todas las filas y reordena el Excel."""
    if not (results_path and os.path.exists(results_path)):
        return
    try:
        from openpyxl import load_workbook
        wb = load_workbook(results_path)
        ws = wb.active
        headers = [c.value for c in ws[1]]
        if COLUMNA_VIDEO in headers:
            col = headers.index(COLUMNA_VIDEO) + 1
        else:
            col = len(headers) + 1
            ws.cell(row=1, column=col).value = COLUMNA_VIDEO
        for fila in range(2, ws.max_row + 1):
            ws.cell(row=fila, column=col).value = video_url or "-"
        wb.save(results_path)
        from osocio.utils.excel_layout import reordenar_archivo
        reordenar_archivo(results_path, log=log)
    except Exception as e:
        log(f"  ⚠ No se pudo escribir la columna Video: {e}")


def run(pais: str, plataforma: str, excel_path: Optional[str] = None, variante: str = "",
        programada: bool = False, build_name: str = "", log_fn: Callable = print,
        stop_event: Optional[threading.Event] = None, enviar: bool = True) -> Dict:
    """Corre el Excel de un mercado en un dispositivo de BrowserStack.

    enviar=False llena todo y no toca Enviar (modo chequeo_campos del motor): para probar
    sin cargar leads en la base.
    """
    if plataforma not in PLATAFORMAS:
        return _resumen_vacio(pais, f"Plataforma de BrowserStack desconocida: {plataforma}")
    excel_path = excel_path or excel_por_defecto(pais, plataforma, variante)
    if not os.path.exists(excel_path):
        return _resumen_vacio(pais, f"Excel no encontrado: {excel_path}")
    build_name = build_name or nombre_build(pais, plataforma, programada)
    try:
        from osocio.providers.browserstack.bs_driver import load_credentials
        load_credentials()
    except Exception as e:
        return _resumen_vacio(pais, str(e))

    return _run_motor(pais, plataforma, excel_path, variante, programada, build_name, log_fn,
                      enviar)


def _clase_motor(plataforma):
    """El motor de la app para cada dispositivo: el de los navegadores tal cual, salvo el
    iPhone, que es el mismo motor con el llenado por la capa nativa."""
    if plataforma == "iphone":
        from osocio.providers.browserstack.iphone_runner import FormFillerIphone
        return FormFillerIphone
    from osocio.core.generic_country_base import GenericCountryBase
    return GenericCountryBase


def _run_motor(pais, plataforma, excel_path, variante, programada, build_name, log_fn, enviar=True):
    from osocio.providers.browserstack.fabrica import FabricaNavegadorBrowserStack

    resumen = _resumen_vacio(pais)
    fab = FabricaNavegadorBrowserStack(plataforma, build_name,
                                       nombre_sesion(pais, plataforma, excel_path), log=log_fn)
    form = _clase_motor(plataforma)(pais, browser=f"browserstack_{plataforma}", browser_factory=fab,
                              is_scheduled=programada, excel_suffix=variante, capturas=False)
    form.EXCEL_PATH = excel_path
    form.on_lead = MarcasDelVideo(fab, enviar)
    if not enviar:
        form.config["chequeo_campos"] = True
    try:
        form.run()
    except Exception as e:
        resumen["error"] = str(e)
        log_fn(f"✗ Error en la corrida BrowserStack {plataforma}: {e}")
    finally:
        summ = getattr(form, "run_summary", None) or {}
        resumen.update(ok=int(summ.get("ok", 0)), failed=int(summ.get("fail", 0)),
                       total=int(summ.get("total", 0)), fail_rows=list(summ.get("fail_rows") or []),
                       warn=int(summ.get("warn", 0)), warn_rows=list(summ.get("warn_rows") or []),
                       results_excel=getattr(form, "RESULTADOS_PATH", None))
        if fab.sesion is not None:
            if enviar:
                # Los avisos (URL / form / TYP) no marcan la sesión como fallida: no son del envío.
                todo_ok = (resumen["failed"] == 0 and resumen["ok"] + resumen["warn"] > 0
                           and not resumen["error"])
                motivo = motivo_de_sesion(resumen["ok"], resumen["failed"], resumen["fail_rows"],
                                          resumen["error"] or "")
            else:
                # Sin envío el motor anota "Verificación visual", no "Lead enviado": eso no
                # es un fallo de la sesión.
                todo_ok = not resumen["error"]
                motivo = "Prueba sin envío (chequeo de campos)"
            info = fab.sesion.cerrar(todo_ok, motivo)
            resumen["session_id"] = fab.sesion.session_id
            resumen["video_url"] = info.get("video_url", "")
            if resumen["video_url"]:
                log_fn(f"  Video: {resumen['video_url']}")
        escribir_video(resumen["results_excel"], resumen["video_url"], log=log_fn)
    return resumen
