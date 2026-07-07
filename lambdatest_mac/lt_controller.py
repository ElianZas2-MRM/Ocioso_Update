"""
lt_controller.py
----------------
API pública para ejecutar LambdaTest Mac desde la interfaz Osocio.

Responsabilidades:
- Leer credenciales desde lambdatest_credentials.txt (BASE_DIR o lambdatest_mac/)
- Exponer run() que llama run_lt_batch() con skip_screenshots y stop_event
- Retornar resumen con video_url y paths de resultados
"""

import os
import sys
import json
import threading
from typing import Callable, Optional

_THIS_DIR   = os.path.dirname(os.path.abspath(__file__))   # lambdatest_mac/
if getattr(sys, 'frozen', False):
    _OSOCIO_DIR = os.path.dirname(sys.executable)
else:
    _OSOCIO_DIR = os.path.dirname(_THIS_DIR)                   # raíz del proyecto
_JSON_DIR   = os.path.join(_OSOCIO_DIR, "json")

# Asegurar que lambdatest_mac esté en sys.path para los imports de lt_runner
if _THIS_DIR not in sys.path:
    sys.path.insert(0, _THIS_DIR)




def run(
    pais: str,
    platform: str = "mac",
    log_fn: Callable = print,
    stop_event: Optional[threading.Event] = None,
    build_name: str = "",
    excel_path: Optional[str] = None,
) -> dict:
    """
    Ejecuta LambdaTest para un país y retorna el resumen.

    Retorna:
      {
        "pais": str,
        "ok": int,
        "failed": int,
        "total": int,
        "results_excel": str | None,
        "video_url": str,
        "session_id": str,
        "error": str | None,
      }
    """
    from lt_runner import run_lt_batch, LTRunOptions, load_credentials

    try:
        username, access_key = load_credentials()
    except Exception as e:
        return {
            "pais": pais, "ok": 0, "failed": 0, "total": 0,
            "results_excel": None, "video_url": "",
            "session_id": "", "error": str(e),
        }

    if not excel_path:
        dev_suffix = "Mac" if "mac" in platform.lower() or "iphone" in platform.lower() else "Main"
        excel_name = f"Lead_information_Formulario_{pais}_{dev_suffix}.xlsx"
        excel_path = os.path.join(_OSOCIO_DIR, "data", excel_name)

    if not os.path.exists(excel_path):
        return {
            "pais": pais, "ok": 0, "failed": 0, "total": 0,
            "results_excel": None, "video_url": "",
            "session_id": "", "error": f"Excel no encontrado: {excel_path}",
        }

    brasil_docs = {}
    try:
        with open(os.path.join(_JSON_DIR, "config_global.json"), "r", encoding="utf-8") as _f:
            brasil_docs = json.load(_f).get("lambdatest", {}).get("brasil_docs", {})
    except Exception:
        pass

    opts = LTRunOptions(
        pais=pais,
        excel_path=excel_path,
        build_name=build_name or f"Osocio LT Mac — {pais}",
        platform=platform,
        brasil_docs=brasil_docs,
    )

    summary = run_lt_batch(opts, log=log_fn, stop_event=stop_event)
    return summary
