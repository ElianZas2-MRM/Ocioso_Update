"""
lt_android_controller.py
------------------------
API pública para ejecutar LambdaTest Android desde la interfaz Osocio.
Espejo de lambdatest_mac/lt_controller.py pero para Android (Samsung + Chrome).
"""

import os
import sys
import json
import threading
from typing import Callable, Optional

_THIS_DIR   = os.path.dirname(os.path.abspath(__file__))   # lambdatest_android/
if getattr(sys, 'frozen', False):
    _OSOCIO_DIR = os.path.dirname(sys.executable)
else:
    _OSOCIO_DIR = os.path.dirname(_THIS_DIR)                   # raíz del proyecto
_JSON_DIR   = os.path.join(_OSOCIO_DIR, "json")

if _THIS_DIR not in sys.path:
    sys.path.insert(0, _THIS_DIR)


def run(
    pais: str,
    device_name: str = "Galaxy S24",
    with_screenshots: bool = False,
    log_fn: Callable = print,
    stop_event: Optional[threading.Event] = None,
    build_name: str = "",
    excel_path: Optional[str] = None,
) -> dict:
    """
    Ejecuta LambdaTest Android para un país.

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
    from lt_android_runner import run_lt_android_batch, LTAndroidRunOptions
    from lt_android_runner import load_credentials  # re-export via lambdatest_mac

    try:
        load_credentials()
    except Exception as e:
        return {
            "pais": pais, "ok": 0, "failed": 0, "total": 0,
            "results_excel": None, "video_url": "",
            "session_id": "", "error": str(e),
        }

    if not excel_path:
        excel_name = f"Lead_information_Formulario_{pais}_Android.xlsx"
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

    opts = LTAndroidRunOptions(
        pais=pais,
        excel_path=excel_path,
        build_name=build_name or f"Osocio LT Android — {pais}",
        device_name=device_name,
        with_screenshots=with_screenshots,
        brasil_docs=brasil_docs,
    )

    return run_lt_android_batch(opts, log=log_fn, stop_event=stop_event)
