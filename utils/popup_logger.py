"""
popup_logger.py — Logger centralizado para la aplicación.
Escribe mensajes a json/runtime.log con timestamp y nivel (INFO/WARNING/ERROR).
En modo debug (variable FORM_AUTOMATION_VERBOSE=1), también imprime en consola.
"""
import os
import sys
from datetime import datetime


def _get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def log_runtime(message, level="INFO"):
    """Escribe logs de runtime en temporales/runtime.log fuera de la UI."""
    try:
        base_dir = _get_base_dir()
        temporales_dir = os.path.join(base_dir, "temporales")
        os.makedirs(temporales_dir, exist_ok=True)

        log_path = os.path.join(temporales_dir, "runtime.log")
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{ts}] [{level}] {message}\n"

        with open(log_path, "a", encoding="utf-8") as f:
            f.write(line)
    except Exception:
        # Nunca romper la ejecución principal por fallo de logging.
        pass


def popup_log(title, message, level="ERROR"):
    """Muestra popup del sistema y persiste el evento en runtime.log."""
    log_runtime(f"{title}: {message}", level=level)

    # Intentar popup nativo de Windows cuando aplica.
    if os.name == "nt":
        try:
            import ctypes

            icon = 0x10 if level.upper() == "ERROR" else 0x40
            ctypes.windll.user32.MessageBoxW(0, str(message), str(title), 0x00000000 | icon | 0x00001000)
            return
        except Exception as exc:
            log_runtime(f"Fallo MessageBoxW: {exc}", level="WARN")

    # Fallback cross-platform con tkinter.
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)

        if level.upper() == "ERROR":
            messagebox.showerror(title, message, parent=root)
        elif level.upper() == "WARN":
            messagebox.showwarning(title, message, parent=root)
        else:
            messagebox.showinfo(title, message, parent=root)

        root.destroy()
    except Exception as exc:
        log_runtime(f"Fallo popup tkinter: {exc}", level="WARN")
