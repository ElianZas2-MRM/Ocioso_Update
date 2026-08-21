"""
popup_logger.py — Logger centralizado para la aplicación.
Escribe mensajes a temporales/runtime.log con timestamp y nivel (INFO/WARNING/ERROR).
En modo debug (variable FORM_AUTOMATION_VERBOSE=1), también imprime en consola.

Sobre los popups: `popup_log` es una NOTIFICACIÓN, no una pregunta — ningún llamador lee su
resultado ni bifurca según lo que conteste el usuario. Por eso nunca tiene que frenar a quien
lo llama, y menos todavía cuando no hay nadie para cerrarlo. Ver `popup_log`.
"""
import os
import sys
import threading
from datetime import datetime

# Modo desatendido: no hay nadie mirando la pantalla, así que un popup no se cierra nunca.
# Lo fija run.py al arrancar con --autonomous; el autodetect de abajo es la red de seguridad
# para procesos hijos que no pasan por ahí.
_unattended_override = None


def set_unattended(value=True):
    """Marca la ejecución como desatendida (sin nadie que pueda cerrar un popup)."""
    global _unattended_override
    _unattended_override = bool(value)


def is_unattended():
    """True si esta ejecución no tiene a nadie delante.

    Se evalúa en cada llamada (no al importar) para que valga tanto el `set_unattended`
    explícito de run.py como la variable de entorno, sin depender del orden de imports.
    """
    if _unattended_override is not None:
        return _unattended_override
    if os.environ.get("OSOCIO_UNATTENDED", "").strip().lower() in ("1", "true", "yes"):
        return True
    # El Programador de tareas lanza `run.py --autonomous [--once]`, y autonomous_runner
    # respawnea con el mismo flag: alcanza con mirar la línea de comandos.
    return "--autonomous" in sys.argv


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


def _show_windows_popup(title, message, level):
    """MessageBoxW pelado. Bloquea al hilo que lo llama hasta que alguien lo cierra."""
    import ctypes

    icon = 0x10 if level.upper() == "ERROR" else 0x40
    # MB_SYSTEMMODAL (0x1000): queda arriba de todo para que no se pierda detrás de la app.
    ctypes.windll.user32.MessageBoxW(0, str(message), str(title), 0x00000000 | icon | 0x00001000)


def popup_log(title, message, level="ERROR"):
    """Persiste el evento en runtime.log y, si hay alguien mirando, lo muestra en un popup.

    Dos reglas, las dos por el mismo motivo — el popup avisa, no pregunta:

    1. En modo desatendido NO se muestra nada. `MessageBoxW` es modal: en una corrida
       programada de madrugada nadie lo cierra, y la corrida quedaba colgada para siempre en
       vez de fallar y liberar el lock del scheduler. Queda solo en runtime.log, que es donde
       se mira al otro día.
    2. Si lo llama un hilo de trabajo (una corrida en curso), el popup se muestra sin frenarlo.
       Antes, un error de Excel a mitad de una corrida congelaba esa corrida hasta que el
       usuario volviera a la máquina. Desde el hilo principal se mantiene bloqueante: ahí el
       llamador suele estar por terminar el proceso y, si no esperáramos, el popup se cerraría
       solo antes de que llegue a leerse.
    """
    log_runtime(f"{title}: {message}", level=level)

    if is_unattended():
        log_runtime(f"Popup omitido (ejecución desatendida): {title}", level="INFO")
        return

    if os.name == "nt":
        try:
            if threading.current_thread() is threading.main_thread():
                _show_windows_popup(title, message, level)
            else:
                threading.Thread(
                    target=_show_windows_popup,
                    args=(title, message, level),
                    daemon=True,
                    name="popup-log",
                ).start()
            return
        except Exception as exc:
            log_runtime(f"Fallo MessageBoxW: {exc}", level="WARN")

    # Fallback cross-platform con tkinter. Solo desde el hilo principal: crear un segundo
    # tk.Tk() desde un hilo de trabajo, con la UI de la app ya corriendo, la tira abajo.
    if threading.current_thread() is not threading.main_thread():
        log_runtime(f"Popup omitido (hilo secundario sin MessageBoxW): {title}", level="WARN")
        return

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
