# -*- coding: utf-8 -*-
"""
driver_update_ui.py — Ventana de progreso de la actualizacion de drivers.

Se usa en dos momentos:

- **Al abrir la app** (`ensure_drivers_with_ui()`): solo se muestra si hay algo que descargar.
  Si los drivers locales ya coinciden con los navegadores instalados, `needs_network()` devuelve
  False y la app abre derecho, sin ventana intermedia y sin tocar la red.
- **Desde el boton "Actualizar drivers"** (`ensure_drivers_with_ui(force=True, parent=root)`):
  ahi la ventana se muestra SIEMPRE y siempre termina con una respuesta explicita, aunque no
  hubiera nada que hacer. El usuario apreto un boton: se merece un "esta todo al dia" y no que
  no pase nada.

La descarga corre en un hilo aparte y se comunica con Tk por una Queue: tocar widgets desde el
hilo de trabajo cuelga Tk, asi que el hilo solo encola eventos y el hilo principal los consume
con `after()`.

Los colores estan repetidos de interface/main_interface.py a proposito: importar ese modulo
arrastra PIL y todo el backend (varios segundos), y en el arranque esta ventana justamente tiene
que aparecer antes de eso. Si cambia la paleta de la app, actualizar tambien estas constantes.
"""
import os
import queue
import threading
import tkinter as tk
from tkinter import ttk

from utils.driver_updater import (
    MISSING,
    NO_BROWSER,
    check_drivers,
    ensure_drivers_ready,
    needs_network,
)

try:
    from utils.paths import ASSET_DIR
except Exception:  # pragma: no cover
    ASSET_DIR = ""

try:
    from utils.popup_logger import log_runtime
except Exception:  # pragma: no cover
    def log_runtime(message, level="INFO"):
        print(f"[{level}] {message}")


# Espejo de la paleta de main_interface.py (ver docstring).
_BG = "#5D3C7A"
_CARD = "#4A2666"
_BORDER = "#7D4E9F"
_TEXT = "#FFFFFF"
_TEXT_SOFT = "#E6D6F2"
_OK = "#82E0AA"
_ERR = "#F1948A"
_ACCENT = "#AED6F1"

_AUTOCLOSE_MS = 900


def _human_mb(num_bytes):
    return f"{num_bytes / (1024 * 1024):.1f} MB"


class _DriverUpdateWindow:
    """Ventana con una fila por driver: nombre, estado y barra de progreso.

    `parent` es la ventana principal de la app cuando esto se abre desde el boton. En ese caso
    hay que usar un Toplevel: crear un segundo tk.Tk() con una UI ya corriendo rompe Tk. En el
    arranque no hay ningun root todavia, asi que la ventana crea el suyo.
    """

    def __init__(self, statuses, parent=None, force=False):
        # Solo se muestran los navegadores instalados: si no hay Firefox en la PC, su fila no
        # aporta nada y ademas tampoco se va a descargar geckodriver.
        self.statuses = [s for s in statuses if s.state != NO_BROWSER]
        self.events = queue.Queue()
        self.cancelled = False
        self.finished = False
        self.force = force
        self.result = statuses
        self.rows = {}
        self.summary = None

        self.owns_root = parent is None
        if self.owns_root:
            self.win = tk.Tk()
        else:
            self.win = tk.Toplevel(parent)
            self.win.transient(parent)
        self.parent = parent

        self.win.title("Osocio - Drivers de navegador")
        self.win.configure(bg=_BG)
        self.win.resizable(False, False)
        self.win.protocol("WM_DELETE_WINDOW", self._skip)

        icon = os.path.join(ASSET_DIR, "icon.ico") if ASSET_DIR else ""
        if icon and os.path.exists(icon):
            try:
                self.win.iconbitmap(icon)
            except Exception:
                pass

        self._build()
        self._center()

    # ── Construccion de la UI ────────────────────────────────────────────────
    def _build(self):
        outer = tk.Frame(self.win, bg=_BG, padx=26, pady=22)
        outer.pack(fill="both", expand=True)

        title = ("Drivers de navegador" if self.force
                 else "Actualizando drivers de navegador")
        subtitle = ("Verificando que cada driver coincida con el navegador instalado."
                    if self.force else
                    "Tu navegador se actualizo y el driver que usa Osocio quedo desfasado.\n"
                    "Se esta descargando la version correcta; esto tarda unos segundos.")

        tk.Label(outer, text=title, bg=_BG, fg=_TEXT,
                 font=("Segoe UI", 14, "bold"), anchor="w").pack(fill="x")

        tk.Label(outer, text=subtitle, bg=_BG, fg=_TEXT_SOFT, font=("Segoe UI", 9),
                 justify="left", anchor="w").pack(fill="x", pady=(6, 16))

        card = tk.Frame(outer, bg=_CARD, highlightbackground=_BORDER,
                        highlightthickness=1, padx=16, pady=14)
        card.pack(fill="x")

        style = ttk.Style(self.win)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(
            "Osocio.Horizontal.TProgressbar",
            troughcolor="#2E1146", background=_ACCENT,
            bordercolor=_CARD, lightcolor=_ACCENT, darkcolor=_ACCENT,
        )

        for status in self.statuses:
            row = tk.Frame(card, bg=_CARD)
            row.pack(fill="x", pady=5)

            tk.Label(row, text=status.driver_name, bg=_CARD, fg=_TEXT,
                     font=("Segoe UI", 9, "bold"), width=18, anchor="w").pack(side="left")

            detail = tk.Label(row, text=status.detail, bg=_CARD, fg=_TEXT_SOFT,
                              font=("Segoe UI", 9), anchor="w")
            detail.pack(side="left", fill="x", expand=True)

            bar = ttk.Progressbar(card, style="Osocio.Horizontal.TProgressbar",
                                  mode="determinate", length=430, maximum=100)
            bar.pack(fill="x", pady=(0, 4))
            self.rows[status.key] = {"detail": detail, "bar": bar}

        self.summary = tk.Label(outer, text="", bg=_BG, fg=_TEXT_SOFT,
                                font=("Segoe UI", 10, "bold"), wraplength=430, justify="left")

        self.skip_button = tk.Button(
            outer, text="Omitir y abrir la app" if not self.force else "Cancelar",
            command=self._skip,
            bg=_CARD, fg=_TEXT_SOFT, activebackground=_BORDER, activeforeground=_TEXT,
            relief="flat", font=("Segoe UI", 9), cursor="hand2", padx=12, pady=5,
            borderwidth=0,
        )
        self.skip_button.pack(pady=(16, 0))

    def _center(self):
        self.win.update_idletasks()
        width = self.win.winfo_width()
        height = self.win.winfo_height()
        x = (self.win.winfo_screenwidth() - width) // 2
        y = (self.win.winfo_screenheight() - height) // 3
        self.win.geometry(f"+{x}+{y}")

    # ── Puente hilo de trabajo -> hilo de Tk ─────────────────────────────────
    def _on_status(self, status):
        self.events.put(("status", status.key, status.detail, status.state, status.error))

    def _on_progress(self, key, done, total):
        self.events.put(("progress", key, done, total, ""))

    def _drain(self):
        try:
            while True:
                kind, key, a, b, c = self.events.get_nowait()
                row = self.rows.get(key)
                if not row:
                    continue
                if kind == "status":
                    color = _ERR if c else (_OK if a.startswith("Actualizado") else _TEXT_SOFT)
                    row["detail"].configure(text=a, fg=color)
                    if c:
                        row["bar"].configure(value=0)
                    elif a.startswith("Actualizado"):
                        row["bar"].configure(value=100)
                elif kind == "progress":
                    if b:
                        row["bar"].configure(mode="determinate", value=a * 100 / b)
                        row["detail"].configure(
                            text=f"Descargando... {_human_mb(a)} de {_human_mb(b)}", fg=_TEXT_SOFT
                        )
                    else:
                        # Sin Content-Length no hay porcentaje posible: se muestran los MB.
                        row["detail"].configure(text=f"Descargando... {_human_mb(a)}", fg=_TEXT_SOFT)
        except queue.Empty:
            pass

        if self.finished:
            self._on_finished()
            return
        self.win.after(80, self._drain)

    def _on_finished(self):
        failed = [s for s in self.result if s.error]
        updated = [s for s in self.result if s.updated]

        if failed:
            # Si algo fallo la ventana NO se cierra sola: el usuario tiene que enterarse de que
            # va a correr con el driver viejo (o sin driver) antes de lanzar una ejecucion.
            detail = "; ".join(f"{s.driver_name}: {s.error}" for s in failed)
            log_runtime(f"Drivers con problemas: {detail}", level="WARN")
            blocking = [s for s in failed if s.state == MISSING]
            text = ("Falta un driver y no se pudo descargar. Revisa la conexion."
                    if blocking else
                    "No se pudo actualizar. Se sigue usando el driver actual.")
            self._show_summary(text, _ERR)
            return

        if not self.force:
            self.win.after(_AUTOCLOSE_MS, self._close)
            return

        # Camino manual: siempre hay una respuesta, aunque no hubiera nada que hacer.
        if updated:
            nombres = ", ".join(s.driver_name for s in updated)
            self._show_summary(f"Listo. Se actualizo: {nombres}", _OK)
        else:
            self._show_summary("Todo al dia: no habia nada que actualizar.", _OK)

    def _show_summary(self, text, color):
        self.summary.configure(text=text, fg=color)
        self.summary.pack(pady=(16, 0), before=self.skip_button)
        self.skip_button.configure(text="Cerrar", fg=_TEXT)
        self.win.update_idletasks()
        self._center()

    def _skip(self):
        self.cancelled = True
        self._close()

    def _close(self):
        try:
            self.win.destroy()
        except Exception:
            pass

    # ── Ejecucion ────────────────────────────────────────────────────────────
    def run(self):
        def worker():
            try:
                self.result = ensure_drivers_ready(
                    on_status=self._on_status,
                    on_progress=self._on_progress,
                    should_cancel=lambda: self.cancelled,
                    force_check=self.force,
                )
            except Exception as exc:  # pragma: no cover - ensure_drivers_ready ya no lanza
                log_runtime(f"Fallo la actualizacion de drivers: {exc}", level="ERROR")
            finally:
                self.finished = True

        threading.Thread(target=worker, daemon=True, name="driver-updater").start()
        self.win.after(80, self._drain)
        if self.owns_root:
            self.win.mainloop()
        else:
            # wait_window corre un loop anidado: la app principal sigue respondiendo mientras
            # esta ventana esta abierta.
            self.parent.wait_window(self.win)
        return self.result


def ensure_drivers_with_ui(force=False, parent=None):
    """Chequea los drivers y muestra la ventana de progreso.

    force=False (arranque): la ventana solo aparece si hay algo que descargar.
    force=True (boton "Actualizar drivers"): la ventana aparece siempre, ignora el cache
    semanal de geckodriver y termina con un mensaje explicito, aunque no hubiera nada que hacer.

    Devuelve la lista de DriverStatus (o None si no se pudo chequear). No lanza nunca: si algo
    sale mal la app tiene que seguir andando.
    """
    try:
        statuses = check_drivers()
    except Exception as exc:
        log_runtime(f"No se pudo chequear el estado de los drivers: {exc}", level="ERROR")
        return None

    if not force and not needs_network(statuses):
        log_runtime("Drivers al dia, no hace falta descargar nada", level="INFO")
        return statuses

    try:
        return _DriverUpdateWindow(statuses, parent=parent, force=force).run()
    except Exception as exc:
        # Si Tk falla (sesion sin escritorio, etc.) igual conviene intentar la actualizacion
        # silenciosa antes de rendirse: sin driver valido la corrida no arranca.
        log_runtime(f"No se pudo mostrar la ventana de drivers: {exc}", level="WARN")
        try:
            return ensure_drivers_ready(force_check=force)
        except Exception:
            return statuses
