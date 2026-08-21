# -*- coding: utf-8 -*-
"""
driver_update_ui.py — Ventana de progreso de la actualizacion de drivers.

Se muestra al abrir la app SOLO cuando hay algo que descargar. Si los drivers locales ya
coinciden con los navegadores instalados, `utils.driver_updater.needs_network()` devuelve False
y la app abre derecho, sin ventana intermedia y sin tocar la red.

La descarga corre en un hilo aparte y se comunica con Tk por una Queue: tocar widgets desde el
hilo de trabajo cuelga Tk, asi que el hilo solo encola eventos y el hilo principal los consume
con `after()`.

Los colores estan repetidos de interface/main_interface.py a proposito: importar ese modulo
arrastra PIL y todo el backend (varios segundos), y esta ventana justamente tiene que aparecer
antes de eso. Si cambia la paleta de la app, actualizar tambien estas constantes.
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
    """Ventana con una fila por driver: nombre, estado y barra de progreso."""

    def __init__(self, statuses):
        # Solo se muestran los navegadores instalados: si no hay Firefox en la PC, su fila no
        # aporta nada y ademas tampoco se va a descargar geckodriver.
        self.statuses = [s for s in statuses if s.state != NO_BROWSER]
        self.events = queue.Queue()
        self.cancelled = False
        self.finished = False
        self.result = statuses
        self.rows = {}

        self.root = tk.Tk()
        self.root.title("Osocio - Preparando drivers")
        self.root.configure(bg=_BG)
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", self._skip)

        icon = os.path.join(ASSET_DIR, "icon.ico") if ASSET_DIR else ""
        if icon and os.path.exists(icon):
            try:
                self.root.iconbitmap(icon)
            except Exception:
                pass

        self._build()
        self._center()

    # ── Construccion de la UI ────────────────────────────────────────────────
    def _build(self):
        outer = tk.Frame(self.root, bg=_BG, padx=26, pady=22)
        outer.pack(fill="both", expand=True)

        tk.Label(
            outer, text="Actualizando drivers de navegador",
            bg=_BG, fg=_TEXT, font=("Segoe UI", 14, "bold"), anchor="w",
        ).pack(fill="x")

        tk.Label(
            outer,
            text=("Tu navegador se actualizo y el driver que usa Osocio quedo desfasado.\n"
                  "Se esta descargando la version correcta; esto tarda unos segundos."),
            bg=_BG, fg=_TEXT_SOFT, font=("Segoe UI", 9), justify="left", anchor="w",
        ).pack(fill="x", pady=(6, 16))

        card = tk.Frame(outer, bg=_CARD, highlightbackground=_BORDER,
                        highlightthickness=1, padx=16, pady=14)
        card.pack(fill="x")

        style = ttk.Style(self.root)
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

        self.skip_button = tk.Button(
            outer, text="Omitir y abrir la app", command=self._skip,
            bg=_CARD, fg=_TEXT_SOFT, activebackground=_BORDER, activeforeground=_TEXT,
            relief="flat", font=("Segoe UI", 9), cursor="hand2", padx=12, pady=5,
            borderwidth=0,
        )
        self.skip_button.pack(pady=(16, 0))

    def _center(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() - width) // 2
        y = (self.root.winfo_screenheight() - height) // 3
        self.root.geometry(f"+{x}+{y}")

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
        self.root.after(80, self._drain)

    def _on_finished(self):
        failed = [s for s in self.result if s.error]
        if not failed:
            self.root.after(_AUTOCLOSE_MS, self._close)
            return
        # Si algo fallo la ventana NO se cierra sola: el usuario tiene que enterarse de que va a
        # correr con el driver viejo (o sin driver) antes de lanzar una ejecucion.
        detail = "; ".join(f"{s.driver_name}: {s.error}" for s in failed)
        log_runtime(f"Drivers con problemas al iniciar: {detail}", level="WARN")
        blocking = [s for s in failed if s.state == MISSING]
        text = ("Falta un driver y no se pudo descargar. Revisa la conexion."
                if blocking else
                "No se pudo actualizar. La app abre igual con el driver actual.")
        tk.Label(self.root, text=text, bg=_BG, fg=_ERR,
                 font=("Segoe UI", 9), wraplength=430, justify="left").pack(padx=26)
        self.skip_button.configure(text="Continuar")
        self.root.update_idletasks()
        self._center()

    def _skip(self):
        self.cancelled = True
        self._close()

    def _close(self):
        try:
            self.root.destroy()
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
                )
            except Exception as exc:  # pragma: no cover - ensure_drivers_ready ya no lanza
                log_runtime(f"Fallo la actualizacion de drivers: {exc}", level="ERROR")
            finally:
                self.finished = True

        threading.Thread(target=worker, daemon=True, name="driver-updater").start()
        self.root.after(80, self._drain)
        self.root.mainloop()
        return self.result


def ensure_drivers_with_ui():
    """Chequea los drivers al abrir la app y muestra la ventana solo si hay que descargar algo.

    Devuelve la lista de DriverStatus (o None si no se pudo chequear). No lanza nunca: si algo
    sale mal la app tiene que abrir igual.
    """
    try:
        statuses = check_drivers()
    except Exception as exc:
        log_runtime(f"No se pudo chequear el estado de los drivers: {exc}", level="ERROR")
        return None

    if not needs_network(statuses):
        log_runtime("Drivers al dia, no hace falta descargar nada", level="INFO")
        return statuses

    try:
        return _DriverUpdateWindow(statuses).run()
    except Exception as exc:
        # Si Tk falla (sesion sin escritorio, etc.) igual conviene intentar la actualizacion
        # silenciosa antes de rendirse: sin driver valido la corrida no arranca.
        log_runtime(f"No se pudo mostrar la ventana de drivers: {exc}", level="WARN")
        try:
            return ensure_drivers_ready()
        except Exception:
            return statuses
