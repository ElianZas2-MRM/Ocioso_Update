"""
console_widget.py
-----------------
Panel de consola en tiempo real para la interfaz Osocio.
Redirige sys.stdout para capturar todos los print() del backend.
"""

import sys
import threading
from tkinter import Frame, Label, Button, Text, Scrollbar, END, BOTH, RIGHT, Y, X, BOTTOM, LEFT


APP_BG_COLOR    = "#2b1d3a"
CONSOLE_BG      = "#1a1a2e"
CONSOLE_FG      = "#d4d4d4"
CONSOLE_HEADER  = "#3d2a52"


class _TeeStream:
    """Escribe a dos streams simultáneamente (consola widget + stdout original)."""

    def __init__(self, widget_write, original):
        self._widget_write = widget_write
        self._original = original

    def write(self, text):
        if text:
            try:
                self._widget_write(text)
            except Exception:
                pass
            try:
                self._original.write(text)
            except Exception:
                pass

    def flush(self):
        try:
            self._original.flush()
        except Exception:
            pass

    def isatty(self):
        return False


class ConsoleWidget:
    """
    Panel de consola que muestra logs en tiempo real.
    Thread-safe: usa root.after() para actualizar el widget desde threads.
    """

    def __init__(self, parent, root):
        self._root = root
        self._lock = threading.Lock()

        # Frame contenedor
        outer = Frame(parent, bg=CONSOLE_HEADER, padx=1, pady=1)
        outer.pack(fill=X, side=BOTTOM, padx=20, pady=(0, 8))

        header = Frame(outer, bg=CONSOLE_HEADER)
        header.pack(fill=X)

        Label(
            header, text="Consola", bg=CONSOLE_HEADER, fg="white",
            font=("Consolas", 9, "bold"), padx=6, pady=3,
        ).pack(side=LEFT)

        Button(
            header, text="Limpiar", bg=CONSOLE_HEADER, fg="#aaa",
            font=("Consolas", 8), relief="flat", bd=0, cursor="hand2",
            command=self.clear, activebackground=CONSOLE_HEADER, activeforeground="white",
        ).pack(side=RIGHT, padx=4)

        inner = Frame(outer, bg=CONSOLE_BG)
        inner.pack(fill=BOTH, expand=True)

        scrollbar = Scrollbar(inner)
        scrollbar.pack(side=RIGHT, fill=Y)

        self._text = Text(
            inner, height=8, bg=CONSOLE_BG, fg=CONSOLE_FG,
            font=("Consolas", 9), relief="flat", bd=0,
            yscrollcommand=scrollbar.set,
            state="disabled", wrap="word",
            insertbackground=CONSOLE_FG,
        )
        self._text.pack(fill=BOTH, expand=True, padx=4, pady=4)
        scrollbar.config(command=self._text.yview)

        # Redirigir stdout
        self._original_stdout = sys.stdout
        sys.stdout = _TeeStream(self.write, self._original_stdout)

    def write(self, text: str):
        """Thread-safe: encola la actualización en el hilo principal de Tk."""
        self._root.after(0, self._append, text)

    def _append(self, text: str):
        self._text.configure(state="normal")
        self._text.insert(END, text)
        self._text.see(END)
        self._text.configure(state="disabled")

    def flush(self):
        pass

    def clear(self):
        self._text.configure(state="normal")
        self._text.delete("1.0", END)
        self._text.configure(state="disabled")

    def restore_stdout(self):
        sys.stdout = self._original_stdout
