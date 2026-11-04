"""Piezas de las ventanas de "procesando": spinner, barra redondeada y cronómetro.

Están hechas para reemplazar a lo que ya usaban esas ventanas sin tocar su lógica:
  - `IconoEstado` en lugar del Label con "↻": las ventanas le hacen `config(text="➔")`
    para animarlo (eso se ignora, gira solo) y al final `config(text="✓", fg=...)`, que
    dibuja el círculo con la marca.
  - `BarraProgreso` en lugar del Canvas de la barra: acepta el mismo
    `create_rectangle` / `coords` / `itemconfig(fill=...)` que hacían sobre el rectángulo,
    pero la dibuja redondeada y con antialias.
  - `Cronometro`: "⏱ 01:23" mientras el ícono gira; se frena solo cuando termina.
"""
import time
import tkinter as tk

from osocio.interface import tema
from osocio.interface.widgets import bg_visible, fondo_redondeado

_GIRANDO = ("↻", "➔")


class IconoEstado(tk.Canvas):
    """Spinner mientras corre; círculo con ✓ / ✕ / ■ al terminar."""

    def __init__(self, master, tam=30, **_kw):
        super().__init__(master, width=tam, height=tam, highlightthickness=0, bd=0,
                         bg=bg_visible(master))
        self._tam = tam
        self._angulo = 0
        self._job = None
        self.girando = False
        self._marca = None
        tema.suscribir(self)
        self.girar()

    def girar(self):
        self.girando = True
        self._marca = None
        if self._job is None:
            self._paso()

    def _paso(self):
        if not self.girando:
            self._job = None
            return
        try:
            t = self._tam
            self.delete("all")
            # Colores clásicos: Tk ya los traduce al tema al dibujar (ver tema.py).
            self.create_oval(4, 4, t - 4, t - 4, outline="#4C226E", width=3)
            self.create_arc(4, 4, t - 4, t - 4, start=self._angulo, extent=110, style="arc",
                            outline=tema.ACENTO, width=3)
        except tk.TclError:
            self._job = None
            return
        self._angulo = (self._angulo - 20) % 360
        self._job = self.after(40, self._paso)

    def _dibujar_marca(self, simbolo, color):
        self.girando = False
        if self._job is not None:
            try:
                self.after_cancel(self._job)
            except tk.TclError:
                pass
            self._job = None
        self._marca = (simbolo, color)
        t = self._tam
        try:
            self.delete("all")
            self.create_oval(2, 2, t - 2, t - 2, fill=color, outline=color)
            self.create_text(t / 2, t / 2 + 1, text=simbolo, fill="#110518",
                             font=("Segoe UI", int(t * 0.42), "bold"))
        except tk.TclError:
            pass

    def _repintar_tema(self):
        if self._marca:
            self._dibujar_marca(*self._marca)

    def configure(self, cnf=None, **kw):
        kw = {**(cnf or {}), **kw}
        texto = kw.pop("text", None)
        color = kw.pop("fg", None) or kw.pop("foreground", None)
        kw.pop("font", None)
        if kw:
            super().configure(**kw)
        if texto is None:
            return
        if texto in _GIRANDO:
            if not self.girando:
                self.girar()
            return
        self._dibujar_marca(texto, color or "#82E0AA")

    config = configure


class BarraProgreso(tk.Canvas):
    """Barra de progreso redondeada. `bg` es el color de la pista."""

    _ITEM = "barra"

    def __init__(self, master, height=8, bg="#35164D", **_kw):
        alto = max(6, int(height))
        super().__init__(master, height=alto, highlightthickness=0, bd=0, bg=bg_visible(master))
        self._pista = bg
        self._color = "#F8C471"
        self._ancho_lleno = 0
        self._imgs = ()
        self.bind("<Configure>", lambda e: self._dibujar(), add="+")
        tema.suscribir(self)

    # --- lo que antes se hacía sobre el rectángulo ---------------------------------------
    def create_rectangle(self, *coords, **kw):
        if kw.get("fill"):
            self._color = kw["fill"]
        if len(coords) >= 3:
            self._ancho_lleno = coords[2]
        self._dibujar()
        return self._ITEM

    def coords(self, item, *args):
        if item == self._ITEM:
            if len(args) == 1 and isinstance(args[0], (list, tuple)):
                args = tuple(args[0])
            if len(args) >= 3:
                self._ancho_lleno = args[2]
                self._dibujar()
            return [0, 0, self._ancho_lleno, self._alto()]
        return super().coords(item, *args)

    def itemconfigure(self, item, cnf=None, **kw):
        if item == self._ITEM:
            kw = {**(cnf or {}), **kw}
            if kw.get("fill"):
                self._color = kw["fill"]
                self._dibujar()
            return None
        return super().itemconfigure(item, cnf, **kw)

    itemconfig = itemconfigure

    def _repintar_tema(self):
        self._dibujar()

    # --- dibujo --------------------------------------------------------------------------
    def _alto(self):
        try:
            return int(float(self.cget("height")))
        except (tk.TclError, ValueError):
            return 8

    def _dibujar(self):
        try:
            w, h = self.winfo_width(), self._alto()
        except tk.TclError:
            return
        if w < 4:
            return
        self.delete("all")
        pista = fondo_redondeado(w, h, self._pista, radio=h // 2)
        self.create_image(0, 0, image=pista, anchor="nw")
        lleno = max(0, min(w, int(self._ancho_lleno or 0)))
        imgs = [pista]
        if lleno >= h:            # más angosto que alto no se ve redondo: se espera un poco
            barra = fondo_redondeado(lleno, h, self._color, radio=h // 2)
            self.create_image(0, 0, image=barra, anchor="nw")
            imgs.append(barra)
        self._imgs = tuple(imgs)


class Cronometro(tk.Label):
    """'⏱ 01:23' mientras `icono` gira. Se frena solo cuando la corrida termina."""

    def __init__(self, master, icono, bg=None, fg="#9B86B5"):
        super().__init__(master, text="⏱ 00:00", font=("Segoe UI", 8), bg=bg or bg_visible(master),
                         fg=fg)
        self._icono = icono
        self._inicio = time.monotonic()
        self._tic()

    def _tic(self):
        try:
            if not self.winfo_exists():
                return
            seg = int(time.monotonic() - self._inicio)
            h, resto = divmod(seg, 3600)
            m, s = divmod(resto, 60)
            txt = f"{h}:{m:02d}:{s:02d}" if h else f"{m:02d}:{s:02d}"
            if getattr(self._icono, "girando", False):
                self.config(text=f"⏱ {txt} transcurrido")
                self.after(1000, self._tic)
            else:
                self.config(text=f"⏱ {txt} en total")
        except tk.TclError:
            pass
