"""Controles propios de la app que Tk no trae: el interruptor (toggle).

Tk 8.6 no tiene bordes redondeados ni antialias, así que el interruptor se dibuja con PIL
a 4x y se achica: queda suave en cualquier pantalla. Usa la MISMA `BooleanVar` que el
Checkbutton al que reemplaza, así que todo lo que ya leía o guardaba esa variable sigue
igual.
"""
import tkinter as tk

from PIL import Image, ImageDraw, ImageTk

_SS = 4  # supersampling


def _hex(c):
    c = c.lstrip("#")
    return tuple(int(c[i:i + 2], 16) for i in (0, 2, 4))


def imagen_switch(prendido, color_on, color_off, borde, w=34, h=18, habilitado=True):
    """El switch como PhotoImage: pista redondeada y perilla blanca."""
    lay = Image.new("RGBA", (w * _SS, h * _SS), (0, 0, 0, 0))
    d = ImageDraw.Draw(lay)
    caja = (0, 0, w * _SS - 1, h * _SS - 1)
    if prendido:
        d.rounded_rectangle(caja, radius=h * _SS // 2, fill=_hex(color_on) + (255,))
    else:
        d.rounded_rectangle(caja, radius=h * _SS // 2, fill=_hex(color_off) + (255,),
                            outline=_hex(borde) + (255,), width=int(1.5 * _SS))
    k = (h - 6) * _SS
    kx = w * _SS - 3 * _SS - k if prendido else 3 * _SS
    d.ellipse((kx, 3 * _SS, kx + k, 3 * _SS + k),
              fill=(255, 255, 255, 255) if prendido else _hex(borde) + (255,))
    lay = lay.resize((w, h), Image.LANCZOS)
    if not habilitado:
        lay.putalpha(lay.split()[3].point(lambda v: v * 0.4))
    return ImageTk.PhotoImage(lay)


class Toggle(tk.Frame):
    """Interruptor con texto al lado. Clic en el switch o en el texto lo cambia.

    `variable`: la BooleanVar que maneja. `command`: se llama después de cada clic.
    `icono`: PhotoImage opcional entre el switch y el texto (p. ej. el escudo de Cadillac).
    """

    def __init__(self, master, text="", variable=None, command=None, bg="#4A2666",
                 fg="#E6D6F2", font=("Segoe UI", 8), color_on="#7C3AED",
                 color_off="#35164D", borde="#7D4E9F", icono=None, w=34, h=18):
        super().__init__(master, bg=bg)
        self.variable = variable if variable is not None else tk.BooleanVar(value=False)
        self.command = command
        self._colores = (color_on, color_off, borde)
        self._tam = (w, h)
        self._habilitado = True
        self._imgs = {}
        self._sw = tk.Label(self, bg=bg, bd=0, cursor="hand2")
        self._sw.pack(side="left")
        self._icono = None
        if icono is not None:
            self._icono = tk.Label(self, image=icono, bg=bg, bd=0, cursor="hand2")
            self._icono.image = icono
            self._icono.pack(side="left", padx=(6, 0))
        self._txt = tk.Label(self, text=text, bg=bg, fg=fg, font=font, cursor="hand2")
        self._txt.pack(side="left", padx=(6, 0))
        self._fg = fg
        for w_ in (self._sw, self._txt, self._icono):
            if w_ is not None:
                w_.bind("<Button-1>", self._clic)
        self._traza = self.variable.trace_add("write", lambda *_: self._pintar())
        self.bind("<Destroy>", self._al_destruir, add="+")
        self._pintar()

    def _al_destruir(self, event):
        if event.widget is self:
            try:
                self.variable.trace_remove("write", self._traza)
            except Exception:
                pass

    def _clic(self, _event=None):
        if not self._habilitado:
            return
        self.variable.set(not bool(self.variable.get()))
        if self.command:
            self.command()

    def _pintar(self):
        prendido = bool(self.variable.get())
        clave = (prendido, self._habilitado) + self._colores
        if clave not in self._imgs:
            self._imgs[clave] = imagen_switch(prendido, *self._colores, *self._tam,
                                              habilitado=self._habilitado)
        try:
            self._sw.config(image=self._imgs[clave])
        except tk.TclError:
            pass

    def set_colores(self, color_on=None, color_off=None, borde=None, bg=None, fg=None):
        """Para cambiar de tema o de modo sin recrear el control."""
        on, off, b = self._colores
        self._colores = (color_on or on, color_off or off, borde or b)
        if bg:
            for w_ in (self, self._sw, self._txt, self._icono):
                if w_ is not None:
                    w_.config(bg=bg)
        if fg:
            self._fg = fg
            self._txt.config(fg=fg)
        self._pintar()

    def config_estado(self, habilitado):
        self._habilitado = bool(habilitado)
        cursor = "hand2" if habilitado else "arrow"
        for w_ in (self._sw, self._txt, self._icono):
            if w_ is not None:
                w_.config(cursor=cursor)
        self._txt.config(fg=self._fg if habilitado else "#8a7a9c")
        self._pintar()


COLOR_T1 = "#E2B33C"   # dorado
COLOR_T3 = "#3DC1E3"   # turquesa, misma intensidad


def imagen_selector(derecha, color, w=42, h=22):
    """Switch siempre lleno con el color del lado elegido; la perilla marca el lado."""
    lay = Image.new("RGBA", (w * _SS, h * _SS), (0, 0, 0, 0))
    d = ImageDraw.Draw(lay)
    d.rounded_rectangle((0, 0, w * _SS - 1, h * _SS - 1), radius=h * _SS // 2,
                        fill=_hex(color) + (255,))
    k = (h - 6) * _SS
    kx = w * _SS - 3 * _SS - k if derecha else 3 * _SS
    d.ellipse((kx, 3 * _SS, kx + k, 3 * _SS + k), fill=(255, 255, 255, 255))
    return ImageTk.PhotoImage(lay.resize((w, h), Image.LANCZOS))


class SelectorModo(tk.Frame):
    """MODO  T1 (switch) T3. `variable` es una StringVar con "T1" o "T3"."""

    def __init__(self, master, variable, command=None, bg="#5D3C7A", fg_apagado="#9b8aa8",
                 titulo="MODO", fg_titulo="#E6D6F2"):
        super().__init__(master, bg=bg)
        self.variable = variable
        self.command = command
        self._fg_apagado = fg_apagado
        self._imgs = {}
        self._titulo = None
        if titulo:
            self._titulo = tk.Label(self, text=titulo, bg=bg, fg=fg_titulo,
                                    font=("Segoe UI", 7, "bold"))
            self._titulo.pack(side="left", padx=(0, 8))
        self._t1 = tk.Label(self, text="T1", bg=bg, font=("Segoe UI", 10, "bold"), cursor="hand2")
        self._sw = tk.Label(self, bg=bg, bd=0, cursor="hand2")
        self._t3 = tk.Label(self, text="T3", bg=bg, font=("Segoe UI", 10, "bold"), cursor="hand2")
        self._t1.pack(side="left")
        self._sw.pack(side="left", padx=6)
        self._t3.pack(side="left")
        self._t1.bind("<Button-1>", lambda e: self._elegir("T1"))
        self._t3.bind("<Button-1>", lambda e: self._elegir("T3"))
        self._sw.bind("<Button-1>", lambda e: self._elegir("T1" if self.es_t3() else "T3"))
        self._traza = variable.trace_add("write", lambda *_: self._pintar())
        self.bind("<Destroy>", self._al_destruir, add="+")
        self._pintar()

    def es_t3(self):
        return str(self.variable.get()).upper() == "T3"

    def _al_destruir(self, event):
        if event.widget is self:
            try:
                self.variable.trace_remove("write", self._traza)
            except Exception:
                pass

    def _elegir(self, modo):
        if str(self.variable.get()).upper() == modo:
            return
        self.variable.set(modo)
        if self.command:
            self.command()

    def _pintar(self):
        t3 = self.es_t3()
        if t3 not in self._imgs:
            self._imgs[t3] = imagen_selector(t3, COLOR_T3 if t3 else COLOR_T1)
        try:
            self._sw.config(image=self._imgs[t3])
            self._t1.config(fg=self._fg_apagado if t3 else COLOR_T1)
            self._t3.config(fg=COLOR_T3 if t3 else self._fg_apagado)
        except tk.TclError:
            pass

    def set_colores(self, bg=None, fg_apagado=None, fg_titulo=None):
        if fg_apagado:
            self._fg_apagado = fg_apagado
        if bg:
            for w_ in (self, self._t1, self._sw, self._t3, self._titulo):
                if w_ is not None:
                    w_.config(bg=bg)
        if fg_titulo and self._titulo is not None:
            self._titulo.config(fg=fg_titulo)
        self._pintar()
