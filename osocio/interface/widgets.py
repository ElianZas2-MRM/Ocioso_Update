"""Controles propios de la app que Tk no trae: interruptores, botones y tarjetas redondeados.

Tk 8.6 no tiene bordes redondeados ni antialias, así que estas piezas se dibujan con PIL
a varias veces su tamaño y se achican: quedan suaves en cualquier pantalla.

`BotonRedondo` y `Marco` se usan igual que `tk.Button` y `tk.Frame` (mismas opciones,
`config`, `cget`, `invoke`...), así que las pantallas los usan sin cambiar su lógica. Los
colores se escriben en clásico, como en todo el resto; `tema.color()` los traduce al
dibujar (ver tema.py).
"""
import tkinter as tk
import tkinter.font as tkfont

from PIL import Image, ImageDraw, ImageTk

from osocio.interface import tema

_SS = 4  # supersampling


def _hex(c):
    """(r, g, b) de un color clásico, ya traducido al tema actual."""
    c = tema.color(c)
    try:
        c = c.lstrip("#")
        return tuple(int(c[i:i + 2], 16) for i in (0, 2, 4))
    except (ValueError, AttributeError):
        return (128, 128, 128)


# --- interruptores ------------------------------------------------------------------------

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
              fill=(255, 255, 255, 255) if prendido else _hex("#C5B2D6") + (255,))
    lay = lay.resize((w, h), Image.LANCZOS)
    if not habilitado:
        lay.putalpha(lay.split()[3].point(lambda v: v * 0.4))
    return ImageTk.PhotoImage(lay)


class Toggle(tk.Frame):
    """Interruptor con texto al lado. Clic en el switch o en el texto lo cambia.

    Usa la MISMA `BooleanVar` que el Checkbutton al que reemplaza, así que todo lo que ya
    leía o guardaba esa variable sigue igual. `command`: se llama después de cada clic.
    `icono`: PhotoImage opcional entre el switch y el texto (p. ej. el escudo de Cadillac).
    """

    def __init__(self, master, text="", variable=None, command=None, bg="#4A2666",
                 fg="#E6D6F2", font=("Segoe UI", 8), color_on=tema.ACENTO,
                 color_off="#35164D", borde="#6E5A86", icono=None, w=34, h=18):
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
        tema.suscribir(self)
        self._pintar()

    def _repintar_tema(self):
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
        clave = (prendido, self._habilitado, tema.tema()) + self._colores
        if clave not in self._imgs:
            self._imgs[clave] = imagen_switch(prendido, *self._colores, *self._tam,
                                              habilitado=self._habilitado)
        try:
            self._sw.config(image=self._imgs[clave])
        except tk.TclError:
            pass

    def set_colores(self, color_on=None, color_off=None, borde=None, bg=None, fg=None):
        """Para cambiar de modo sin recrear el control."""
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

    def set_texto(self, texto):
        self._txt.config(text=texto)

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
        tema.suscribir(self)
        self._pintar()

    def _repintar_tema(self):
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
        clave = (t3, tema.tema())
        if clave not in self._imgs:
            self._imgs[clave] = imagen_selector(t3, COLOR_T3 if t3 else COLOR_T1)
        try:
            self._sw.config(image=self._imgs[clave])
            self._t1.config(fg=self._fg_apagado if t3 else COLOR_T1)
            self._t3.config(fg=COLOR_T3 if t3 else self._fg_apagado)
        except tk.TclError:
            pass


# --- piezas redondeadas -------------------------------------------------------------------

_CACHE_FONDOS = {}


def _rgb_hex(rgb):
    return "#%02x%02x%02x" % tuple(rgb)


def fondo_redondeado(w, h, relleno, borde=None, radio=10, grosor=1, borde_tema=False):
    """Rectángulo redondeado con antialias, como PhotoImage con esquinas transparentes.

    `relleno` y `borde` son clásicos; con `borde_tema=True` el borde ya viene en colores
    del tema (lo calculado a partir del relleno).
    """
    w, h = max(2, int(w)), max(2, int(h))
    clave = (w, h, relleno, borde, radio, grosor, borde_tema, tema.tema())
    img = _CACHE_FONDOS.get(clave)
    if img is not None and _viva(img):
        return img
    ss = 3
    lay = Image.new("RGBA", (w * ss, h * ss), (0, 0, 0, 0))
    d = ImageDraw.Draw(lay)
    r = max(1, min(radio, h // 2, w // 2)) * ss
    caja = (0, 0, w * ss - 1, h * ss - 1)
    if borde:
        b = tuple(int(borde[i:i + 2], 16) for i in (1, 3, 5)) if borde_tema else _hex(borde)
        d.rounded_rectangle(caja, radius=r, fill=_hex(relleno) + (255,),
                            outline=b + (255,), width=max(1, round(grosor * ss)))
    else:
        d.rounded_rectangle(caja, radius=r, fill=_hex(relleno) + (255,))
    img = ImageTk.PhotoImage(lay.resize((w, h), Image.LANCZOS))
    if len(_CACHE_FONDOS) > 800:
        _CACHE_FONDOS.clear()
    _CACHE_FONDOS[clave] = img
    return img


def _viva(img):
    """¿La imagen sigue existiendo en Tk? (si se cerró la ventana donde se creó, no)."""
    try:
        img._PhotoImage__photo.tk.call("image", "width", str(img))
        return True
    except (tk.TclError, AttributeError, RuntimeError):
        return False


def ajustar_scroll(canvas):
    """Área de scroll de un canvas que nunca es más chica que lo que se ve.

    Con `scrollregion=bbox("all")` a secas, si el contenido entra entero en la ventana el
    área queda más chica que el canvas y Tk igual deja correr la vista con la rueda: la
    página "baja" y queda un hueco vacío arriba. Así, si entra, no hay a dónde bajar.
    """
    try:
        caja = canvas.bbox("all")
        if not caja:
            return
        _x0, _y0, x1, y1 = caja
        visible = canvas.winfo_height()
        canvas.configure(scrollregion=(0, 0, x1, max(y1, visible)))
        if y1 <= visible:
            canvas.yview_moveto(0)
    except tk.TclError:
        pass


def bg_visible(widget):
    """El color que se VE detrás de un widget: el de su tarjeta si está en una redondeada."""
    v = getattr(widget, "_bg_visible", None)
    if v:
        return v
    try:
        return widget.cget("bg")
    except (tk.TclError, AttributeError):
        return "#5D3C7A"


_OPC_BOTON = {
    "text": "", "image": "", "compound": "none", "command": None, "bg": "#35164D",
    "fg": "#FFFFFF", "font": ("Segoe UI", 9), "padx": 1, "pady": 1, "relief": "flat",
    "bd": 0, "overrelief": "", "activebackground": "", "activeforeground": "",
    "highlightthickness": 0, "highlightbackground": "", "highlightcolor": "", "state": "normal",
    "width": 0, "height": 0, "anchor": "center", "justify": "center", "wraplength": 0,
    "disabledforeground": "", "textvariable": "", "default": "disabled", "repeatdelay": 0,
    "repeatinterval": 0, "underline": -1,
    "bordesutil": True,   # propia: borde apenas más claro que el relleno (False: sin borde)
}
_ALIAS = {"background": "bg", "foreground": "fg", "borderwidth": "bd"}


class BotonRedondo(tk.Canvas):
    """Un `tk.Button` con esquinas redondeadas. Acepta las mismas opciones."""

    def __init__(self, master=None, cnf=None, **kw):
        kw = {**(cnf or {}), **kw}
        self._o = dict(_OPC_BOTON)
        propias, del_canvas = self._separar(kw)
        self._o.update(propias)
        self._apretado = False
        self._traza_var = None
        self._fondo_ref = None
        del_canvas.setdefault("cursor", "hand2")
        super().__init__(master, highlightthickness=0, bd=0, bg=bg_visible(master), **del_canvas)
        self.bindtags((str(self), "BotonRedondo") + tuple(self.bindtags()[1:]))
        _clase_boton(self)
        self._seguir_textvariable()
        tema.suscribir(self)
        self._ajustar_tamano()

    # --- opciones ------------------------------------------------------------------------
    @staticmethod
    def _separar(kw):
        propias, otras = {}, {}
        for k, v in kw.items():
            k = _ALIAS.get(k, k)
            if k in _OPC_BOTON:
                propias[k] = v
            else:
                otras[k] = v
        return propias, otras

    def configure(self, cnf=None, **kw):
        if cnf is None and not kw:
            return super().configure()
        if isinstance(cnf, str):
            return self.cget(cnf)
        kw = {**(cnf or {}), **kw}
        propias, otras = self._separar(kw)
        if otras:
            super().configure(**otras)
        if propias:
            self._o.update(propias)
            if "textvariable" in propias:
                self._seguir_textvariable()
            self._ajustar_tamano()

    config = configure

    def cget(self, key):
        k = _ALIAS.get(key, key)
        if k in self._o:
            if k == "text" and self._o.get("textvariable"):
                try:
                    return self.getvar(str(self._o["textvariable"]))
                except tk.TclError:
                    pass
            return self._o[k]
        return super().cget(key)

    def __getitem__(self, key):
        return self.cget(key)

    def __setitem__(self, key, value):
        self.configure(**{key: value})

    def keys(self):
        return sorted(set(super().keys()) | set(_OPC_BOTON))

    def invoke(self):
        if str(self._o["state"]) == "disabled":
            return None
        cmd = self._o.get("command")
        if callable(cmd):
            return cmd()
        return None

    def flash(self):
        pass

    def _seguir_textvariable(self):
        var = self._o.get("textvariable")
        if var and hasattr(var, "trace_add") and self._traza_var is None:
            self._traza_var = var.trace_add("write", lambda *_: self._ajustar_tamano())

    def _repintar_tema(self):
        self._dibujar()

    # --- medidas y dibujo ----------------------------------------------------------------
    def _fuente(self):
        f = self._o["font"]
        root = self._root()
        # Por ventana: una Font queda atada a la raíz donde se creó (los tests abren varias).
        cache = _FUENTES.setdefault(str(id(root)), {})
        if getattr(root, "_osocio_fuentes", None) is not cache:
            cache.clear()
            root._osocio_fuentes = cache
        clave = str(f)
        if clave not in cache:
            try:
                cache[clave] = tkfont.Font(root=root, font=f)
            except tk.TclError:
                cache[clave] = tkfont.nametofont("TkDefaultFont", root=root)
        return cache[clave]

    def _imagen(self):
        img = self._o.get("image")
        if not img:
            return None, 0, 0
        try:
            nombre = str(img)
            return (nombre, int(self.tk.call("image", "width", nombre)),
                    int(self.tk.call("image", "height", nombre)))
        except tk.TclError:
            return None, 0, 0

    def _contenido(self):
        """(fuente, texto, imagen, iw, ih, tw, th, ancho total, alto total, compound)."""
        f = self._fuente()
        texto = str(self.cget("text") or "")
        lineas = texto.split("\n") if texto else []
        tw = max((f.measure(l) for l in lineas), default=0)
        th = f.metrics("linespace") * len(lineas) if lineas else 0
        img, iw, ih = self._imagen()
        comp = str(self._o.get("compound") or "none")
        if img and (not texto or comp == "none"):
            return f, "", img, iw, ih, 0, 0, iw, ih, "none"
        if img and comp in ("left", "right"):
            return f, texto, img, iw, ih, tw, th, iw + tw, max(ih, th), comp
        if img and comp in ("top", "bottom"):
            return f, texto, img, iw, ih, tw, th, max(iw, tw), ih + th, comp
        if img:
            return f, texto, img, iw, ih, tw, th, max(iw, tw), max(ih, th), "center"
        return f, texto, None, 0, 0, tw, th, tw, th, "none"

    def _ajustar_tamano(self):
        f, texto, img, iw, ih, tw, th, cw, ch, comp = self._contenido()
        try:
            ancho, alto = int(self._o.get("width") or 0), int(self._o.get("height") or 0)
        except (TypeError, ValueError):
            ancho, alto = 0, 0
        if ancho > 0:
            cw = ancho if (img and not texto) else ancho * f.measure("0")
        if alto > 0:
            ch = alto if (img and not texto) else alto * f.metrics("linespace")
        padx, pady = int(self._o.get("padx") or 0), int(self._o.get("pady") or 0)
        ht = int(self._o.get("highlightthickness") or 0)
        try:
            super().configure(width=cw + 2 * padx + 2 * ht + 10, height=ch + 2 * pady + 2 * ht + 6)
        except tk.TclError:
            return
        self._dibujar()

    def _dibujar(self):
        try:
            w, h = self.winfo_width(), self.winfo_height()
        except tk.TclError:
            return
        if w <= 2 or h <= 2:
            try:
                w = int(float(super().cget("width")))
                h = int(float(super().cget("height")))
            except (tk.TclError, ValueError):
                return
        self.delete("all")
        deshab = str(self._o["state"]) == "disabled"
        relleno = self._o.get("bg") or "#35164D"
        if self._apretado and self._o.get("activebackground"):
            relleno = self._o["activebackground"]
        ht = int(self._o.get("highlightthickness") or 0)
        radio = min(10, h // 2)
        try:
            super().configure(bg=bg_visible(self.master))
        except tk.TclError:
            pass
        if ht > 0 and self._o.get("highlightbackground"):
            fondo = fondo_redondeado(w, h, relleno, self._o["highlightbackground"], radio=radio,
                                     grosor=ht)
        elif not self._o.get("bordesutil"):
            fondo = fondo_redondeado(w, h, relleno, None, radio=radio)
        else:
            # Borde sutil del mismo tono: los botones se despegan del fondo en los dos temas.
            borde = _rgb_hex(min(255, round(c + (255 - c) * 0.12)) for c in _hex(relleno))
            fondo = fondo_redondeado(w, h, relleno, borde, radio=radio, borde_tema=True)
        self.create_image(0, 0, image=fondo, anchor="nw")
        self._fondo_ref = fondo

        f, texto, img, iw, ih, tw, th, cw, ch, comp = self._contenido()
        fg = self._o.get("fg") or "#FFFFFF"
        if self._apretado and self._o.get("activeforeground"):
            fg = self._o["activeforeground"]
        if deshab:
            fg = self._o.get("disabledforeground") or "#8A7A9C"
        anchor = str(self._o.get("anchor") or "center")
        margen = int(self._o.get("padx") or 0) + ht + 5
        if anchor in ("w", "nw", "sw"):
            x0 = margen
        elif anchor in ("e", "ne", "se"):
            x0 = w - margen - cw
        else:
            x0 = (w - cw) / 2
        y0 = (h - ch) / 2
        cy = h / 2
        just = str(self._o.get("justify") or "center")
        if img and comp == "none":
            self.create_image(x0 + cw / 2, cy, image=img)
        elif img and comp == "left":
            self.create_image(x0, cy, image=img, anchor="w")
            self.create_text(x0 + iw, cy, text=texto, font=f, fill=fg, anchor="w", justify=just)
        elif img and comp == "right":
            self.create_text(x0, cy, text=texto, font=f, fill=fg, anchor="w", justify=just)
            self.create_image(x0 + tw, cy, image=img, anchor="w")
        elif img and comp == "top":
            self.create_image(x0 + cw / 2, y0, image=img, anchor="n")
            self.create_text(x0 + cw / 2, y0 + ih, text=texto, font=f, fill=fg, anchor="n",
                             justify=just)
        elif img and comp == "bottom":
            self.create_text(x0 + cw / 2, y0, text=texto, font=f, fill=fg, anchor="n", justify=just)
            self.create_image(x0 + cw / 2, y0 + th, image=img, anchor="n")
        elif img:
            self.create_image(x0 + cw / 2, cy, image=img)
            self.create_text(x0 + cw / 2, cy, text=texto, font=f, fill=fg, justify=just)
        elif texto:
            self.create_text(x0 + cw / 2, cy, text=texto, font=f, fill=fg, justify=just)


_FUENTES = {}
_clase_lista = {"ok": False}


def _clase_boton(w):
    """Los eventos propios del botón, una sola vez para toda la clase."""
    if _clase_lista["ok"]:
        return
    _clase_lista["ok"] = True

    def yo(event):
        return event.widget if isinstance(event.widget, BotonRedondo) else None

    def press(event):
        b = yo(event)
        if b and str(b._o["state"]) != "disabled":
            b._apretado = True
            b._dibujar()

    def release(event):
        b = yo(event)
        if not b or not b._apretado:
            return
        b._apretado = False
        adentro = 0 <= event.x < b.winfo_width() and 0 <= event.y < b.winfo_height()
        b._dibujar()
        if adentro:
            b.invoke()

    def leave(event):
        b = yo(event)
        if b and b._apretado:
            b._apretado = False
            b._dibujar()

    def configure(event):
        b = yo(event)
        if b:
            b._dibujar()

    w.bind_class("BotonRedondo", "<ButtonPress-1>", press)
    w.bind_class("BotonRedondo", "<ButtonRelease-1>", release)
    w.bind_class("BotonRedondo", "<Leave>", leave, add="+")
    w.bind_class("BotonRedondo", "<Configure>", configure)


class Marco(tk.Frame):
    """Un `tk.Frame`. Si tiene borde (`highlightthickness`), se dibuja como tarjeta redondeada.

    El borde y el fondo pasan a un dibujo detrás de los hijos; el Frame en sí toma el color
    de lo que tiene alrededor, así las esquinas quedan transparentes.
    """

    def __init__(self, master=None, cnf=None, **kw):
        kw = {**(cnf or {}), **kw}
        ht = int(kw.get("highlightthickness", 0) or 0)
        self._redondo = ht > 0
        self._fondo = None
        if not self._redondo:
            super().__init__(master, **kw)
            return
        bg = kw.pop("bg", None) or kw.pop("background", None) or "#4A2666"
        self._card = {"bg": bg,
                      "highlightbackground": kw.pop("highlightbackground", None) or "#7D4E9F",
                      "highlightcolor": kw.pop("highlightcolor", ""),
                      "highlightthickness": ht}
        kw["highlightthickness"] = 0
        super().__init__(master, bg=bg_visible(master), **kw)
        self._bg_visible = bg
        self.bind("<Configure>", lambda e: self._dibujar(), add="+")
        tema.suscribir(self)
        self._asegurar_fondo()

    def _asegurar_fondo(self):
        try:
            if self._fondo is not None and self._fondo.winfo_exists():
                return self._fondo
        except tk.TclError:
            pass
        self._fondo = tk.Canvas(self, highlightthickness=0, bd=0, bg=bg_visible(self.master))
        self._fondo.place(x=0, y=0, relwidth=1, relheight=1)
        self.tk.call("lower", self._fondo._w)  # Canvas.lower es para ítems, no ventanas
        for ev in ("<Button-1>", "<ButtonRelease-1>", "<Double-Button-1>"):
            self._fondo.bind(ev, lambda e, ev=ev: self.event_generate(ev, x=e.x, y=e.y))
        return self._fondo

    def _dibujar(self):
        if not self._redondo:
            return
        try:
            w, h = self.winfo_width(), self.winfo_height()
        except tk.TclError:
            return
        if w < 4 or h < 4:
            return
        c = self._asegurar_fondo()
        afuera = bg_visible(self.master)
        try:
            tk.Frame.configure(self, bg=afuera)
            c.configure(bg=afuera)
        except tk.TclError:
            return
        borde = self._card["highlightbackground"] if self._card["highlightthickness"] else None
        img = fondo_redondeado(w, h, self._card["bg"], borde, radio=8 if h < 70 else 12,
                               grosor=max(1, self._card["highlightthickness"]))
        c.delete("all")
        c.create_image(0, 0, image=img, anchor="nw")
        c._img = img

    def _repintar_tema(self):
        self._dibujar()

    def configure(self, cnf=None, **kw):
        if not getattr(self, "_redondo", False):
            return super().configure(cnf, **kw)
        if cnf is None and not kw:
            return super().configure()
        if isinstance(cnf, str):
            return self.cget(cnf)
        kw = {**(cnf or {}), **kw}
        cambio = False
        for k in ("bg", "background", "highlightbackground", "highlightcolor", "highlightthickness"):
            if k in kw:
                v = kw.pop(k)
                k = "bg" if k == "background" else k
                self._card[k] = int(v or 0) if k == "highlightthickness" else v
                cambio = True
        if kw:
            super().configure(**kw)
        if cambio:
            self._bg_visible = self._card["bg"]
            self._dibujar()

    config = configure

    def cget(self, key):
        if getattr(self, "_redondo", False):
            k = "bg" if key == "background" else key
            if k in self._card:
                return self._card[k]
        return super().cget(key)

    def __getitem__(self, key):
        return self.cget(key)

    def __setitem__(self, key, value):
        self.configure(**{key: value})


class Segmentado(Marco):
    """Opciones excluyentes en una cápsula: [ 1 por país | Consolidado ]. Maneja una StringVar."""

    def __init__(self, master, variable, opciones, command=None, bg="#35164D",
                 font=("Segoe UI", 9), fg="#C5B2D6"):
        super().__init__(master, bg=bg, highlightthickness=1, highlightbackground="#7D4E9F")
        self.variable = variable
        self._bg, self._fg = bg, fg
        self.botones = {}
        for valor, texto in opciones:
            b = BotonRedondo(self, text=texto, bg=bg, fg=fg, font=font, padx=10, pady=1,
                             bordesutil=False,
                             command=lambda v=valor: self._elegir(v, command))
            b.pack(side="left", padx=(3, 0) if self.botones else (3, 0), pady=3)
            self.botones[valor] = b
        tk.Frame(self, bg=bg, width=3).pack(side="left")
        self._traza = variable.trace_add("write", lambda *_: self._pintar())
        self.bind("<Destroy>", self._al_destruir, add="+")
        self._pintar()

    def _al_destruir(self, event):
        if event.widget is self:
            try:
                self.variable.trace_remove("write", self._traza)
            except Exception:
                pass

    def _elegir(self, valor, command):
        self.variable.set(valor)
        if command:
            command()

    def _pintar(self):
        actual = self.variable.get()
        for v, b in self.botones.items():
            try:
                if v == actual:
                    b.config(bg="#602D8A", fg="#FFFFFF")
                else:
                    b.config(bg=self._bg, fg=self._fg)
            except tk.TclError:
                pass

