"""Menú lateral de la app: el oso, el modo T1/T3, las secciones y, abajo, el tema y los drivers.

Se puede achicar a una tira de íconos (botón «) para que no ocupe lugar; se recuerda.
"""
import os
import tkinter as tk

from PIL import Image, ImageDraw, ImageFont, ImageTk

from osocio.interface import tema
from osocio.interface.widgets import (COLOR_T1, COLOR_T3, BotonRedondo, Marco, SelectorModo,
                                      Toggle, _hex)

ANCHO = 228
ANCHO_COMPACTO = 68

_TXT = "#E6D6F2"
_TXT_GRUPO = "#9B86B5"
_HOVER = tema.SIDEBAR_HOVER


def _icono(asset_dir, nombre, color, tam=16):
    """Ícono de tabler teñido del color pedido (clásico, se traduce al tema)."""
    try:
        im = Image.open(os.path.join(asset_dir, "tabler_icons", nombre)).convert("RGBA")
    except (OSError, ValueError):
        return None
    im = im.resize((tam, tam), Image.LANCZOS)
    lleno = Image.new("RGBA", im.size, _hex(color) + (255,))
    lleno.putalpha(im.split()[3])
    return ImageTk.PhotoImage(lleno)


def _emoji(caracter, tam=16):
    """Un emoji a color como imagen (Tk no los dibuja a color)."""
    try:
        f = ImageFont.truetype("C:/Windows/Fonts/seguiemj.ttf", 64)
    except OSError:
        return None
    lay = Image.new("RGBA", (96, 96), (0, 0, 0, 0))
    ImageDraw.Draw(lay).text((8, 8), caracter, font=f, embedded_color=True)
    caja = lay.getbbox()
    if not caja:
        return None
    lay = lay.crop(caja)
    lay.thumbnail((tam, tam), Image.LANCZOS)
    return ImageTk.PhotoImage(lay)


def _oso(asset_dir, ancho, redondo=False):
    try:
        im = Image.open(os.path.join(asset_dir, "osopng.png")).convert("RGBA")
    except OSError:
        return None
    if redondo:
        lado = min(im.size)
        im = im.crop(((im.width - lado) // 2, 0, (im.width - lado) // 2 + lado, lado))
        alto = ancho
    else:
        alto = round(im.height * ancho / im.width)
    ss = 3
    im = im.resize((ancho * ss, alto * ss), Image.LANCZOS)
    mask = Image.new("L", im.size, 0)
    d = ImageDraw.Draw(mask)
    if redondo:
        d.ellipse((0, 0, im.width - 1, im.height - 1), fill=255)
    else:
        d.rounded_rectangle((0, 0, im.width - 1, im.height - 1), radius=14 * ss, fill=255)
    im.putalpha(mask)
    return ImageTk.PhotoImage(im.resize((ancho, alto), Image.LANCZOS))


class MenuLateral(tk.Frame):
    """`grupos`: [(titulo, [(clave, texto, icono), ...]), ...]. `al_elegir(clave)` al hacer clic."""

    def __init__(self, master, asset_dir, grupos, al_elegir, var_modo, al_cambiar_modo,
                 al_cambiar_tema, al_actualizar_drivers, compacto=False, al_compactar=None):
        super().__init__(master, bg=tema.SIDEBAR_BG)
        self.asset_dir = asset_dir
        self.grupos = grupos
        self.al_elegir = al_elegir
        self.var_modo = var_modo
        self.al_cambiar_modo = al_cambiar_modo
        self.al_cambiar_tema = al_cambiar_tema
        self.al_actualizar_drivers = al_actualizar_drivers
        self.al_compactar = al_compactar
        self.compacto = bool(compacto)
        self.activa = None
        self.botones = {}
        self.btn_drivers = None
        self._imgs = []
        self.var_oscuro = tk.BooleanVar(value=tema.es_oscuro())
        self._traza_modo = var_modo.trace_add("write", lambda *_: self._refrescar_modo())
        self._sep = tk.Frame(master, bg=tema.SIDEBAR_SEP, width=1)
        tema.suscribir(self)
        self._construir()

    def pack(self, **kw):
        super().pack(side="left", fill="y", **kw)
        self.pack_propagate(False)
        self._sep.pack(side="left", fill="y", after=self)

    # --- armado -------------------------------------------------------------------------
    def _construir(self):
        for c in self.winfo_children():
            c.destroy()
        self._imgs = []
        self.botones = {}
        self.configure(width=ANCHO_COMPACTO if self.compacto else ANCHO)
        if self.compacto:
            self._construir_compacto()
        else:
            self._construir_expandido()
        if self.activa:
            self.marcar(self.activa)

    def _img(self, img):
        if img is not None:
            self._imgs.append(img)
        return img

    def _construir_expandido(self):
        bg = tema.SIDEBAR_BG
        arriba = tk.Frame(self, bg=bg)
        arriba.pack(fill="x", padx=14, pady=(14, 0))
        oso = self._img(_oso(self.asset_dir, ANCHO - 28))
        if oso:
            tk.Label(arriba, image=oso, bg=bg, bd=0).pack(anchor="w")
        tit = tk.Frame(arriba, bg=bg)
        tit.pack(anchor="w", pady=(4, 10))
        tk.Label(tit, text="Osocio", font=("Segoe UI", 11, "bold"), bg=bg, fg="#FFFFFF").pack(side="left")
        tk.Label(tit, text="Form Automation", font=("Segoe UI", 8), bg=bg, fg=_TXT_GRUPO).pack(
            side="left", padx=(5, 0), pady=(3, 0))

        caja = Marco(arriba, bg=tema.MODO_CAJA, highlightthickness=1,
                     highlightbackground=tema.SIDEBAR_SEP)
        caja.pack(fill="x", pady=(0, 6))
        tk.Label(caja, text="MODO", font=("Segoe UI", 7, "bold"), bg=tema.MODO_CAJA,
                 fg=_TXT).pack(anchor="w", padx=12, pady=(8, 0))
        SelectorModo(caja, self.var_modo, command=self.al_cambiar_modo, bg=tema.MODO_CAJA,
                     fg_apagado="#9B86B5", titulo="").pack(anchor="w", padx=12, pady=(2, 0))
        self._lbl_excels = tk.Label(caja, text="", font=("Segoe UI", 8, "italic"),
                                    bg=tema.MODO_CAJA)
        self._lbl_excels.pack(anchor="w", padx=12, pady=(1, 8))
        self._refrescar_modo()

        medio = tk.Frame(self, bg=bg)
        medio.pack(fill="x", padx=10, pady=(4, 0))
        for titulo, items in self.grupos:
            tk.Label(medio, text=titulo.upper(), font=("Segoe UI", 7, "bold"), bg=bg,
                     fg=_TXT_GRUPO).pack(anchor="w", padx=8, pady=(8, 2))
            for clave, texto, icono in items:
                self._item(medio, clave, "  " + texto, icono)

        abajo = tk.Frame(self, bg=bg)
        abajo.pack(side="bottom", fill="x", padx=12, pady=(0, 12))
        tk.Frame(abajo, bg=tema.SIDEBAR_SEP, height=1).pack(fill="x", pady=(0, 8))
        fila_tema = tk.Frame(abajo, bg=bg)
        fila_tema.pack(fill="x", padx=4)
        luna = self._img(_emoji("🌙" if tema.es_oscuro() else "☀️"))
        tk.Label(fila_tema, image=luna, bg=bg).pack(side="left")
        tk.Label(fila_tema, text="Tema oscuro", font=("Segoe UI", 9), bg=bg, fg=_TXT).pack(
            side="left", padx=(8, 0))
        Toggle(fila_tema, text="", variable=self.var_oscuro, command=self._tema_clic, bg=bg,
               color_off="#35164D", borde="#7D4E9F").pack(side="right")
        fila_dr = tk.Frame(abajo, bg=bg)
        fila_dr.pack(fill="x", pady=(8, 0))
        self.btn_drivers = BotonRedondo(
            fila_dr, text=" Actualizar drivers", anchor="w", compound="left",
            image=self._img(_icono(self.asset_dir, "refresh_lavender.png", _TXT, 14)),
            bg=tema.SIDEBAR_BG, fg=_TXT, activebackground=_HOVER, font=("Segoe UI", 9),
            padx=2, pady=4, command=self.al_actualizar_drivers)
        self.btn_drivers.pack(side="left", fill="x", expand=True)
        self._hover(self.btn_drivers)
        b = BotonRedondo(fila_dr, text="«", font=("Segoe UI", 11, "bold"), bg="#35164D",
                         fg=_TXT, padx=6, pady=0, command=self._alternar)
        b.pack(side="right", padx=(6, 0))

    def _construir_compacto(self):
        bg = tema.SIDEBAR_BG
        arriba = tk.Frame(self, bg=bg)
        arriba.pack(fill="x", pady=(14, 0))
        oso = self._img(_oso(self.asset_dir, 44, redondo=True))
        if oso:
            tk.Label(arriba, image=oso, bg=bg, bd=0).pack()
        self._chip_modo = BotonRedondo(arriba, text=self.var_modo.get(), font=("Segoe UI", 10, "bold"),
                                       bg=tema.MODO_CAJA, padx=6, pady=6, width=3,
                                       highlightthickness=1,
                                       command=self._alternar_modo)
        self._chip_modo.pack(pady=(12, 6))
        self._refrescar_modo()
        medio = tk.Frame(self, bg=bg)
        medio.pack(fill="x", pady=(6, 0))
        for _titulo, items in self.grupos:
            tk.Frame(medio, bg=tema.SIDEBAR_SEP, height=1).pack(fill="x", padx=16, pady=6)
            for clave, texto, icono in items:
                self._item(medio, clave, "", icono, tooltip=texto)
        abajo = tk.Frame(self, bg=bg)
        abajo.pack(side="bottom", fill="x", pady=(0, 12))
        luna = self._img(_emoji("🌙" if tema.es_oscuro() else "☀️"))
        bt = BotonRedondo(abajo, image=luna, bg=bg, padx=8, pady=6,
                          command=lambda: (self.var_oscuro.set(not self.var_oscuro.get()),
                                           self._tema_clic()))
        bt.pack(pady=2)
        self._hover(bt)
        self.btn_drivers = BotonRedondo(
            abajo, image=self._img(_icono(self.asset_dir, "refresh_lavender.png", _TXT, 16)),
            bg=bg, padx=8, pady=6, command=self.al_actualizar_drivers)
        self.btn_drivers.pack(pady=2)
        self._hover(self.btn_drivers)
        _Tooltip(self.btn_drivers, "Actualizar drivers")
        BotonRedondo(abajo, text="»", font=("Segoe UI", 11, "bold"), bg="#35164D", fg=_TXT,
                     padx=6, pady=0, command=self._alternar).pack(pady=(6, 0))

    def _item(self, padre, clave, texto, icono, tooltip=None):
        img = self._img(_icono(self.asset_dir, icono, _TXT, 16))
        img_sel = self._img(_icono(self.asset_dir, icono, "#FFFFFF", 16))
        b = BotonRedondo(padre, text=texto, image=img, compound="left", anchor="w",
                         bg=tema.SIDEBAR_BG, fg=_TXT, activebackground=_HOVER,
                         font=("Segoe UI", 10), padx=8 if texto else 10, pady=6,
                         command=lambda: self.al_elegir(clave))
        b.pack(fill="x" if texto else None, pady=1)
        b.config(bordesutil=False)
        b._imgs_menu = (img, img_sel)
        self._hover(b, clave)
        if tooltip:
            _Tooltip(b, tooltip)
        self.botones[clave] = b

    def _hover(self, b, clave=None):
        def entra(_e):
            if clave is None or clave != self.activa:
                b.config(bg=_HOVER)

        def sale(_e):
            if clave is None or clave != self.activa:
                b.config(bg=tema.SIDEBAR_BG)
        b.bind("<Enter>", entra)
        b.bind("<Leave>", sale)

    # --- estado -------------------------------------------------------------------------
    def marcar(self, clave):
        """Pinta como elegida la sección `clave`."""
        self.activa = clave
        for k, b in self.botones.items():
            normal, sel = getattr(b, "_imgs_menu", (None, None))
            if k == clave:
                b.config(bg=tema.MENU_SEL, fg="#FFFFFF", image=sel or "",
                         font=("Segoe UI", 10, "bold"))
            else:
                b.config(bg=tema.SIDEBAR_BG, fg=_TXT, image=normal or "", font=("Segoe UI", 10))

    def _refrescar_modo(self):
        t3 = str(self.var_modo.get()).upper() == "T3"
        color = COLOR_T3 if t3 else COLOR_T1
        try:
            if not self.compacto and getattr(self, "_lbl_excels", None) is not None:
                self._lbl_excels.config(text=f"Excels …_{'T3' if t3 else 'T1'}.xlsx", fg=color)
            if self.compacto and getattr(self, "_chip_modo", None) is not None:
                self._chip_modo.config(text="T3" if t3 else "T1", fg=color, highlightbackground=color)
        except tk.TclError:
            pass

    def _alternar_modo(self):
        self.var_modo.set("T1" if str(self.var_modo.get()).upper() == "T3" else "T3")
        if self.al_cambiar_modo:
            self.al_cambiar_modo()

    def _alternar(self):
        self.compacto = not self.compacto
        self._construir()
        if self.al_compactar:
            self.al_compactar(self.compacto)

    def _tema_clic(self):
        if self.al_cambiar_tema:
            self.al_cambiar_tema(tema.OSCURO if self.var_oscuro.get() else tema.CLASICO)

    def _repintar_tema(self):
        # El emoji de la luna/sol y los íconos teñidos dependen del tema: se rearma.
        self.var_oscuro.set(tema.es_oscuro())
        self._construir()

    def set_drivers_estado(self, estado):
        try:
            if self.btn_drivers is not None:
                self.btn_drivers.config(state=estado)
        except tk.TclError:
            pass


class _Tooltip:
    """Cartelito con el nombre de la sección, para la tira de íconos."""

    def __init__(self, widget, texto):
        self.widget, self.texto, self.win = widget, texto, None
        widget.bind("<Enter>", self._mostrar, add="+")
        widget.bind("<Leave>", self._ocultar, add="+")

    def _mostrar(self, _e):
        if self.win is not None:
            return
        x = self.widget.winfo_rootx() + self.widget.winfo_width() + 8
        y = self.widget.winfo_rooty() + 4
        self.win = tk.Toplevel(self.widget)
        self.win.wm_overrideredirect(True)
        self.win.geometry(f"+{x}+{y}")
        tk.Label(self.win, text=self.texto, font=("Segoe UI", 9), bg="#35164D", fg="#FFFFFF",
                 padx=8, pady=3).pack()

    def _ocultar(self, _e):
        if self.win is not None:
            self.win.destroy()
            self.win = None
