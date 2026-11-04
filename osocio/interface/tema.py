"""Tema de la app: oscuro (negro + violeta) o clásico (los violetas de siempre).

Cómo funciona
-------------
Toda la interfaz está escrita con los colores CLÁSICOS (las constantes de siempre,
`CARD_BG_COLOR`, `#C5A9DF`, etc.). Este módulo se mete en el único punto por donde pasa
cada color camino a Tk (`Misc._options`, más los estilos de ttk) y lo traduce al tema
elegido. Al leerlo de vuelta (`cget`) se devuelve el clásico, así el código que compara
colores sigue funcionando igual en los dos temas.

Así el tema alcanza a TODO lo que se crea —pantallas, diálogos, filas que se arman al
vuelo— sin tener que tocar cada pantalla, y el cambio de tema en vivo es recorrer los
widgets y volver a aplicar sus colores.

Lo que se dibuja con PIL (botones redondeados, tarjetas, toggles) no pasa por Tk: esos
controles llaman a `color()` al pintarse y se registran con `al_cambiar()` para repintarse.
"""
import colorsys
import re
import tkinter as tk
import weakref
from tkinter import ttk

OSCURO = "oscuro"
CLASICO = "clasico"

# Colores propios de piezas nuevas, escritos en "clásico" como todo lo demás.
SIDEBAR_BG = "#34184A"
SIDEBAR_SEP = "#4E2A6B"
MODO_CAJA = "#3F1F58"
ACENTO = "#9B5CD6"            # toggles prendidos, ítem elegido del menú

# Clásico → oscuro, para los colores con un papel claro. El resto de los violetas se
# convierte a gris con la misma luminosidad (ver `_gris`).
_OSCURO = {
    "#5d3c7a": "#0f0f12",   # fondo de la app
    "#4a2666": "#18181c",   # tarjetas
    "#7d4e9f": "#2e2e36",   # bordes de tarjetas (y pestaña activa vieja)
    "#35164d": "#2a2a32",   # botón no elegido: gris que se ve sobre la tarjeta
    "#602d8a": "#7c3aed",   # botón elegido: violeta
    "#4c226e": "#3a3a44",   # hover
    "#2e1146": "#0b0b0e",   # entradas de texto
    "#e6d6f2": "#e4e4e7",   # texto secundario
    "#c5b2d6": "#b4b4bb",
    "#c5a9df": "#a1a1aa",   # ayudas en itálica
    "#3a1d52": "#232328",
    "#100518": "#0a0a0c",
    "#110518": "#0a0a0d",
    "#9662bd": "#6b6b75",
    "#9b86b5": "#9a9aa3",
    "#8b5fb5": "#5e5e68",
    "#8a6e9e": "#8a8a93",
    "#8a6db0": "#85858e",
    "#d8b4fe": "#c4b5fd",
    "#231830": "#141418",   # modales
    "#38234d": "#26262c",
    "#2b1d3a": "#111114",   # consola
    "#3d2a52": "#1f1f24",
    "#1a1a2e": "#0b0b0e",
    "#9c6fb4": "#3a3a44",
    "#110830": "#0b0b0e",
    "#34184a": "#0b0b0e",   # SIDEBAR_BG
    "#4e2a6b": "#222228",   # SIDEBAR_SEP
    "#3f1f58": "#1e1e24",   # MODO_CAJA
    "#9b5cd6": "#7c3aed",   # ACENTO
    "#7e5109": "#e0b25c",   # texto de aviso (marrón sobre amarillo claro en clásico)
    "#fcf3cf": "#2a2418",   # caja de aviso
    "#f4d03f": "#6b5a1f",   # borde de aviso
}

# Entradas de texto en el tema oscuro: los fondos blancos pasan a gris muy oscuro y los
# textos negros a claro (si no, quedarían cajas blancas encandilando sobre el negro).
ENTRADA_OSCURA = "#1c1c22"
TEXTO_ENTRADA_OSCURA = "#e4e4e7"
_CLAVES_FONDO = {"bg", "background", "fieldbackground", "disabledbackground",
                 "readonlybackground"}
_CLAVES_TEXTO = {"fg", "foreground", "insertbackground"}

_HEX = re.compile(r"^#[0-9a-fA-F]{6}$")

_estado = {"tema": OSCURO}
_ida = {}       # clásico (minúsculas) → color del tema
_vuelta = {}    # color del tema (minúsculas) → clásico tal como lo escribió el código
_suscriptos = weakref.WeakSet()
_callbacks = []
_estilos = []   # (style, nombre, kw, es_map) registrados, para re-aplicar al cambiar


def _gris(c):
    """Un violeta cualquiera → gris neutro de luminosidad parecida."""
    r, g, b = (int(c[i:i + 2], 16) / 255 for i in (1, 3, 5))
    h, s, v = colorsys.rgb_to_hsv(r, g, b)
    if not (235 / 360 <= h <= 290 / 360) or s < 0.12:
        return c
    if v < 0.55:              # fondos: más oscuros que en clásico
        v = 0.04 + v * 0.35
    s = 0.06
    r, g, b = colorsys.hsv_to_rgb(250 / 360, s, v)
    return "#%02x%02x%02x" % (round(r * 255), round(g * 255), round(b * 255))


def _calcular(c):
    """El color del tema para un clásico. Único: dos clásicos nunca dan el mismo."""
    base = _OSCURO.get(c) or _gris(c)
    usado = set(_ida.values())
    cand = base
    k = 0
    while cand in usado or (cand != c and cand in _OSCURO):
        k += 1
        n = int(base[1:], 16)
        cand = "#%06x" % min(0xFFFFFF, max(0, n + (k if k % 2 else -k)))
    return cand


def tema():
    return _estado["tema"]


def es_oscuro():
    return _estado["tema"] == OSCURO


def color(c):
    """Un color clásico traducido al tema actual (para lo que se dibuja con PIL)."""
    if not isinstance(c, str) or not _HEX.match(c) or _estado["tema"] == CLASICO:
        return c
    low = c.lower()
    if low not in _ida:
        _ida[low] = _calcular(low)
    dest = _ida[low]
    _vuelta.setdefault(dest, c)
    return dest


def clasico(c):
    """El color clásico detrás de uno ya traducido (lo que devuelve `cget`)."""
    if not isinstance(c, str) or _estado["tema"] == CLASICO:
        return c
    return _vuelta.get(c.lower(), c)


def _luz(c):
    """(saturación, brillo) de un color, o None si no es un color #rrggbb ni un nombre conocido."""
    c = str(c).strip().lower()
    nombres = {"white": "#ffffff", "black": "#000000", "systemwindow": "#ffffff",
               "systemwindowtext": "#000000", "#fff": "#ffffff", "#000": "#000000"}
    c = nombres.get(c, c)
    if not _HEX.match(c):
        return None
    r, g, b = (int(c[i:i + 2], 16) / 255 for i in (1, 3, 5))
    _h, s, v = colorsys.rgb_to_hsv(r, g, b)
    return s, v


def color_opcion(clave, v):
    """Como `color()`, pero sabiendo para qué opción es (fondo blanco → entrada oscura)."""
    if not isinstance(v, str) or _estado["tema"] == CLASICO:
        return v
    clave = str(clave).lstrip("-").lower()
    if clave in _CLAVES_FONDO or clave in _CLAVES_TEXTO:
        sv = _luz(v)
        if sv is not None:
            s, b = sv
            if clave in _CLAVES_FONDO and s < 0.10 and b > 0.92:
                _vuelta.setdefault(ENTRADA_OSCURA, v)
                return ENTRADA_OSCURA
            if clave in _CLAVES_TEXTO and s < 0.15 and b < 0.25:
                _vuelta.setdefault(TEXTO_ENTRADA_OSCURA, v)
                return TEXTO_ENTRADA_OSCURA
    return color(v)


def _mapear_dict(d):
    if not isinstance(d, dict):
        return d
    return {k: color_opcion(k, v) for k, v in d.items()}


_DEFECTOS = {  # opción de la base de Tk: (clásico, oscuro)
    "*Entry.background": ("white", ENTRADA_OSCURA),
    "*Entry.foreground": ("black", TEXTO_ENTRADA_OSCURA),
    "*Entry.insertBackground": ("black", TEXTO_ENTRADA_OSCURA),
    "*Text.background": ("white", ENTRADA_OSCURA),
    "*Text.foreground": ("black", TEXTO_ENTRADA_OSCURA),
    "*Text.insertBackground": ("black", TEXTO_ENTRADA_OSCURA),
    "*Listbox.background": ("white", ENTRADA_OSCURA),
    "*Listbox.foreground": ("black", TEXTO_ENTRADA_OSCURA),
    "*Spinbox.background": ("white", ENTRADA_OSCURA),
    "*Spinbox.foreground": ("black", TEXTO_ENTRADA_OSCURA),
}


def preparar(root):
    """Colores por defecto de las entradas que no eligen color propio. Después de crear root."""
    oscuro = _estado["tema"] == OSCURO
    for patron, (cl, osc) in _DEFECTOS.items():
        root.option_add(patron, osc if oscuro else cl, "widgetDefault")


# --- enganche en Tkinter --------------------------------------------------------------

_instalado = {"ok": False}


def instalar(nombre=OSCURO):
    """Activa el tema. Va ANTES de crear la ventana principal."""
    _estado["tema"] = OSCURO if nombre == OSCURO else CLASICO
    if _instalado["ok"]:
        return
    _instalado["ok"] = True

    orig_options = tk.Misc._options

    def _options(self, cnf, kw=None):
        return orig_options(self, _mapear_dict(cnf), _mapear_dict(kw))
    tk.Misc._options = _options

    orig_cget = tk.Misc.cget

    def cget(self, key):
        v = orig_cget(self, key)
        return clasico(v) if isinstance(v, str) else v
    tk.Misc.cget = cget
    tk.Misc.__getitem__ = cget

    orig_conf = ttk.Style.configure

    def configure(self, style, query_opt=None, **kw):
        if query_opt is None and kw:
            _estilos.append((self, style, dict(kw), False))
        return orig_conf(self, style, query_opt, **_mapear_dict(kw))
    ttk.Style.configure = configure

    orig_map = ttk.Style.map

    def map_(self, style, query_opt=None, **kw):
        if query_opt is None and kw:
            _estilos.append((self, style, dict(kw), True))
        return orig_map(self, style, query_opt, **_mapear_map(kw))
    ttk.Style.map = map_

    orig_tag = ttk.Treeview.tag_configure

    def tag_configure(self, tagname, option=None, **kw):
        return orig_tag(self, tagname, option, **_mapear_dict(kw))
    ttk.Treeview.tag_configure = tag_configure

    _instalado["orig_conf"] = orig_conf
    _instalado["orig_map"] = orig_map


def _mapear_map(kw):
    out = {}
    for k, v in kw.items():
        if isinstance(v, (list, tuple)):
            out[k] = [tuple(x[:-1]) + (color_opcion(k, x[-1]),)
                      if isinstance(x, (list, tuple)) and x else x for x in v]
        else:
            out[k] = v
    return out


# --- cambio de tema en vivo -------------------------------------------------------------

_OPCIONES_COLOR = ("background", "foreground", "activebackground", "activeforeground",
                   "highlightbackground", "highlightcolor", "selectcolor", "insertbackground",
                   "disabledbackground", "disabledforeground", "selectbackground",
                   "selectforeground", "readonlybackground", "troughcolor")
_OPCIONES_ITEM = ("fill", "outline", "activefill", "activeoutline", "disabledfill")


def suscribir(widget):
    """Un control dibujado con PIL: se le llama `_repintar_tema()` al cambiar de tema."""
    _suscriptos.add(widget)


def al_cambiar(fn):
    """Una función a llamar después de cada cambio de tema."""
    _callbacks.append(fn)


def _recorrer(w):
    yield w
    for c in w.winfo_children():
        yield from _recorrer(c)


def cambiar(nombre, root):
    """Pasa a otro tema en vivo: re-aplica los colores de todos los widgets."""
    nuevo = OSCURO if nombre == OSCURO else CLASICO
    if nuevo == _estado["tema"]:
        return
    pendientes = []
    for w in _recorrer(root):
        for op in _OPCIONES_COLOR:
            try:
                v = w.tk.call(w._w, "cget", "-" + op)
            except tk.TclError:
                continue
            v = str(v)
            if v.startswith("#") or v.lower() in ("systemwindow", "systemwindowtext"):
                pendientes.append((w, "configure", None, op, clasico(v)))
        if isinstance(w, tk.Canvas):
            for item in w.find_all():
                for op in _OPCIONES_ITEM:
                    try:
                        v = str(w.itemcget(item, op))
                    except tk.TclError:
                        continue
                    if v.startswith("#"):
                        pendientes.append((w, "itemconfigure", item, op, clasico(v)))
        if isinstance(w, tk.Text):
            for tag in w.tag_names():
                for op in ("background", "foreground"):
                    v = str(w.tag_cget(tag, op))
                    if v.startswith("#"):
                        pendientes.append((w, "tag", tag, op, clasico(v)))
        if isinstance(w, ttk.Treeview):
            for tag in _tags_treeview(w):
                for op in ("background", "foreground"):
                    try:
                        v = str(w.tk.call(w._w, "tag", "configure", tag, "-" + op))
                    except tk.TclError:
                        continue
                    if v.startswith("#"):
                        pendientes.append((w, "tvtag", tag, op, clasico(v)))
    _estado["tema"] = nuevo
    preparar(root)
    for w, tipo, item, op, v in pendientes:
        try:
            if tipo == "configure":
                w.tk.call(w._w, "configure", "-" + op, color_opcion(op, v))
            elif tipo == "itemconfigure":
                w.tk.call(w._w, "itemconfigure", item, "-" + op, color(v))
            elif tipo == "tag":
                w.tk.call(w._w, "tag", "configure", item, "-" + op, color(v))
            else:
                w.tk.call(w._w, "tag", "configure", item, "-" + op, color_opcion(op, v))
        except tk.TclError:
            pass
    for style, nombre_estilo, kw, es_map in list(_estilos):
        try:
            if es_map:
                _instalado["orig_map"](style, nombre_estilo, **_mapear_map(kw))
            else:
                _instalado["orig_conf"](style, nombre_estilo, **_mapear_dict(kw))
        except tk.TclError:
            pass
    for w in list(_suscriptos):
        try:
            w._repintar_tema()
        except tk.TclError:
            pass
    for fn in list(_callbacks):
        try:
            fn()
        except Exception:
            pass


def _tags_treeview(tv):
    try:
        return tv.tk.splitlist(tv.tk.call(tv._w, "tag", "names"))
    except tk.TclError:
        return ()
