# -*- coding: utf-8 -*-
"""
interface_demo.py — Demo visual interactiva del nuevo diseño Figma-style.
Esta versión incluye scrollbars horizontales, scroll general de ventana, colores pastel suaves y advertencias dinámicas.
"""
import tkinter as tk
from tkinter import ttk, messagebox
import os
import sys

# Evitar UnicodeEncodeError cuando el backend hace print() con emojis bajo consola cp1252
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import time
import threading
from datetime import datetime as _dt, date as _date
from PIL import Image, ImageTk

# === Backend real (coexiste con la app vieja; import defensivo) ===
_APP_BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
for _p in (_APP_BASE, os.path.join(_APP_BASE, "forms"), os.path.join(_APP_BASE, "core")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

BACKEND_OK = True
_BACKEND_IMPORT_ERROR = ""
try:
    from utils.paths import DATA_DIR, JSON_DIR, BASE_DIR
    from utils.fixed_field_mapping_store import build_excel_columns_for_country
    from utils.data_generator import generar_fila_datos
    from core.country_configs import COUNTRY_CONFIGS
    from interface.helpers_interface import (
        cargar_config_global, guardar_config_global, obtener_email_destinatario,
    )
    from utils.scheduling import guardar_programacion, cargar_programacion, limpiar_programacion
    from interface.field_validation_ui import build_field_validation_tab
    from interface.dealer_comparator_ui import build_dealer_comparator_tab
    from interface.ids_dinamicos_ui import abrir_popup_ids_dinamicos
except Exception as _imp_err:  # noqa: BLE001
    BACKEND_OK = False
    _BACKEND_IMPORT_ERROR = str(_imp_err)
    build_field_validation_tab = None
    build_dealer_comparator_tab = None
    abrir_popup_ids_dinamicos = None
    BASE_DIR = _APP_BASE
    DATA_DIR = os.path.join(_APP_BASE, "data")
    JSON_DIR = os.path.join(_APP_BASE, "json")
    COUNTRY_CONFIGS = {}

    def build_excel_columns_for_country(pais):
        return ["URL", "Formulario", "Modelo", "Nombre", "Apellido", "Documento", "Celular",
                "Email", "Region", "Ciudad", "Concesionario", "Fecha de compra", "Comentario"]

    def generar_fila_datos(pais, device=None, doc_types=None):
        return {}

    def cargar_config_global():
        return {}

    def guardar_config_global(cfg):
        return False

    def obtener_email_destinatario():
        return []

    def guardar_programacion(p):
        return False

    def cargar_programacion():
        return None

    def limpiar_programacion():
        return False


def _excel_path_for(pais):
    """Excel principal (editable / previsualización / scheduler) del país: el del
    primer dispositivo disponible (Chrome→Firefox→Edge→Mac→Android) o el genérico.
    Si ninguno existe, devuelve la ruta de Chrome (para creación por defecto)."""
    for suf in ("Chrome", "Firefox", "Edge", "Mac", "Android", "Generico"):
        p = os.path.join(DATA_DIR, f"Lead_information_Formulario_{pais}_{suf}.xlsx")
        if os.path.exists(p):
            return p
    return os.path.join(DATA_DIR, f"Lead_information_Formulario_{pais}_Chrome.xlsx")


def _t3_tag(t3):
    """Sufijo de nombre para Excels de formularios T3 2.0 (Adobe AEM)."""
    return "_T3" if t3 else ""


def _lead_excel_name(pais, suffix, t3=False):
    """Nombre de archivo Excel de leads: …_{Pais}_{Dispositivo}[_T3].xlsx."""
    return f"Lead_information_Formulario_{pais}_{suffix}{_t3_tag(t3)}.xlsx"


def _generic_excel_path_for(pais, t3=False):
    """Excel único 'compartido' del país: mismos datos para todos los dispositivos."""
    return os.path.join(DATA_DIR, _lead_excel_name(pais, "Generico", t3))


def _device_excel_suffix(dev):
    """Sufijo de archivo Excel por dispositivo. LT usa el mismo nombre que leen
    los controllers de LambdaTest (…_Mac.xlsx / …_Android.xlsx)."""
    d = str(dev).strip().lower()
    if d in ("mac lt", "mac"):
        return "Mac"
    if d in ("android lt", "android"):
        return "Android"
    return str(dev).strip().capitalize()  # Chrome / Firefox / Edge


# Serializa la reserva del nº de resultados (setup_directories_and_files) para que
# sesiones en paralelo NO escriban el mismo resultados_<Pais>N.xlsx (choque de carpetas).
_SETUP_LOCK = threading.Lock()
_SETUP_PATCHED = [False]


def _ensure_serialized_setup():
    if _SETUP_PATCHED[0]:
        return
    try:
        from core import base_form_filler as _bff
        _orig_setup = _bff.BaseFormFiller.setup_directories_and_files

        def _locked_setup(self, _orig=_orig_setup):
            with _SETUP_LOCK:
                return _orig(self)

        _bff.BaseFormFiller.setup_directories_and_files = _locked_setup
        _SETUP_PATCHED[0] = True
    except Exception:
        pass


# País ↔ abreviatura (para la pestaña Validación de Campos)
_VAL_ABBR = {"Argentina": "AR", "Bolivia": "BO", "Brasil": "BR", "Chile": "CH", "Colombia": "CO",
             "Ecuador": "EC", "Paraguay": "PA", "Peru": "PE", "Uruguay": "UY"}
_VAL_FULL = {v: k for k, v in _VAL_ABBR.items()}
_VAL_FULL.update({"CL": "Chile", "PY": "Paraguay"})  # alias comunes

# === PALETA DE COLORES PREMIUM AMIGABLE PARA LOS OJOS ===
APP_BG_COLOR = "#5D3C7A"          # Morado base de Osocio (Ventana principal)
CARD_BG_COLOR = "#4A2666"         # Morado oscuro/rico para las tarjetas contenedoras (Cards)
BORDER_COLOR = "#7D4E9F"          # Borde de contraste de paneles
TEXT_PRIMARY = "#FFFFFF"          # Texto principal blanco
TEXT_SECONDARY = "#E6D6F2"        # Texto secundario lavanda muy claro

# Botones de selección y hovers en tonos pastel súper suaves (para no dañar la vista)
ACCENT_COLOR = "#AED6F1"          # Azul pastel muy suave para elementos activos y chequeados
ACCENT_MUTED = "#D4E6F1"          # Azul pastel aún más claro para hovers secundarios
BUTTON_INACTIVE = "#35164D"       # Botón inactivo morado oscuro
BUTTON_ACTIVE = "#602D8A"         # Botón activo/seleccionado (morado medio suave)
BUTTON_HOVER = "#4C226E"          # Hover morado sutil (muy oscuro)
ENTRY_BG = "#2E1146"              # Fondo para campos de entrada

# Colores dedicados para pestañas (Tabs)
TAB_ACTIVE_BG = "#7D4E9F"         # Morado brillante de contraste para pestaña activa
TAB_ACTIVE_FG = "#FFFFFF"         # Texto blanco para pestaña activa
TAB_INACTIVE_BG = "#35164D"       # Morado oscuro para pestaña inactiva
TAB_INACTIVE_FG = "#C5B2D6"       # Texto secundario lavanda claro para pestaña inactiva

# Colores suaves para los textos de los botones de acción (fondos oscuros)
TEXT_ADD = "#82E0AA"              # Texto verde pastel suave para Agregar
TEXT_DELETE = "#F1948A"           # Texto rojo/coral pastel suave para Eliminar
TEXT_SAVE = "#85C1E9"             # Texto azul pastel suave para Guardar
TEXT_EXCEL = "#F8C471"            # Texto naranja pastel suave para Abrir Excel

# Botones de acción principal en colores pasteles apagados (Fondos súper oscuros para evitar fatiga)
EXECUTE_BG = "#27AE60"            # Verde esmeralda vívido para el CTA Ejecutar
EXECUTE_FG = "#FFFFFF"            # Texto blanco para máximo contraste
EXECUTE_HOVER = "#2ECC71"          # Hover verde más brillante

VALIDATE_BG = "#1D5270"           # Azul acero oscuro para Validación
VALIDATE_FG = "#85C1E9"           # Celeste pastel para el texto
VALIDATE_HOVER = "#28729C"         # Hover azul acero

# CTA "Guardar" — mismo estilo en toda la app (azul medio, no tan pastel)
SAVE_BG = "#4F86C6"
SAVE_FG = "#FFFFFF"
SAVE_HOVER = "#5E95D3"

# Colores de la caja de advertencia (Muted Yellow)
WARN_BG = "#FCF3CF"
WARN_BORDER = "#F4D03F"
WARN_TEXT = "#7E5109"

# === Diálogo de Programación Semanal (portado de weekly_scheduler.py, recoloreado) ===
_S_BG = APP_BG_COLOR
_S_CARD = CARD_BG_COLOR
_S_SEL = BUTTON_ACTIVE
_S_UNSEL = BUTTON_INACTIVE
_S_MUTED = TEXT_SECONDARY
_S_BORDER = BORDER_COLOR
_S_ACCENT = ACCENT_COLOR
_S_AMBER = "#F8C471"
_S_GREEN = "#82E0AA"
_S_RED = "#F1948A"
_S_WHITE = "#FFFFFF"

_SCH_DAYS = [("Lun", "Lunes"), ("Mar", "Martes"), ("Mié", "Miércoles"), ("Jue", "Jueves"),
             ("Vie", "Viernes"), ("Sáb", "Sábado"), ("Dom", "Domingo")]
_SCH_HOURS = [f"{h:02d}:{m:02d}" for h in range(24) for m in (0, 15, 30, 45)]
_SCH_COUNTRIES = ["Argentina", "Bolivia", "Brasil", "Chile", "Colombia", "Ecuador", "Paraguay", "Peru", "Uruguay"]


class _DemoSchedulerPanel(tk.Frame):
    """Panel de configuración semanal (días, horarios, países) integrado directamente en la pestaña."""

    def __init__(self, parent, on_save, initial_config=None, on_close=None):
        super().__init__(parent, bg=_S_BG, bd=1, highlightthickness=1, highlightbackground=_S_BORDER)

        self._on_save = on_save
        self._on_close = on_close
        self._selected_day = None
        self._copy_open_state = False
        self._copy_selected_days = []
        self._hour_btns = {}
        self._badges_frame = None
        self._copy_frame_inner = None
        self._val_lbl = None
        self._save_btn = None
        self._edit_all_days = False
        self._custom_hour_var = tk.StringVar()

        cfg = initial_config or {}
        self._schedule = {k: list(v) for k, v in cfg.get("horarios", {}).items()}
        self._countries = list(cfg.get("paises", []))

        self._build_ui()

    def _close(self):
        if self._on_close:
            self._on_close()

    def _build_ui(self):
        header = tk.Frame(self, bg=_S_CARD)
        header.pack(fill="x")
        tk.Label(header, text="⚙  Configurar automatización", font=("Segoe UI", 11, "bold"),
                 bg=_S_CARD, fg=_S_WHITE).pack(side="left", padx=16, pady=10)
        
        btn_close = tk.Button(header, text="✕", font=("Segoe UI", 11), bg=_S_CARD, fg=_S_MUTED, relief="flat",
                              cursor="hand2", activebackground=_S_BG, activeforeground=_S_WHITE, bd=0,
                              command=self._close)
        btn_close.pack(side="right", padx=12, pady=6)
        btn_close.bind("<Enter>", lambda e: btn_close.config(bg="#E74C3C", fg="white"))
        btn_close.bind("<Leave>", lambda e: btn_close.config(bg=_S_CARD, fg=_S_MUTED))
        
        tk.Frame(self, bg=_S_BORDER, height=1).pack(fill="x")

        self._build_footer(self, dict(padx=14, pady=10))
        tk.Frame(self, bg=_S_BORDER, height=1).pack(fill="x", side="bottom")

        body = tk.Frame(self, bg=_S_BG)
        body.pack(fill="both", expand=True)

        p = dict(padx=14, pady=8)
        self._build_days_section(body, p)
        self._build_countries_section(body, p)

    def _card_frame(self, parent):
        return tk.Frame(parent, bg=_S_CARD, highlightbackground=_S_BORDER, highlightthickness=1)

    def _build_days_section(self, parent, pad):
        card = self._card_frame(parent)
        card.pack(fill="x", **pad)

        hrow = tk.Frame(card, bg=_S_CARD)
        hrow.pack(fill="x", padx=12, pady=(12, 6))
        tk.Label(hrow, text="DÍAS DE LA SEMANA", font=("Segoe UI", 9, "bold"), bg=_S_CARD, fg=_S_MUTED).pack(side="left")
        self._total_badge = tk.Label(hrow, text="", font=("Segoe UI", 8, "bold"), bg=_S_BG, fg=_S_ACCENT, padx=8, pady=2)
        self._total_badge.pack(side="left", padx=6)
        self._clear_all_lbl = tk.Label(hrow, text="Desmarcar todos", font=("Segoe UI", 9, "underline"),
                                       bg=_S_CARD, fg=_S_ACCENT, cursor="hand2")
        self._clear_all_lbl.pack(side="right")
        self._clear_all_lbl.bind("<Button-1>", lambda e: self._clear_all())

        day_row = tk.Frame(card, bg=_S_CARD)
        day_row.pack(fill="x", padx=12, pady=(4, 8))
        self._day_btns = {}
        for short, full in _SCH_DAYS:
            col_f = tk.Frame(day_row, bg=_S_CARD)
            col_f.pack(side="left", expand=True, fill="x", padx=2)
            btn = tk.Button(col_f, text=f"{short}\n—", font=("Segoe UI", 8, "bold"), relief="flat",
                            cursor="hand2", pady=8, wraplength=60)
            btn.pack(fill="x")
            btn.config(command=lambda d=full: self._select_day(d))
            self._day_btns[full] = btn

        tk.Label(card, text="① Tocá un día para abrirlo.   ② Agregá u elegí sus horarios.   ③ Copialos a otros días con \"Todos\" o eligiendo días.",
                 font=("Segoe UI", 8, "italic"), bg=_S_CARD, fg=_S_MUTED, justify="left").pack(anchor="w", padx=12, pady=(0, 6))

        self._hours_outer = tk.Frame(card, bg=_S_BG)
        self._hours_outer.pack(fill="x", padx=12, pady=(0, 8))
        self._hours_outer.pack_forget()
        self._update_day_buttons()

    def _select_day(self, day):
        if self._selected_day == day:
            self._close_hours_panel()
            return
        self._selected_day = day
        self._hours_outer.pack(fill="x", padx=12, pady=(0, 8))
        self._build_hours_panel()
        self._update_day_buttons()

    def _close_hours_panel(self):
        self._selected_day = None
        self._hours_outer.pack_forget()
        self._copy_open_state = False
        self._copy_selected_days = []
        self._update_day_buttons()

    def _build_hours_panel(self):
        for w in self._hours_outer.winfo_children():
            w.destroy()

        self._apply_days = set()

        # Header + Guardar
        h_hdr = tk.Frame(self._hours_outer, bg=_S_BG)
        h_hdr.pack(fill="x", pady=(8, 4))
        tk.Label(h_hdr, text=f"🕐  Horarios para el {self._selected_day}", font=("Segoe UI", 10, "bold"),
                 bg=_S_BG, fg=_S_WHITE).pack(side="left")
        tk.Button(h_hdr, text="✓ Guardar", font=("Segoe UI", 8, "bold"), bg=SAVE_BG, fg=SAVE_FG, relief="flat",
                  activebackground=SAVE_HOVER, activeforeground=SAVE_FG,
                  cursor="hand2", padx=10, pady=3, command=self._close_hours_panel).pack(side="right")

        # 1) Agregar horario personalizado
        cust_row = tk.Frame(self._hours_outer, bg=_S_BG)
        cust_row.pack(fill="x", pady=(2, 6))
        tk.Label(cust_row, text="1)  Agregar horario:", font=("Segoe UI", 9, "bold"), bg=_S_BG, fg=_S_WHITE).pack(side="left")
        cust_entry = tk.Entry(cust_row, textvariable=self._custom_hour_var, font=("Segoe UI", 10), width=8,
                              bg=_S_UNSEL, fg=_S_WHITE, insertbackground=_S_WHITE, relief="flat")
        cust_entry.pack(side="left", padx=(6, 2), ipady=2)
        cust_entry.bind("<Return>", lambda e: self._add_custom_hour(self._custom_hour_var.get()))
        tk.Label(cust_row, text="HH:MM", font=("Segoe UI", 8), bg=_S_BG, fg=_S_MUTED).pack(side="left", padx=(2, 6))
        tk.Button(cust_row, text="+ Agregar", font=("Segoe UI", 8, "bold"), bg=_S_SEL, fg=_S_WHITE, relief="flat",
                  cursor="hand2", padx=10, pady=3,
                  command=lambda: self._add_custom_hour(self._custom_hour_var.get())).pack(side="left")

        # Horarios elegidos (badges)
        tk.Label(self._hours_outer, text="Horarios elegidos (tocá para quitar):", font=("Segoe UI", 8, "bold"),
                 bg=_S_BG, fg=_S_MUTED).pack(anchor="w", pady=(4, 2))
        self._badges_frame = tk.Frame(self._hours_outer, bg=_S_BG)
        self._badges_frame.pack(fill="x")
        self._refresh_badges()

        # 2) Aplicar estos horarios a otros días
        tk.Frame(self._hours_outer, bg=_S_BORDER, height=1).pack(fill="x", pady=(6, 6))
        tk.Label(self._hours_outer, text="2)  Aplicar estos horarios a otros días (se copia al instante):", font=("Segoe UI", 9, "bold"),
                 bg=_S_BG, fg=_S_WHITE).pack(anchor="w")

        chips_row = tk.Frame(self._hours_outer, bg=_S_BG)
        chips_row.pack(fill="x", pady=4)
        self._apply_chip_btns = {}

        def _apply_day(d):
            src = list(self._schedule.get(self._selected_day, []))
            if src:
                self._schedule[d] = list(src)
            else:
                self._schedule.pop(d, None)

        def _toggle_apply(d):
            if d in self._apply_days:
                self._apply_days.discard(d)
                self._schedule.pop(d, None)
            else:
                self._apply_days.add(d)
                _apply_day(d)
            _refresh_chips()
            self._update_day_buttons()
            self._update_footer()

        def _all_other():
            others = [d for _, d in _SCH_DAYS if d != self._selected_day]
            if self._apply_days == set(others):
                for d in others:
                    self._schedule.pop(d, None)
                self._apply_days = set()
            else:
                self._apply_days = set(others)
                for d in others:
                    _apply_day(d)
            _refresh_chips()
            self._update_day_buttons()
            self._update_footer()

        def _refresh_chips():
            for d, b in self._apply_chip_btns.items():
                on = d in self._apply_days
                b.config(bg=_S_SEL if on else _S_UNSEL, fg=_S_WHITE if on else _S_MUTED,
                         relief="raised" if on else "flat", bd=2 if on else 0)
            all_on = bool(self._apply_days) and self._apply_days == set(d for _, d in _SCH_DAYS if d != self._selected_day)
            todos_btn.config(relief="raised" if all_on else "flat", bd=2 if all_on else 0)

        todos_btn = tk.Button(chips_row, text="Todos", font=("Segoe UI", 8, "bold"), bg=_S_AMBER, fg=_S_BG, relief="flat",
                              cursor="hand2", padx=8, pady=4, command=_all_other)
        todos_btn.pack(side="left", padx=(0, 8))
        for _, d in _SCH_DAYS:
            if d == self._selected_day:
                continue
            b = tk.Button(chips_row, text=d[:3], font=("Segoe UI", 8, "bold"), bg=_S_UNSEL, fg=_S_MUTED, relief="flat",
                          cursor="hand2", padx=8, pady=4, command=lambda dd=d: _toggle_apply(dd))
            b.pack(side="left", padx=2)
            self._apply_chip_btns[d] = b

        # 3) …o elegí de la grilla (cada 15 min)
        tk.Frame(self._hours_outer, bg=_S_BORDER, height=1).pack(fill="x", pady=(8, 6))
        tk.Label(self._hours_outer, text="…o elegí de la grilla (cada 15 min):", font=("Segoe UI", 8, "italic"),
                 bg=_S_BG, fg=_S_MUTED).pack(anchor="w", pady=(0, 2))
        grid = tk.Frame(self._hours_outer, bg=_S_BG)
        grid.pack(fill="x", pady=(0, 4))
        current = self._schedule.get(self._selected_day, [])
        self._hour_btns = {}
        for c in range(6):
            grid.columnconfigure(c, weight=1)

        for idx, hour in enumerate(_SCH_HOURS):
            r, c = divmod(idx, 6)
            picked = hour in current
            btn = tk.Button(grid, text=hour, font=("Segoe UI", 9),
                            bg=_S_SEL if picked else _S_UNSEL, fg=_S_WHITE if picked else _S_MUTED,
                            relief="raised" if picked else "flat", cursor="hand2", padx=4, pady=6, width=5)
            btn.grid(row=r, column=c, padx=2, pady=2, sticky="ew")
            btn.config(command=lambda h=hour: self._toggle_hour(h))
            self._hour_btns[hour] = btn

    def _refresh_badges(self):
        if not self._badges_frame:
            return
        for w in self._badges_frame.winfo_children():
            w.destroy()
        selected = sorted(self._schedule.get(self._selected_day, []))
        MAX_SHOW = 12
        for h in selected[:MAX_SHOW]:
            is_custom = h not in _SCH_HOURS
            bg = "#3A1A55" if is_custom else _S_BG
            fg = _S_AMBER if is_custom else _S_ACCENT
            tk.Button(self._badges_frame, text=f"✕ {h}", font=("Segoe UI", 9, "bold"), bg=bg, fg=fg,
                      padx=6, pady=3, relief="solid", bd=1, cursor="hand2",
                      command=lambda x=h: self._toggle_hour(x)).pack(side="left", padx=2)
        if len(selected) > MAX_SHOW:
            tk.Label(self._badges_frame, text=f"+{len(selected) - MAX_SHOW} más", font=("Segoe UI", 9),
                     bg=_S_BG, fg=_S_MUTED).pack(side="left", padx=4)

    def _toggle_hour(self, hour):
        day = self._selected_day
        dh = self._schedule.setdefault(day, [])
        if hour in dh:
            dh.remove(hour)
        else:
            dh.append(hour)
            dh.sort()
        if not self._schedule.get(day):
            self._schedule.pop(day, None)
        picked = hour in self._schedule.get(day, [])
        btn = self._hour_btns.get(hour)
        if btn:
            btn.config(bg=_S_SEL if picked else _S_UNSEL, fg=_S_WHITE if picked else _S_MUTED,
                       relief="raised" if picked else "flat")
        self._refresh_badges()
        self._update_day_buttons()
        self._update_footer()

    def _add_custom_hour(self, time_str):
        time_str = time_str.strip()
        try:
            if ':' not in time_str:
                raise ValueError
            h, m = time_str.split(':', 1)
            h, m = int(h), int(m)
            if not (0 <= h <= 23 and 0 <= m <= 59):
                raise ValueError
            hora = f"{h:02d}:{m:02d}"
        except (ValueError, TypeError):
            messagebox.showwarning("Horario inválido", "Ingresá un horario válido en formato HH:MM\n(ej: 09:33, 14:05)", parent=self)
            return
        day = self._selected_day
        if not day:
            return
        dh = self._schedule.setdefault(day, [])
        if hora not in dh:
            dh.append(hora)
            dh.sort()
        self._custom_hour_var.set("")
        gbtn = self._hour_btns.get(hora)
        if gbtn:
            gbtn.config(bg=_S_SEL, fg=_S_WHITE, relief="raised")
        self._refresh_badges()
        self._update_day_buttons()
        self._update_footer()

    def _apply_hours_to_days(self, days):
        src = list(self._schedule.get(self._selected_day, []))
        if not src or not days:
            return
        for d in days:
            self._schedule[d] = list(src)
        self._update_day_buttons()
        self._update_footer()
        messagebox.showinfo("Aplicado", f"Horarios copiados a {len(days)} día(s): {', '.join(sorted(days))}.", parent=self)

    def _update_day_buttons(self):
        total = sum(len(v) for v in self._schedule.values())
        if total > 0:
            self._total_badge.config(text=f"{total} horario{'s' if total != 1 else ''}")
            self._total_badge.pack(side="left", padx=6)
            self._clear_all_lbl.pack(side="right")
        else:
            self._total_badge.pack_forget()
            self._clear_all_lbl.pack_forget()

        for short, full in _SCH_DAYS:
            count = len(self._schedule.get(full, []))
            is_open = self._selected_day == full
            btn = self._day_btns[full]
            if is_open:
                btn.config(bg=_S_SEL, fg=_S_WHITE, text=f"{short}\n● abierto")
            elif count > 0:
                btn.config(bg=_S_UNSEL, fg=_S_ACCENT, text=f"{short}\n{count} sel.")
            else:
                btn.config(bg=_S_UNSEL, fg=_S_MUTED, text=f"{short}\n—")

    def _clear_all(self):
        self._schedule = {}
        self._close_hours_panel()
        self._update_day_buttons()
        self._update_footer()

    def _build_countries_section(self, parent, pad):
        card = self._card_frame(parent)
        card.pack(fill="x", **pad)

        hrow = tk.Frame(card, bg=_S_CARD)
        hrow.pack(fill="x", padx=12, pady=(12, 6))
        tk.Label(hrow, text="🌎  PAÍSES A TESTEAR", font=("Segoe UI", 9, "bold"), bg=_S_CARD, fg=_S_MUTED).pack(side="left")
        self._toggle_all_lbl = tk.Label(hrow, text="Seleccionar todos", font=("Segoe UI", 9, "underline"),
                                        bg=_S_CARD, fg=_S_ACCENT, cursor="hand2")
        self._toggle_all_lbl.pack(side="right")
        self._toggle_all_lbl.bind("<Button-1>", lambda e: self._toggle_all())

        grid = tk.Frame(card, bg=_S_CARD)
        grid.pack(fill="x", padx=12, pady=(0, 4))
        for c in range(3):
            grid.columnconfigure(c, weight=1)

        self._country_vars = {}
        self._country_items = {}
        for idx, country in enumerate(_SCH_COUNTRIES):
            r, c = divmod(idx, 3)
            checked = country in self._countries
            var = tk.BooleanVar(value=checked)
            self._country_vars[country] = var
            bg = _S_SEL if checked else _S_UNSEL
            border = _S_ACCENT if checked else _S_BORDER
            item = tk.Frame(grid, bg=bg, highlightbackground=border, highlightthickness=1)
            item.grid(row=r, column=c, padx=3, pady=3, sticky="ew")
            self._country_items[country] = item
            cb = tk.Checkbutton(item, text=country, variable=var, font=("Segoe UI", 9), bg=bg,
                                fg=_S_WHITE, activebackground=_S_SEL, activeforeground=_S_WHITE,
                                selectcolor=_S_UNSEL, padx=8, pady=6)
            cb.pack(fill="x")
            var.trace("w", lambda *a, co=country, it=item, cb_=cb, v=var: self._on_country_toggle(co, it, cb_, v))

        self._count_lbl = tk.Label(card, text="", font=("Segoe UI", 8), bg=_S_CARD, fg=_S_MUTED)
        self._count_lbl.pack(anchor="w", padx=12, pady=(2, 10))
        self._update_country_count()

    def _on_country_toggle(self, country, item, cb, var):
        checked = var.get()
        bg = _S_SEL if checked else _S_UNSEL
        border = _S_ACCENT if checked else _S_BORDER
        item.config(bg=bg, highlightbackground=border)
        cb.config(bg=bg, fg=_S_WHITE)
        if checked and country not in self._countries:
            self._countries.append(country)
        elif not checked and country in self._countries:
            self._countries.remove(country)
        self._update_country_count()
        self._update_toggle_all_label()
        self._update_footer()

    def _toggle_all(self):
        all_sel = len(self._countries) == len(_SCH_COUNTRIES)
        for var in self._country_vars.values():
            var.set(not all_sel)

    def _update_toggle_all_label(self):
        self._toggle_all_lbl.config(text="Desmarcar todos" if len(self._countries) == len(_SCH_COUNTRIES) else "Seleccionar todos")

    def _update_country_count(self):
        n = len(self._countries)
        self._count_lbl.config(text=f"{n} de {len(_SCH_COUNTRIES)} países seleccionados" if n > 0 else "")

    def _build_footer(self, parent, pad):
        footer = tk.Frame(parent, bg=_S_BG)
        footer.pack(fill="x", side="bottom", **pad)
        self._val_lbl = tk.Label(footer, text="", font=("Segoe UI", 9), bg=_S_BG, fg=_S_AMBER,
                                 padx=10, pady=6, justify="left", anchor="w", wraplength=500, relief="flat")
        btn_row = tk.Frame(footer, bg=_S_BG)
        btn_row.pack(fill="x")
        self._save_btn = tk.Button(btn_row, text="💾  Guardar configuración", font=("Segoe UI", 10, "bold"),
                                   bg=_S_SEL, fg=_S_WHITE, relief="flat", cursor="hand2", padx=16, pady=8,
                                   command=self._save)
        self._save_btn.pack(side="right")
        self._update_footer()

    def _update_footer(self):
        if self._val_lbl is None:
            return
        total = sum(len(v) for v in self._schedule.values())
        n = len(self._countries)
        can = total > 0 and n > 0
        if not can:
            if total == 0 and n == 0:
                msg = "⚠  Seleccioná al menos un horario y un país para continuar."
            elif total == 0:
                msg = "⚠  Seleccioná al menos un horario para continuar."
            else:
                msg = "⚠  Seleccioná al menos un país para continuar."
            self._val_lbl.config(text=msg)
            self._val_lbl.pack(fill="x", pady=(0, 8))
        else:
            self._val_lbl.pack_forget()
        self._save_btn.config(state="normal" if can else "disabled", bg=_S_SEL if can else _S_UNSEL,
                              cursor="hand2" if can else "arrow")

    def _save(self):
        total = sum(len(v) for v in self._schedule.values())
        if total == 0 or not self._countries:
            return
        config = {"horarios": {k: v for k, v in self._schedule.items() if v}, "paises": list(self._countries)}
        if self._on_save(config) is False:
            return
        self._close()


# Directorio de Assets
from utils.paths import BASE_DIR, ASSET_DIR

ICONS_CACHE = {}

def get_button_icon(name):
    if not name:
        return None
    if name not in ICONS_CACHE:
        icon_path = os.path.join(ASSET_DIR, "tabler_icons", name)
        if os.path.exists(icon_path):
            try:
                img = Image.open(icon_path)
                ICONS_CACHE[name] = ImageTk.PhotoImage(img)
            except Exception:
                ICONS_CACHE[name] = None
        else:
            ICONS_CACHE[name] = None
    return ICONS_CACHE[name]

_tray_instance = None

if os.name == 'nt':
    try:
        import queue
        import threading
        import win32gui
        import win32con
        import win32api

        # OJO — por qué el tray corre en su propio hilo:
        #
        # Antes esto era un WNDPROC de ctypes (ctypes.WINFUNCTYPE) colgado del hilo de Tk.
        # El mainloop de Tk bombea los mensajes de Windows CON EL GIL LIBERADO, así que cuando
        # Windows invocaba el WNDPROC de Python el intérprete se caía en seco:
        #   Fatal Python error: PyEval_RestoreThread ... the GIL is released
        # Por eso el icono se veía (Shell_NotifyIcon no necesita callback) pero el click derecho
        # mostraba un recuadro blanco y mataba el proceso.
        # (Además DefWindowProcW se llamaba sin argtypes → "OverflowError: int too long to
        # convert" en cada mensaje, en Windows de 64 bits.)
        #
        # Solución: la ventana del icono vive en un hilo propio con su propio PumpMessages().
        # Ahí pywin32 sí adquiere el GIL bien. Los clicks no tocan Tk directamente: se encolan
        # y el hilo de Tk los consume con after() (Tk no es thread-safe).
        _TRAY_MSG = win32con.WM_USER + 20

        class SysTrayIcon:
            def __init__(self, icon_path, hover_text, on_quit, on_double_click=None,
                         on_right_click=None):
                self.icon_path = icon_path
                self.hover_text = hover_text
                self.on_quit = on_quit
                self.on_double_click = on_double_click
                # El menú del click derecho lo dibuja Tk (ver _menu_tray). TrackPopupMenu no
                # sirve acá: la ventana propietaria del icono es invisible y de 0x0,
                # SetForegroundWindow sobre ella falla y Windows 11 dibuja el menú vacío.
                self.on_right_click = on_right_click

                self.hwnd = None
                self.notify_id = None
                self.events = queue.Queue()      # ("restore" | "menu", x, y)
                self._listo = threading.Event()

                self._hilo = threading.Thread(target=self._correr, daemon=True)
                self._hilo.start()
                self._listo.wait(timeout=5)

            # ── hilo del icono ────────────────────────────────────────────────
            def _correr(self):
                hinst = win32api.GetModuleHandle(None)
                wc = win32gui.WNDCLASS()
                wc.hInstance = hinst
                # Nombre de clase ÚNICO por instancia: en pywin32 el WNDPROC va atado a la
                # CLASE, no a la ventana. Si se reusa la clase de un tray anterior (al
                # minimizar → restaurar → minimizar de nuevo), los mensajes del icono nuevo
                # caen en los handlers del objeto viejo — cuyo hilo ya murió y cuya cola no
                # lee nadie: el icono aparecía pero no respondía ni al click derecho ni al
                # doble click.
                self._clase_nombre = f"OsocioTrayIcon_{id(self)}"
                wc.lpszClassName = self._clase_nombre
                wc.lpfnWndProc = {
                    win32con.WM_DESTROY: self._on_destroy,
                    win32con.WM_CLOSE: self._on_close_msg,
                    _TRAY_MSG: self._on_tray_msg,
                }
                self._clase_atom = win32gui.RegisterClass(wc)

                self.hwnd = win32gui.CreateWindow(
                    self._clase_atom, "Osocio Tray", win32con.WS_OVERLAPPED,
                    0, 0, 0, 0, 0, 0, hinst, None
                )

                if os.path.exists(self.icon_path):
                    self.hicon = win32gui.LoadImage(
                        hinst, self.icon_path, win32con.IMAGE_ICON, 0, 0,
                        win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
                    )
                else:
                    self.hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)

                self.show_icon()
                self._listo.set()
                win32gui.PumpMessages()          # loop propio, no el de Tk

            def _on_close_msg(self, hwnd, msg, wparam, lparam):
                # DestroyWindow sólo se puede llamar desde el hilo dueño de la ventana:
                # destroy() postea WM_CLOSE y la destrucción real pasa acá.
                win32gui.DestroyWindow(hwnd)
                return 0

            def _on_destroy(self, hwnd, msg, wparam, lparam):
                self.remove_icon()
                win32gui.PostQuitMessage(0)      # corta el PumpMessages de este hilo
                try:
                    # Liberar la clase, si no queda una registrada por cada minimizado
                    win32gui.UnregisterClass(self._clase_atom,
                                             win32api.GetModuleHandle(None))
                except Exception:
                    pass
                return 0

            def _on_tray_msg(self, hwnd, msg, wparam, lparam):
                # Solo encolar: tocar Tk desde este hilo lo rompe.
                if lparam in (win32con.WM_LBUTTONDBLCLK, win32con.WM_LBUTTONUP):
                    self.events.put(("restore", 0, 0))
                elif lparam == win32con.WM_RBUTTONUP:
                    x, y = win32gui.GetCursorPos()
                    self.events.put(("menu", x, y))
                return 0

            # ── API desde el hilo de Tk ───────────────────────────────────────
            def procesar_eventos(self):
                """La llama el hilo de Tk periódicamente (after)."""
                while True:
                    try:
                        accion, x, y = self.events.get_nowait()
                    except queue.Empty:
                        return
                    if accion == "restore" and self.on_double_click:
                        self.on_double_click()
                    elif accion == "menu" and self.on_right_click:
                        self.on_right_click(x, y)

            def show_icon(self):
                flags = win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP
                nid = (self.hwnd, 0, flags, _TRAY_MSG, self.hicon, self.hover_text)
                if self.notify_id is None:
                    win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid)
                else:
                    win32gui.Shell_NotifyIcon(win32gui.NIM_MODIFY, nid)
                self.notify_id = nid

            def remove_icon(self):
                if self.notify_id:
                    try:
                        win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, self.notify_id)
                    except Exception:
                        pass
                    self.notify_id = None

            def destroy(self):
                """Saca el icono y baja el hilo del tray."""
                self.remove_icon()
                if self.hwnd:
                    try:
                        # WM_CLOSE (no WM_DESTROY): postear WM_DESTROY llama al handler pero
                        # no destruye la ventana. El hilo dueño hace el DestroyWindow real.
                        win32gui.PostMessage(self.hwnd, win32con.WM_CLOSE, 0, 0)
                    except Exception:
                        pass
                    self.hwnd = None
    except Exception:
        pass
else:
    SysTrayIcon = None

def iniciar_interfaz():
    # AppUserModelID propio: evita que Windows agrupe la app bajo el ícono de python.exe en la barra de tareas
    if os.name == 'nt':
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Osocio.FormAutomation.App")
        except Exception:
            pass

    root = tk.Tk()
    root.title("Osocio - Form Automation")
    root.geometry("1150x680")  # Ventana con tamaño considerable
    root.minsize(1100, 600)
    root.configure(bg=APP_BG_COLOR)

    # Establecer el icono del oso de Ocioso.
    # iconbitmap(default=...) lo aplica a la ventana Y a las Toplevel; sin default= el .exe
    # empaquetado se queda con el icono genérico de Tk en la barra de tareas.
    icon_path = os.path.join(ASSET_DIR, "icon.ico")
    if os.path.exists(icon_path):
        try:
            root.iconbitmap(default=icon_path)
        except Exception:
            try:
                root.iconbitmap(icon_path)
            except Exception:
                pass
        # Refuerzo vía Win32: fija el icono grande (barra de tareas / Alt-Tab) y el chico
        # (barra de título) directo en la ventana, que es de donde Windows los toma.
        try:
            import ctypes
            root.update_idletasks()
            hwnd = ctypes.windll.user32.GetParent(root.winfo_id()) or root.winfo_id()
            LR_LOADFROMFILE, IMAGE_ICON, WM_SETICON = 0x0010, 1, 0x0080
            for size, which in ((32, 1), (16, 0)):  # 1 = ICON_BIG, 0 = ICON_SMALL
                h = ctypes.windll.user32.LoadImageW(None, icon_path, IMAGE_ICON,
                                                    size, size, LR_LOADFROMFILE)
                if h:
                    ctypes.windll.user32.SendMessageW(hwnd, WM_SETICON, which, h)
        except Exception:
            pass

    # Barra de título en modo oscuro (Windows DWM)
    def _dark_titlebar(window):
        try:
            import ctypes
            window.update_idletasks()
            hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
            val = ctypes.c_int(1)
            for attr in (20, 19):  # DWMWA_USE_IMMERSIVE_DARK_MODE (20, y 19 en builds viejas)
                ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, attr, ctypes.byref(val), ctypes.sizeof(val))
        except Exception:
            pass
    _dark_titlebar(root)
 
    # Variables globales compartidas entre pestañas para consistencia
    active_p_tab = ["Argentina"]
    
    # Helper constructora de botones con hovers amigables y Tabler Icons cargados dinámicamente
    def make_icon_btn(parent, text, text_color, command=None, pack_btn=True):
        icon_name = None
        clean_text = text
        if "Agregar Regla" in text:
            icon_name = "plus_green.png"
            clean_text = " Agregar Regla"
        elif "Agregar" in text or "+" in text:
            icon_name = "plus_green.png"
            clean_text = " Agregar"
        elif "Eliminar" in text or "🗑" in text:
            icon_name = "trash_coral.png"
            clean_text = " Eliminar"
        elif "Guardar" in text or "💾" in text:
            icon_name = "save_blue.png"
            clean_text = " " + text.replace("💾", "").strip()
        elif "Abrir Excel" in text or "📂" in text:
            icon_name = "folder_yellow.png"
            clean_text = " Abrir Excel"
        elif "Clonar" in text or "📋" in text:
            icon_name = "copy_lavender.png"
            clean_text = " Clonar"
        elif "Actualizar" in text or "🔄" in text:
            icon_name = "refresh_lavender.png"
            clean_text = " Actualizar"
        elif "Limpiar" in text or "🧹" in text:
            icon_name = "broom_lavender.png"
            clean_text = " Limpiar campos"
        elif "Regex" in text or "🔍" in text:
            icon_name = "search_lavender.png"
            clean_text = " Generador Regex"
 
        photo = get_button_icon(icon_name)

        # Todos los CTA "Guardar" comparten el mismo azul medio (relleno sólido)
        _is_guardar = ("Guardar" in text or "💾" in text)
        _base_bg = SAVE_BG if _is_guardar else BUTTON_INACTIVE
        _base_fg = SAVE_FG if _is_guardar else text_color
        _hover_bg = SAVE_HOVER if _is_guardar else BUTTON_HOVER

        if photo:
            btn = tk.Button(parent, text=clean_text, image=photo, compound="left",
                            font=("Segoe UI", 9, "bold"), bg=_base_bg, fg=_base_fg,
                            relief="flat", bd=0, activebackground=_hover_bg, activeforeground=_base_fg,
                            padx=12, pady=4, cursor="hand2", command=command)
        else:
            btn = tk.Button(parent, text=clean_text, font=("Segoe UI", 9, "bold"), bg=_base_bg, fg=_base_fg,
                            relief="flat", bd=0, activebackground=_hover_bg, activeforeground=_base_fg,
                            padx=12, pady=4, cursor="hand2", command=command)

        if pack_btn:
            btn.pack(side="left", padx=3)
        btn.bind("<Enter>", lambda e: btn.config(bg=_hover_bg))
        btn.bind("<Leave>", lambda e: btn.config(bg=_base_bg))
        return btn

    # Estilos ttk para Treeviews y Scrollbars de diseño premium
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass
        
    style.configure("Treeview", 
                    background=CARD_BG_COLOR, 
                    foreground="white", 
                    fieldbackground=CARD_BG_COLOR,
                    rowheight=26,
                    bordercolor=BORDER_COLOR,
                    borderwidth=0)
    style.configure("Treeview.Heading", 
                    background=BUTTON_INACTIVE, 
                    foreground="white", 
                    font=("Segoe UI", 9, "bold"),
                    borderwidth=0)
    style.map("Treeview.Heading", background=[('active', BUTTON_HOVER)])
    style.map("Treeview", background=[('selected', ACCENT_COLOR)], foreground=[('selected', BUTTON_INACTIVE)])

    # DISEÑO PREMIUM PARA LAS BARRAS DE SCROLL (ttk Scrollbar - TScrollbar unificado para clam)
    style.configure("TScrollbar", 
                    troughcolor="#100518",     # Fondo negroso del canal
                    background="#7D4E9F",      # Manija morada media distinguible
                    bordercolor=BORDER_COLOR, 
                    arrowcolor=TEXT_SECONDARY, 
                    lightcolor="#7D4E9F",
                    darkcolor="#7D4E9F",
                    relief="flat", 
                    borderwidth=0, 
                    arrowsize=10)
    style.map("TScrollbar",
              background=[('active', "#9662BD"), ('pressed', ACCENT_COLOR)],
              arrowcolor=[('active', 'white')])
    # Mismos colores para las variantes orientadas (algunas vistas usan estos nombres)
    for _sb in ("Vertical.TScrollbar", "Horizontal.TScrollbar"):
        style.configure(_sb, troughcolor="#100518", background="#7D4E9F",
                        bordercolor=BORDER_COLOR, arrowcolor=TEXT_SECONDARY,
                        lightcolor="#7D4E9F", darkcolor="#7D4E9F", relief="flat",
                        borderwidth=0, arrowsize=10)
        style.map(_sb, background=[('active', "#9662BD"), ('pressed', ACCENT_COLOR)],
                  arrowcolor=[('active', 'white')])

    # ==========================================
    # Helper para agregar logs a la consola inferior
    # ==========================================
    def log_message(msg):
        print(msg)

    # ==========================================
    # 1. CABECERA FIJA (Imagen de Portada)
    # ==========================================
    _fullheader_path = os.path.join(ASSET_DIR, "Fullheader.png")
    _fullheader_orig = None
    if os.path.exists(_fullheader_path):
        try:
            _fullheader_orig = Image.open(_fullheader_path)
        except Exception:
            pass

    _FH_HEIGHT = 120  # Altura un poco más compacta
    _FH_WIDTH = 1200
    if _fullheader_orig:
        _fh_orig_w, _fh_orig_h = _fullheader_orig.size
        _FH_WIDTH = max(1, int(round((_fh_orig_w * _FH_HEIGHT) / _fh_orig_h)))

    _HEADER_SPLIT_LEFT_COLOR = "#110830"
    _HEADER_SPLIT_RIGHT_COLOR = "#28164B"
    
    header_canvas = tk.Canvas(root, bd=0, highlightthickness=0, height=_FH_HEIGHT, bg=APP_BG_COLOR)
    header_canvas.pack(fill="x", side="top")

    _fh_photo_ref = [None]
    if _fullheader_orig:
        try:
            _fh_photo_ref[0] = ImageTk.PhotoImage(_fullheader_orig.resize((_FH_WIDTH, _FH_HEIGHT), Image.Resampling.LANCZOS))
        except Exception:
            _fh_photo_ref[0] = ImageTk.PhotoImage(_fullheader_orig)

    def _render_fixed_header(_event=None):
        header_canvas.delete("all")
        canvas_w = max(1, header_canvas.winfo_width())
        split_x = canvas_w // 2
        header_canvas.create_rectangle(0, 0, split_x, _FH_HEIGHT, fill=_HEADER_SPLIT_LEFT_COLOR, outline="")
        header_canvas.create_rectangle(split_x, 0, canvas_w, _FH_HEIGHT, fill=_HEADER_SPLIT_RIGHT_COLOR, outline="")
        if _fh_photo_ref[0]:
            x = (canvas_w - _FH_WIDTH) // 2
            header_canvas.create_image(x, 0, anchor="nw", image=_fh_photo_ref[0])
        else:
            header_canvas.create_text(canvas_w // 2, _FH_HEIGHT // 2, 
                                      text="OSOCIO FORM AUTOMATION", 
                                      fill="white", font=("Segoe UI", 16, "bold"))

    header_canvas.bind("<Configure>", _render_fixed_header)
    root.after(50, _render_fixed_header)

    # Separador debajo del header
    sep = tk.Frame(root, height=2, bg=BORDER_COLOR)
    sep.pack(fill="x")

    # ==========================================
    # 2. TOP BAR (Email Destinatario)
    # ==========================================
    top_bar = tk.Frame(root, bg=APP_BG_COLOR)
    top_bar.pack(fill="x", padx=20, pady=(6, 4))

    _cfg = {}
    try:
        _cfg = cargar_config_global()
    except Exception:
        pass
    _ui_prefs = _cfg.get("ui_prefs", {})
    var_ver_navegador = tk.BooleanVar(value=bool(_ui_prefs.get("visible_browser", False)))
    # Por defecto DESTILDADO: al cerrar la ventana el programa se cierra directo.
    # Si el usuario lo tilda, la X minimiza a la bandeja (persistente).
    var_minimizar_a_bandeja = tk.BooleanVar(value=bool(_ui_prefs.get("minimizar_a_bandeja", False)))
    var_pausar_autenticacion = tk.BooleanVar(value=bool(_ui_prefs.get("pausar_autenticacion", False)))

    def _on_pause_change(*_):
        if var_pausar_autenticacion.get():
            var_ver_navegador.set(True)

    def _on_visible_change(*_):
        if not var_ver_navegador.get() and var_pausar_autenticacion.get():
            var_pausar_autenticacion.set(False)

    var_pausar_autenticacion.trace_add("write", _on_pause_change)
    var_ver_navegador.trace_add("write", _on_visible_change)

    email_frame = tk.Frame(top_bar, bg=APP_BG_COLOR)
    email_frame.pack(side="right")

    # Ícono de sobre a la izquierda
    mail_ico = tk.Label(email_frame, text="✉", font=("Segoe UI", 11), bg=APP_BG_COLOR, fg=TEXT_SECONDARY)
    mail_ico.pack(side="left", padx=(0, 6))

    email_lbl = tk.Label(email_frame, text="Email destinatario:", font=("Segoe UI", 9, "bold"), bg=APP_BG_COLOR, fg=TEXT_PRIMARY)
    email_lbl.pack(side="left", padx=(0, 8))
    
    email_entry = tk.Entry(email_frame, font=("Segoe UI", 10), bg=ENTRY_BG, fg=TEXT_PRIMARY,
                           insertbackground="white", bd=0, relief="flat", width=34,
                           highlightthickness=1, highlightbackground=BORDER_COLOR, highlightcolor=ACCENT_COLOR,
                           disabledbackground=BUTTON_INACTIVE, disabledforeground=TEXT_SECONDARY)
    email_entry.pack(side="left", ipady=3)
    try:
        email_entry.insert(0, obtener_email_destinatario() or "")
    except Exception:
        pass
    email_entry.config(state="disabled")  # Habilitado sólo al activar "Enviar mail"

    # Variables de control para las opciones de email
    # "Enviar mail" nunca se restaura tildado al abrir la app (para evitar envíos
    # accidentales) aunque haya quedado guardado en una sesión anterior.
    var_enviar_email = tk.BooleanVar(value=False)
    var_adjuntar_res = tk.BooleanVar(value=bool(_cfg.get("adjuntar_resultados", False)))
    var_adjuntar_ss = tk.BooleanVar(value=bool(_cfg.get("adjuntar_screenshots", False)))
    var_modo_email = tk.StringVar(value=_cfg.get("email_modo", "consolidado"))

    # Frame para las opciones extras de email
    opts_frame = tk.Frame(email_frame, bg=APP_BG_COLOR)

    cb_adjuntar = tk.Checkbutton(opts_frame, text="Adjuntar resultados", variable=var_adjuntar_res,
                                 bg=APP_BG_COLOR, fg=TEXT_PRIMARY, selectcolor=ENTRY_BG, bd=0,
                                 activebackground=APP_BG_COLOR, activeforeground="white",
                                 font=("Segoe UI", 9), cursor="hand2")
    cb_adjuntar.pack(side="left", padx=(10, 0))
    
    cb_ss = tk.Checkbutton(opts_frame, text="Adjuntar screenshots", variable=var_adjuntar_ss,
                           bg=APP_BG_COLOR, fg=TEXT_PRIMARY, selectcolor=ENTRY_BG, bd=0,
                           activebackground=APP_BG_COLOR, activeforeground="white",
                           font=("Segoe UI", 9), cursor="hand2")
    cb_ss.pack(side="left", padx=(10, 0))
    
    lbl_modo = tk.Label(opts_frame, text="MODO:", font=("Segoe UI", 9, "bold"), bg=APP_BG_COLOR, fg=TEXT_SECONDARY)
    lbl_modo.pack(side="left", padx=(15, 5))
    
    rb_pais = tk.Radiobutton(opts_frame, text="1 por país", variable=var_modo_email, value="por_pais",
                             bg=APP_BG_COLOR, fg=TEXT_PRIMARY, selectcolor=ENTRY_BG, bd=0,
                             activebackground=APP_BG_COLOR, activeforeground="white",
                             font=("Segoe UI", 9), cursor="hand2")
    rb_pais.pack(side="left", padx=2)
    
    rb_cons = tk.Radiobutton(opts_frame, text="Consolidado", variable=var_modo_email, value="consolidado",
                             bg=APP_BG_COLOR, fg=TEXT_PRIMARY, selectcolor=ENTRY_BG, bd=0,
                             activebackground=APP_BG_COLOR, activeforeground="white",
                             font=("Segoe UI", 9), cursor="hand2")
    rb_cons.pack(side="left", padx=2)

    def toggle_email_options():
        if var_enviar_email.get():
            email_entry.config(state="normal")
            opts_frame.pack(side="left", padx=(10, 0))
        else:
            email_entry.config(state="disabled")
            opts_frame.pack_forget()

    cb_enviar = tk.Checkbutton(email_frame, text="Enviar mail", variable=var_enviar_email,
                               bg=APP_BG_COLOR, fg=TEXT_PRIMARY, selectcolor=ENTRY_BG, bd=0,
                               activebackground=APP_BG_COLOR, activeforeground="white",
                               font=("Segoe UI", 9, "bold"), cursor="hand2", command=toggle_email_options)
    cb_enviar.pack(side="left", padx=(10, 0))

    # ==========================================
    # 3. TABS BAR
    # ==========================================
    tab_bar = tk.Frame(root, bg=APP_BG_COLOR)
    tab_bar.pack(fill="x", padx=20, pady=(0, 4))

    content_area = tk.Frame(root, bg=APP_BG_COLOR)
    content_area.pack(fill="both", expand=True, padx=20, pady=(0, 4))

    tabs = {}
    tab_buttons = []

    def switch_tab(target_name):
        for name, frame in tabs.items():
            if name == target_name:
                frame.pack(fill="both", expand=True)
            else:
                frame.pack_forget()
        
        for btn in tab_buttons:
            if btn.tab_name == target_name:
                btn.config(bg=TAB_ACTIVE_BG, fg=TAB_ACTIVE_FG, highlightthickness=1, highlightbackground=BORDER_COLOR)
            else:
                btn.config(bg=TAB_INACTIVE_BG, fg=TAB_INACTIVE_FG, highlightthickness=0)

    tabs_data = [
        ("Envío de Leads", "leads", False),
        ("Programación de Tests", "scheduler", False),
        ("Validación de Campos", "validation", False),
        ("Generar Excels con Datos", "excel", False),
        ("Comparar Dealers vs Form", "dealers", False),
    ]
    for t_text, t_name, t_disabled in tabs_data:
        btn = tk.Button(tab_bar, text=t_text, font=("Segoe UI", 9, "bold"), bg=TAB_INACTIVE_BG, fg=TAB_INACTIVE_FG,
                        relief="flat", bd=0, activebackground=TAB_ACTIVE_BG, activeforeground=TAB_ACTIVE_FG,
                        padx=18, pady=8, highlightthickness=0)
        btn.tab_name = t_name
        btn.pack(side="left", padx=2)
        if t_disabled:
            btn.config(state="disabled", disabledforeground="#8A6E9E", cursor="arrow")
        else:
            btn.config(command=lambda name=t_name: switch_tab(name))
        tab_buttons.append(btn)

    global_config_btn = tk.Button(
        tab_bar, text="⚙ Configurar", font=("Segoe UI", 9, "bold"),
        bg="#1F618D", fg="#f4f4f4", activebackground="#2980B9", activeforeground="#f4f4f4",
        relief="flat", bd=0, padx=15, pady=8, cursor="hand2"
    )
    global_config_btn.pack(side="right", padx=(0, 2))

    for _, t_name, _ in tabs_data:
        tabs[t_name] = tk.Frame(content_area, bg=APP_BG_COLOR)

    # Canvases registrados para el scroll con rueda de mouse
    scrollable_canvases = []

    # Helper para crear un panel scrolleable verticalmente dentro de una pestaña
    def make_scrollable_tab_container(parent):
        canvas = tk.Canvas(parent, bg=APP_BG_COLOR, bd=0, highlightthickness=0)
        v_scroll = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview, style="TScrollbar")
        
        inner_frame = tk.Frame(canvas, bg=APP_BG_COLOR)
        
        # Sincronizar el ancho del frame interno con el canvas
        def _on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)
            
        canvas.bind("<Configure>", _on_canvas_configure)
        inner_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        canvas_window = canvas.create_window((0, 0), window=inner_frame, anchor="nw")
        canvas.configure(yscrollcommand=v_scroll.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        v_scroll.pack(side="right", fill="y")
        scrollable_canvases.append(canvas)
        return inner_frame

    # ==========================================
    # TAB 1: ENVÍO DE LEADS (leads)
    # ==========================================
    leads_scroll_frame = make_scrollable_tab_container(tabs["leads"])

    # --- Barra superior: acceso a IDs Dinámicos (campos no mapeados) ---
    def _abrir_ids_dinamicos():
        if abrir_popup_ids_dinamicos is None:
            messagebox.showwarning(
                "IDs Dinámicos",
                "No se pudo cargar el módulo de IDs Dinámicos.\n" + (_BACKEND_IMPORT_ERROR or ""),
            )
            return
        try:
            abrir_popup_ids_dinamicos(leads_scroll_frame.winfo_toplevel())
        except Exception as exc:
            messagebox.showerror("IDs Dinámicos", f"No se pudo abrir IDs Dinámicos:\n{exc}")

    topbar_leads = tk.Frame(leads_scroll_frame, bg=APP_BG_COLOR)
    topbar_leads.pack(fill="x", pady=(0, 6))
    tk.Button(
        topbar_leads,
        text="⚙  IDs Dinámicos",
        command=_abrir_ids_dinamicos,
        bg="#FFC845",
        fg="#3A1D52",
        activebackground="#FFD873",
        activeforeground="#3A1D52",
        relief="flat",
        bd=0,
        highlightthickness=2,
        highlightbackground="#FFE08A",
        font=("Segoe UI", 10, "bold"),
        cursor="hand2",
        padx=16,
        pady=6,
    ).pack(side="right")

    # Modo de Excel: "por_dispositivo" (default, un Excel por dispositivo) o
    # "compartido" (un único Excel genérico con los mismos datos para todos).
    try:
        _excel_mode_ini = (cargar_config_global() or {}).get("excel_mode", "por_dispositivo")
    except Exception:
        _excel_mode_ini = "por_dispositivo"
    excel_mode_holder = [_excel_mode_ini if _excel_mode_ini in ("por_dispositivo", "compartido") else "por_dispositivo"]

    # --- 1. CONFIGURACIÓN DE EJECUCIÓN (Card Panel) ---
    config_card = tk.Frame(leads_scroll_frame, bg=CARD_BG_COLOR, bd=0, highlightthickness=1, highlightbackground=BORDER_COLOR)
    config_card.pack(fill="x", pady=(0, 8), ipady=6)

    sec_head = tk.Frame(config_card, bg=CARD_BG_COLOR)
    sec_head.pack(fill="x", padx=15, pady=(6, 2))
    sec_title = tk.Label(sec_head, text="Configurá tu envío: elegí mercado, modo de ejecución y dispositivos",
                         font=("Segoe UI", 10, "bold"), bg=CARD_BG_COLOR, fg=TEXT_PRIMARY)
    sec_title.pack(side="left")
    sec_hint = tk.Label(config_card,
                        text="① Elegí el modo de mercados y de Excels.   ② Tildá los dispositivos/navegadores.   "
                             "③ Con el botón “Configurar” arriba a la derecha cambiás si cada dispositivo usa su Excel o uno compartido.",
                        font=("Segoe UI", 8, "italic"), bg=CARD_BG_COLOR, fg="#C5A9DF", justify="left")
    sec_hint.pack(anchor="w", padx=15, pady=(0, 6))

    row_config = tk.Frame(config_card, bg=CARD_BG_COLOR)
    row_config.pack(fill="x", padx=15)

    # Piloto de grupos con botones pastel suaves. Devuelve un holder [valor] con la selección.
    def make_pill_group(parent, label_text, subtitle_text, options, default_val, sub_texts=None, on_change=None):
        group_frame = tk.Frame(parent, bg=CARD_BG_COLOR)
        group_frame.pack(side="left", padx=(0, 25), anchor="n")

        lbl = tk.Label(group_frame, text=label_text, font=("Segoe UI", 8, "bold"), bg=CARD_BG_COLOR, fg=TEXT_SECONDARY)
        lbl.pack(anchor="w", pady=(0, 3))

        btn_row = tk.Frame(group_frame, bg=CARD_BG_COLOR)
        btn_row.pack(anchor="w")

        selected_val = [default_val]
        btns = {}
        sub_lbl = None

        def on_click(val):
            selected_val[0] = val
            for v, b in btns.items():
                if v == val:
                    b.config(bg=BUTTON_ACTIVE, fg="white", highlightthickness=1, highlightbackground=ACCENT_COLOR)
                else:
                    b.config(bg=BUTTON_INACTIVE, fg=TEXT_SECONDARY, highlightthickness=1, highlightbackground=BUTTON_INACTIVE)
            if sub_texts and sub_lbl is not None:
                sub_lbl.config(text=sub_texts.get(val, subtitle_text))
            if on_change:
                try:
                    on_change(val)
                except Exception:
                    pass

        for opt in options:
            # opt puede ser "Texto" (el valor se deriva en minúscula) o ("Texto visible", "valor_interno"),
            # para poder cambiar la etiqueta sin romper lo que ya está guardado en config/schedule.
            opt_txt, opt_val = opt if isinstance(opt, tuple) else (opt, opt.lower())
            b = tk.Button(btn_row, text=opt_txt, font=("Segoe UI", 8, "bold"), bg=BUTTON_INACTIVE, fg=TEXT_SECONDARY,
                          relief="flat", bd=0, activebackground=BUTTON_HOVER, activeforeground="white",
                          highlightthickness=1, highlightbackground=BUTTON_INACTIVE,
                          padx=12, pady=4, cursor="hand2")
            b.pack(side="left", padx=2)
            btns[opt_val] = b
            b.config(command=lambda v=opt_val: on_click(v))

            def on_enter(e, btn=b, val=opt_val):
                if selected_val[0] != val:
                    btn.config(bg=BUTTON_HOVER)
            def on_leave(e, btn=b, val=opt_val):
                if selected_val[0] != val:
                    btn.config(bg=BUTTON_INACTIVE)
            b.bind("<Enter>", on_enter)
            b.bind("<Leave>", on_leave)

        on_click(default_val)

        sub_lbl = tk.Label(group_frame, text=(sub_texts.get(default_val, subtitle_text) if sub_texts else subtitle_text),
                           font=("Segoe UI", 8, "italic"), bg=CARD_BG_COLOR, fg="#C5A9DF", wraplength=180, justify="left")
        sub_lbl.pack(anchor="w", pady=(3, 0))
        return selected_val

    saved_mercados_mode = _ui_prefs.get("mercados_mode", "consecutivo")
    saved_excels_mode = _ui_prefs.get("excels_mode", "consecutivo")

    mercados_mode = make_pill_group(
        row_config, "MERCADOS", "", [("Secuencial", "consecutivo"), ("Paralelo", "paralelo")], saved_mercados_mode,
        sub_texts={"consecutivo": "Un mercado a la vez (AR → BO → …)", "paralelo": "Todos los mercados a la vez"},
        on_change=lambda v: _save_ui_prefs())
    excels_mode = make_pill_group(
        row_config, "EXCELS POR DISPOSITIVO", "", [("Secuencial", "consecutivo"), ("Paralelo", "paralelo")], saved_excels_mode,
        sub_texts={"consecutivo": "Un dispositivo a la vez (Chrome → Firefox → …)", "paralelo": "Todos a la vez (solo browsers locales Chrome/FF/Edge)"},
        on_change=lambda v: [_refresh_excel_par_warning(), _save_ui_prefs()])

    # Dispositivos y Navegadores
    disp_frame = tk.Frame(row_config, bg=CARD_BG_COLOR)
    disp_frame.pack(side="left", anchor="n", padx=(0, 20))

    disp_lbl = tk.Label(disp_frame, text="DISPOSITIVOS / NAVEGADORES", font=("Segoe UI", 8, "bold"), bg=CARD_BG_COLOR, fg=TEXT_SECONDARY)
    disp_lbl.pack(anchor="w", pady=(0, 3))

    disp_btn_row = tk.Frame(disp_frame, bg=CARD_BG_COLOR)
    disp_btn_row.pack(anchor="w")

    dispositivos = ["Chrome", "Firefox", "Edge", "Mac LT", "Android LT"]
    saved_disp = _ui_prefs.get("selected_disp", {})
    selected_disp = {d.lower(): False for d in dispositivos}
    if saved_disp:
        for k in selected_disp:
            if k in saved_disp:
                selected_disp[k] = bool(saved_disp[k])
    else:
        selected_disp["chrome"] = True
    disp_btns = {}

    status_lbl = tk.Label(disp_frame, text="1 dispositivo seleccionado.", font=("Segoe UI", 8, "italic"), bg=CARD_BG_COLOR, fg="#C5A9DF")

    def _load_lt_creds():
        path = os.path.join(BASE_DIR, "lambdatest_credentials.txt")
        u = ak = ""
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if "=" not in line or line.startswith("#"):
                            continue
                        k, v = line.split("=", 1)
                        k = k.strip().lower()
                        if k == "username":
                            u = v.strip()
                        elif k == "access_key":
                            ak = v.strip()
            except Exception:
                pass
        return u, ak

    def toggle_disp(name):
        key = name.lower()
        selected_disp[key] = not selected_disp[key]
        btn = disp_btns[key]
        if selected_disp[key]:
            btn.config(bg=BUTTON_ACTIVE, fg="white", highlightthickness=1, highlightbackground=ACCENT_COLOR)
        else:
            btn.config(bg=BUTTON_INACTIVE, fg=TEXT_SECONDARY, highlightthickness=1, highlightbackground=BUTTON_INACTIVE)
        
        # Panel LT dinámico
        if selected_disp["mac lt"] or selected_disp["android lt"]:
            # Recargar credenciales desde el archivo cuando se selecciona mac o android
            u, ak = _load_lt_creds()
            lt_user_ent.delete(0, tk.END)
            lt_user_ent.insert(0, u)
            lt_key_ent.delete(0, tk.END)
            lt_key_ent.insert(0, ak)
            lt_creds_frame.pack(side="left", anchor="n", padx=(10, 0))
        else:
            lt_creds_frame.pack_forget()
            
        count = sum(1 for v in selected_disp.values() if v)
        status_lbl.config(text=f"{count} dispositivo{'s' if count != 1 else ''} seleccionado{'s' if count != 1 else ''}.")
        actualizar_warning()
        try:
            _refresh_excel_par_warning()
        except Exception:
            pass

        refresh_ver_nav_state()
        _save_ui_prefs()

    for disp in dispositivos:
        d_key = disp.lower()
        is_sel = selected_disp[d_key]
        init_bg = BUTTON_ACTIVE if is_sel else BUTTON_INACTIVE
        init_fg = "white" if is_sel else TEXT_SECONDARY
        init_hb = ACCENT_COLOR if is_sel else BUTTON_INACTIVE
        
        b = tk.Button(disp_btn_row, text=disp, font=("Segoe UI", 8, "bold"), bg=init_bg, fg=init_fg,
                      relief="flat", bd=0, activebackground=BUTTON_HOVER, activeforeground="white",
                      highlightthickness=1, highlightbackground=init_hb,
                      padx=10, pady=4, cursor="hand2")
        b.pack(side="left", padx=2)
        disp_btns[d_key] = b
        b.config(command=lambda n=disp: toggle_disp(n))

        def make_disp_hover(btn=b, k=d_key):
            def on_enter(e):
                if not selected_disp[k]:
                    btn.config(bg=BUTTON_HOVER)
            def on_leave(e):
                if not selected_disp[k]:
                    btn.config(bg=BUTTON_INACTIVE)
            btn.bind("<Enter>", on_enter)
            btn.bind("<Leave>", on_leave)
        make_disp_hover()

    status_lbl.pack(anchor="w", pady=(3, 0))

    # "Ver navegador" y "Minimizar a la bandeja al cerrar" viven en Configuración
    # avanzada (botón "Configurar"). Acá sólo quedan las variables persistentes.

    def _save_ui_prefs(*_):
        try:
            cfg = cargar_config_global()
            if "ui_prefs" not in cfg:
                cfg["ui_prefs"] = {}
            cfg["ui_prefs"]["visible_browser"] = bool(var_ver_navegador.get())
            cfg["ui_prefs"]["minimizar_a_bandeja"] = bool(var_minimizar_a_bandeja.get())
            cfg["ui_prefs"]["pausar_autenticacion"] = bool(var_pausar_autenticacion.get())
            cfg["ui_prefs"]["selected_disp"] = selected_disp
            if 'mercados_mode' in globals() or 'mercados_mode' in locals():
                cfg["ui_prefs"]["mercados_mode"] = mercados_mode[0]
            if 'excels_mode' in globals() or 'excels_mode' in locals():
                cfg["ui_prefs"]["excels_mode"] = excels_mode[0]
            if 'var_url_parallel' in globals() or 'var_url_parallel' in locals():
                cfg["ui_prefs"]["url_parallel"] = bool(var_url_parallel.get())
            if 'url_max_var' in globals() or 'url_max_var' in locals():
                cfg["ui_prefs"]["url_max"] = url_max_var.get()
            if 'var_t3' in globals() or 'var_t3' in locals():
                cfg["ui_prefs"]["t3"] = bool(var_t3.get())
            
            # Email fields (in both root and ui_prefs)
            enviar_mail = bool(var_enviar_email.get())
            dest = email_entry.get().strip()
            adj_res = bool(var_adjuntar_res.get())
            adj_ss = bool(var_adjuntar_ss.get())
            email_modo = var_modo_email.get()
            
            cfg["enviar_mail"] = enviar_mail
            cfg["email_destinatario"] = dest
            cfg["adjuntar_resultados"] = adj_res
            cfg["adjuntar_screenshots"] = adj_ss
            cfg["email_modo"] = email_modo
            
            cfg["ui_prefs"]["enviar_mail"] = enviar_mail
            cfg["ui_prefs"]["adjuntar_resultados"] = adj_res
            cfg["ui_prefs"]["adjuntar_screenshots"] = adj_ss
            cfg["ui_prefs"]["email_modo"] = email_modo
            
            guardar_config_global(cfg)
        except Exception:
            pass

    var_ver_navegador.trace_add("write", _save_ui_prefs)
    var_pausar_autenticacion.trace_add("write", _save_ui_prefs)
    var_minimizar_a_bandeja.trace_add("write", _save_ui_prefs)

    # Enviar en paralelo POR URL: una sesión (navegador) por URL, todas en simultáneo
    var_url_parallel = tk.BooleanVar(value=bool(_ui_prefs.get("url_parallel", False)))
    url_max_var = tk.StringVar(value=_ui_prefs.get("url_max", "6"))
    url_par_row = tk.Frame(disp_frame, bg=CARD_BG_COLOR)
    url_par_row.pack(anchor="w", pady=(2, 0))
    cb_url_par = tk.Checkbutton(url_par_row, text="⚡ Enviar en paralelo por URL (una sesión por URL)", variable=var_url_parallel,
                                bg=CARD_BG_COLOR, fg=TEXT_SECONDARY, selectcolor=ENTRY_BG, bd=0,
                                activebackground=CARD_BG_COLOR, activeforeground="white",
                                font=("Segoe UI", 8), cursor="hand2")
    cb_url_par.pack(side="left")
    tk.Label(url_par_row, text="máx. simultáneas:", font=("Segoe UI", 8), bg=CARD_BG_COLOR, fg=TEXT_SECONDARY).pack(side="left", padx=(10, 4))
    tk.Entry(url_par_row, textvariable=url_max_var, width=4, font=("Segoe UI", 8), bg=ENTRY_BG, fg="white",
             bd=0, relief="flat", highlightthickness=1, highlightbackground=BORDER_COLOR, justify="center").pack(side="left", ipady=1)

    # Formularios T3 2.0 (Adobe AEM): usa los Excels con nombre …_T3.xlsx
    var_t3 = tk.BooleanVar(value=bool(_ui_prefs.get("t3", False)))
    t3_row = tk.Frame(disp_frame, bg=CARD_BG_COLOR)
    t3_row.pack(anchor="w", pady=(2, 0))
    tk.Checkbutton(t3_row, text="🧩 Formularios T3 2.0 (usa los Excels …_T3)", variable=var_t3,
                   command=lambda: update_table_data(active_p_tab[0]),
                   bg=CARD_BG_COLOR, fg=TEXT_SECONDARY, selectcolor=ENTRY_BG, bd=0,
                   activebackground=CARD_BG_COLOR, activeforeground="white",
                   font=("Segoe UI", 8), cursor="hand2").pack(side="left")

    def refresh_ver_nav_state():
        # "Ver navegador" sólo aplica a browsers de escritorio (chrome/firefox/edge).
        # Si no hay ninguno seleccionado, se fuerza apagado (el checkbox está en Config avanzada).
        if not any(selected_disp.get(b) for b in ("chrome", "firefox", "edge")):
            var_ver_navegador.set(False)
    refresh_ver_nav_state()

    var_url_parallel.trace_add("write", _save_ui_prefs)
    url_max_var.trace_add("write", _save_ui_prefs)
    var_t3.trace_add("write", _save_ui_prefs)
    var_enviar_email.trace_add("write", _save_ui_prefs)
    var_adjuntar_res.trace_add("write", _save_ui_prefs)
    var_adjuntar_ss.trace_add("write", _save_ui_prefs)
    var_modo_email.trace_add("write", _save_ui_prefs)
    email_entry.bind("<FocusOut>", lambda e: _save_ui_prefs())

    # Disclaimer dinámico: "Excels en paralelo" no aplica a LambdaTest.
    excel_par_warn_lbl = tk.Label(config_card, text="", font=("Segoe UI", 8, "bold"),
                                  bg=CARD_BG_COLOR, fg=WARN_TEXT, wraplength=760, justify="left")
    excel_par_warn_lbl.pack(anchor="w", padx=15, pady=(2, 0))

    def _refresh_excel_par_warning(*_):
        try:
            _lt_sel = bool(selected_disp.get("mac lt") or selected_disp.get("android lt"))
            _par = (excels_mode[0] == "paralelo")
        except Exception:
            return
        if _par and _lt_sel:
            excel_par_warn_lbl.config(
                text="⚠ 'Excels en paralelo' NO aplica a LambdaTest (Mac/Android): esos dispositivos se "
                     "ejecutan igual de forma consecutiva. El paralelo sólo acelera los browsers locales "
                     "(Chrome/Firefox/Edge).")
        else:
            excel_par_warn_lbl.config(text="")
    _refresh_excel_par_warning()

    # CREDENCIALES LAMBDATEST INLINE
    lt_creds_frame = tk.Frame(row_config, bg=CARD_BG_COLOR)
    if selected_disp["mac lt"] or selected_disp["android lt"]:
        lt_creds_frame.pack(side="left", anchor="n", padx=(10, 0))
    else:
        lt_creds_frame.pack_forget()

    tk.Label(lt_creds_frame, text="CREDENCIALES LT", font=("Segoe UI", 8, "bold"), bg=CARD_BG_COLOR, fg=TEXT_SECONDARY).grid(row=0, column=0, columnspan=3, sticky="w")
    
    tk.Label(lt_creds_frame, text="User:", font=("Segoe UI", 8), bg=CARD_BG_COLOR, fg="white").grid(row=1, column=0, sticky="w", pady=1)
    lt_user_ent = tk.Entry(lt_creds_frame, font=("Segoe UI", 8), bg=ENTRY_BG, fg="white", bd=0, highlightthickness=1, highlightbackground=BORDER_COLOR, width=28)
    lt_user_ent.grid(row=1, column=1, sticky="w", padx=4, pady=1)
    # Cargar credenciales al inicializar
    _u_init, _ak_init = _load_lt_creds()
    lt_user_ent.insert(0, _u_init)

    tk.Label(lt_creds_frame, text="Key:", font=("Segoe UI", 8), bg=CARD_BG_COLOR, fg="white").grid(row=2, column=0, sticky="w", pady=1)
    lt_key_ent = tk.Entry(lt_creds_frame, font=("Segoe UI", 8), bg=ENTRY_BG, fg="white", bd=0, highlightthickness=1, highlightbackground=BORDER_COLOR, width=28, show="*")
    lt_key_ent.grid(row=2, column=1, sticky="w", padx=4, pady=1)
    lt_key_ent.insert(0, _ak_init)

    # CTA de guardado de credenciales LambdaTest
    def _save_lt_creds():
        user_val = lt_user_ent.get().strip()
        key_val = lt_key_ent.get().strip()
        path = os.path.join(BASE_DIR, "lambdatest_credentials.txt")
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(f"username={user_val}\n")
                f.write(f"access_key={key_val}\n")
            messagebox.showinfo("LambdaTest", "✓ Credenciales de LambdaTest guardadas con éxito en lambdatest_credentials.txt")
            log_message("[INFO] Credenciales de LambdaTest actualizadas en lambdatest_credentials.txt.")
        except Exception as e:
            messagebox.showerror("LambdaTest", f"Error al guardar credenciales:\n{e}")
    
    # Debajo de User/Key (antes iba a la derecha en una 3ra columna y quedaba cortado)
    lt_save_btn = make_icon_btn(lt_creds_frame, "💾 Guardar", TEXT_SAVE, command=_save_lt_creds, pack_btn=False)
    lt_save_btn.grid(row=3, column=0, columnspan=2, padx=4, pady=(5, 1), sticky="ew")

    # ADVERTENCIA DINÁMICA DE MÚLTIPLES DISPOSITIVOS (Figma Style)
    warning_box = tk.Frame(config_card, bg=WARN_BG, bd=1, highlightthickness=1, highlightbackground=WARN_BORDER)
    warning_box.pack_forget()

    warn_lbl = tk.Label(warning_box, text="⚠ Cuidado. Si corrés el mismo form en varios navegadores o dispositivos, cada uno necesita datos únicos — de lo contrario los leads pueden rechazarse o duplicarse.",
                        font=("Segoe UI", 8, "bold"), bg=WARN_BG, fg=WARN_TEXT, justify="left")
    warn_lbl.pack(side="left", padx=12, pady=6)

    warn_cta = tk.Button(warning_box, text="Ir a Generar Excels ➜", font=("Segoe UI", 8, "bold"), bg=WARN_BG, fg=WARN_TEXT,
                         relief="flat", bd=0, highlightthickness=1, highlightbackground=WARN_BORDER, padx=10, pady=2, cursor="hand2", command=lambda: switch_tab("excel"))
    warn_cta.pack(side="right", padx=12, pady=6)

    _WARN_SHARED = ("⚠ Modo Excel compartido activo: TODOS los dispositivos usan el mismo Excel genérico "
                    "(los mismos datos). Los leads pueden duplicarse o rechazarse. Cambiá el modo con la ⚙.")

    def actualizar_warning():
        # El warning solo aplica al modo compartido (mismos datos → riesgo de duplicados).
        # En modo por-dispositivo cada Excel tiene datos únicos, así que no se muestra.
        if excel_mode_holder[0] == "compartido":
            warn_lbl.config(text=_WARN_SHARED)
            warning_box.pack(fill="x", padx=15, pady=(8, 0), before=row_config)
        else:
            warning_box.pack_forget()

    def abrir_config_avanzada():
        win = tk.Toplevel(root)
        win.title("Configuración avanzada")
        win.configure(bg=CARD_BG_COLOR)
        win.transient(root)
        win.resizable(False, False)
        w, h = 560, 480
        px = root.winfo_rootx() + (root.winfo_width() - w) // 2
        py = root.winfo_rooty() + (root.winfo_height() - h) // 2
        win.geometry(f"{w}x{h}+{max(px,0)}+{max(py,0)}")

        tk.Label(win, text="⚙  Configuración avanzada", font=("Segoe UI", 13, "bold"),
                 bg=CARD_BG_COLOR, fg=TEXT_PRIMARY).pack(anchor="w", padx=18, pady=(16, 2))
        tk.Label(win, text="Opciones que no forman parte de la pantalla principal.",
                 font=("Segoe UI", 8), bg=CARD_BG_COLOR, fg=TEXT_SECONDARY).pack(anchor="w", padx=18, pady=(0, 12))

        tk.Label(win, text="EXCEL POR DISPOSITIVO (SOLO ENVÍO DE LEADS / PROGRAMADOS)", font=("Segoe UI", 9, "bold"),
                 bg=CARD_BG_COLOR, fg=TEXT_SECONDARY).pack(anchor="w", padx=18)

        mode_var = tk.StringVar(value=excel_mode_holder[0])
        opts_frame = tk.Frame(win, bg=CARD_BG_COLOR)
        opts_frame.pack(fill="x", padx=18, pady=(6, 0))

        tk.Radiobutton(opts_frame, variable=mode_var, value="por_dispositivo",
                       text="Un Excel por dispositivo (recomendado)",
                       font=("Segoe UI", 9), bg=CARD_BG_COLOR, fg=TEXT_PRIMARY,
                       selectcolor=CARD_BG_COLOR, activebackground=CARD_BG_COLOR,
                       activeforeground=TEXT_PRIMARY, anchor="w").pack(anchor="w", pady=2)
        tk.Label(opts_frame, text="Cada dispositivo usa su propio Excel con datos únicos (…_Chrome/_Firefox/_Edge/_Mac/_Android.xlsx).",
                 font=("Segoe UI", 8), bg=CARD_BG_COLOR, fg=TEXT_SECONDARY, wraplength=500, justify="left").pack(anchor="w", padx=(24, 0), pady=(0, 8))

        tk.Radiobutton(opts_frame, variable=mode_var, value="compartido",
                       text="Un Excel compartido para todos los dispositivos",
                       font=("Segoe UI", 9), bg=CARD_BG_COLOR, fg=TEXT_PRIMARY,
                       selectcolor=CARD_BG_COLOR, activebackground=CARD_BG_COLOR,
                       activeforeground=TEXT_PRIMARY, anchor="w").pack(anchor="w", pady=2)
        tk.Label(opts_frame, text="Todos los dispositivos (incluido LambdaTest) corren los MISMOS datos desde un único Excel genérico (…_Generico.xlsx).",
                 font=("Segoe UI", 8), bg=CARD_BG_COLOR, fg=TEXT_SECONDARY, wraplength=500, justify="left").pack(anchor="w", padx=(24, 0), pady=(0, 4))

        disc = tk.Label(win, text="⚠ Al compartir los mismos datos, los leads pueden duplicarse o ser rechazados por el formulario.",
                        font=("Segoe UI", 8, "bold"), bg=CARD_BG_COLOR, fg=WARN_TEXT, wraplength=500, justify="left")
        disc.pack(anchor="w", padx=18, pady=(6, 0))

        # ── Comportamiento de la app ────────────────────────────────────────────
        tk.Label(win, text="COMPORTAMIENTO DE LA APP", font=("Segoe UI", 9, "bold"),
                 bg=CARD_BG_COLOR, fg=TEXT_SECONDARY).pack(anchor="w", padx=18, pady=(14, 4))
        beh_frame = tk.Frame(win, bg=CARD_BG_COLOR)
        beh_frame.pack(fill="x", padx=18)
        tk.Checkbutton(beh_frame, text="Ver navegador mientras corre  (solo Chrome / Firefox / Edge)",
                       variable=var_ver_navegador, font=("Segoe UI", 9), bg=CARD_BG_COLOR, fg=TEXT_PRIMARY,
                       selectcolor=CARD_BG_COLOR, activebackground=CARD_BG_COLOR, activeforeground=TEXT_PRIMARY,
                       anchor="w").pack(anchor="w", pady=2)
        tk.Checkbutton(beh_frame, text="Minimizar a la bandeja al cerrar (en vez de cerrar el programa)",
                       variable=var_minimizar_a_bandeja, font=("Segoe UI", 9), bg=CARD_BG_COLOR, fg=TEXT_PRIMARY,
                       selectcolor=CARD_BG_COLOR, activebackground=CARD_BG_COLOR, activeforeground=TEXT_PRIMARY,
                       anchor="w").pack(anchor="w", pady=2)
        tk.Checkbutton(beh_frame, text="Pausar para login manual en primer formulario (Envío de Leads / Comparador)",
                       variable=var_pausar_autenticacion, font=("Segoe UI", 9), bg=CARD_BG_COLOR, fg=TEXT_PRIMARY,
                       selectcolor=CARD_BG_COLOR, activebackground=CARD_BG_COLOR, activeforeground=TEXT_PRIMARY,
                       anchor="w").pack(anchor="w", pady=2)

        btns = tk.Frame(win, bg=CARD_BG_COLOR)
        btns.pack(side="bottom", fill="x", padx=18, pady=14)

        def _guardar():
            nuevo = mode_var.get()
            excel_mode_holder[0] = nuevo
            try:
                _cfg = cargar_config_global() or {}
                _cfg["excel_mode"] = nuevo
                guardar_config_global(_cfg)
            except Exception:
                pass
            actualizar_warning()
            try:
                update_excel_calculation()
            except Exception:
                pass
            win.destroy()

        tk.Button(btns, text="Guardar", font=("Segoe UI", 9, "bold"), bg=SAVE_BG, fg=SAVE_FG,
                  activebackground=SAVE_HOVER, activeforeground=SAVE_FG,
                  relief="flat", bd=0, padx=18, pady=6, cursor="hand2", command=_guardar).pack(side="right")
        tk.Button(btns, text="Cancelar", font=("Segoe UI", 9), bg=BUTTON_INACTIVE, fg=TEXT_SECONDARY,
                  relief="flat", bd=0, padx=14, pady=6, cursor="hand2", command=win.destroy).pack(side="right", padx=(0, 8))

    actualizar_warning()  # reflejar el modo persistido al abrir la app
    global_config_btn.config(command=abrir_config_avanzada)

    # --- 2. PAÍSES A EJECUTAR (Card Panel) ---
    paises_card = tk.Frame(leads_scroll_frame, bg=CARD_BG_COLOR, bd=0, highlightthickness=1, highlightbackground=BORDER_COLOR)
    paises_card.pack(fill="x", pady=(0, 8), ipady=5)

    p_header = tk.Frame(paises_card, bg=CARD_BG_COLOR)
    p_header.pack(fill="x", padx=15, pady=(6, 4))
    
    tk.Label(p_header, text="🌐 PAÍSES A EJECUTAR", font=("Segoe UI", 10, "bold"), bg=CARD_BG_COLOR, fg=TEXT_PRIMARY).pack(side="left")

    paises_list = ["Argentina", "Bolivia", "Brasil", "Chile", "Colombia", "Ecuador", "Paraguay", "Peru", "Uruguay"]
    p_codes = {"Argentina": "AR", "Bolivia": "BO", "Brasil": "BR", "Chile": "CL", "Colombia": "CO", "Ecuador": "EC", "Paraguay": "PY", "Peru": "PE", "Uruguay": "UY"}
    selected_countries = {p: False for p in paises_list}
    country_cards = {}
    country_labels = {}

    counter_lbl = tk.Label(p_header, text="0 seleccionados", font=("Segoe UI", 8, "bold"), bg=CARD_BG_COLOR, fg=TEXT_SECONDARY)
    counter_lbl.pack(side="right", padx=(10, 0))

    tk.Label(paises_card, text="Elegí el o los países que querés ejecutar.",
             font=("Segoe UI", 8, "italic"), bg=CARD_BG_COLOR, fg="#C5A9DF").pack(anchor="w", padx=15, pady=(0, 4))

    def update_country_counter():
        count = sum(1 for v in selected_countries.values() if v)
        counter_lbl.config(text=f"{count} seleccionados")

    def toggle_country(name):
        selected_countries[name] = not selected_countries[name]
        card = country_cards[name]
        c_lbl = country_labels[name]
        if selected_countries[name]:
            card.config(highlightbackground=ACCENT_COLOR, bg=BUTTON_INACTIVE)
            c_lbl.config(fg=ACCENT_COLOR, bg=BUTTON_INACTIVE)
        else:
            card.config(highlightbackground=BORDER_COLOR, bg=CARD_BG_COLOR)
            c_lbl.config(fg="white", bg=CARD_BG_COLOR)
        update_country_counter()
        refresh_execute_state()

    grid_frame = tk.Frame(paises_card, bg=CARD_BG_COLOR)
    grid_frame.pack(fill="x", padx=15, pady=2)

    for idx, pais in enumerate(paises_list):
        code = p_codes[pais]
        card = tk.Frame(grid_frame, bg=CARD_BG_COLOR, bd=0, highlightthickness=1, highlightbackground=BORDER_COLOR, cursor="hand2")
        card.grid(row=idx // 9, column=idx % 9, padx=3, pady=3, sticky="nsew")
        grid_frame.columnconfigure(idx % 9, weight=1)

        code_lbl = tk.Label(card, text=code, font=("Segoe UI", 11, "bold"), bg=CARD_BG_COLOR, fg=TEXT_PRIMARY, cursor="hand2")
        code_lbl.pack(pady=(5, 1))
        
        name_lbl = tk.Label(card, text=pais, font=("Segoe UI", 8), bg=CARD_BG_COLOR, fg=TEXT_SECONDARY, cursor="hand2")
        name_lbl.pack(pady=(0, 5))

        country_cards[pais] = card
        country_labels[pais] = name_lbl

        def bind_click(w, p=pais):
            w.bind("<Button-1>", lambda e: toggle_country(p))
        bind_click(card)
        bind_click(code_lbl)
        bind_click(name_lbl)

    links_frame = tk.Frame(paises_card, bg=CARD_BG_COLOR)
    links_frame.pack(fill="x", padx=15, pady=(2, 2))

    def select_all_countries():
        for p in paises_list:
            if not selected_countries[p]:
                toggle_country(p)
    def select_none_countries():
        for p in paises_list:
            if selected_countries[p]:
                toggle_country(p)

    all_btn = tk.Label(links_frame, text="Todos", font=("Segoe UI", 8, "underline"), bg=CARD_BG_COLOR, fg=TEXT_SECONDARY, cursor="hand2")
    all_btn.pack(side="left")
    all_btn.bind("<Button-1>", lambda e: select_all_countries())

    none_btn = tk.Label(links_frame, text="Ninguno", font=("Segoe UI", 8, "underline"), bg=CARD_BG_COLOR, fg=TEXT_SECONDARY, cursor="hand2")
    none_btn.pack(side="left", padx=12)
    none_btn.bind("<Button-1>", lambda e: select_none_countries())

    # --- 3. PROGRAMACIÓN SEMANAL / TEST AUTOMÁTICO (Card Panel — fondo distinto para resaltar) ---
    SCHED_BG = "#3B1E5E"      # Morado un poco más azulado/oscuro que las demás cards
    SCHED_BORDER = "#8B5FB5"  # Borde más visible
    scheduler_scroll_frame = make_scrollable_tab_container(tabs["scheduler"])

    programacion_card = tk.Frame(scheduler_scroll_frame, bg=SCHED_BG, bd=0, highlightthickness=1, highlightbackground=SCHED_BORDER)
    programacion_card.pack(fill="x", padx=20, pady=(20, 10), ipady=5)

    left_col = tk.Frame(programacion_card, bg=SCHED_BG)
    left_col.pack(side="left", fill="both", expand=True, padx=(15, 10), pady=10)

    right_col = tk.Frame(programacion_card, bg=SCHED_BG)
    right_col.pack(side="right", fill="both", expand=True, padx=(10, 15), pady=10)

    # --- Columna Izquierda ---
    tk.Label(left_col, text="📅 PROGRAMACIÓN DE TEST AUTOMÁTICO", font=("Segoe UI", 9, "bold"), bg=SCHED_BG, fg=TEXT_SECONDARY).pack(anchor="w", pady=(0, 10))

    prog_btns = tk.Frame(left_col, bg=SCHED_BG)

    # --- Columna Derecha ---
    header_ctas = tk.Frame(right_col, bg=SCHED_BG)
    header_ctas.pack(fill="x", pady=(0, 6))

    prog_badge = tk.Label(header_ctas, text="", font=("Segoe UI", 8, "bold"), bg=BUTTON_INACTIVE, fg=TEXT_SECONDARY, padx=10, pady=2)
    prog_badge.pack(side="right")

    btn_config_prog = tk.Button(header_ctas, text=" Configurar automatización", image=get_button_icon("gear_white.png"), compound="left",
                                font=("Segoe UI", 8, "bold"),
                                bg=BUTTON_ACTIVE, fg="white", relief="flat", bd=0, activebackground=BUTTON_HOVER,
                                activeforeground="white", padx=12, pady=3, cursor="hand2",
                                command=lambda: toggle_scheduler_config())
    btn_config_prog.pack(side="right", padx=(0, 8))
    btn_config_prog.bind("<Enter>", lambda e: btn_config_prog.config(bg=BUTTON_HOVER))
    btn_config_prog.bind("<Leave>", lambda e: btn_config_prog.config(bg=BUTTON_ACTIVE))

    prog_hint = tk.Label(right_col,
                         text="📖  Guía de inicio rápido:\n"
                              "① Definí el modo de Mercados y los Dispositivos abajo.  "
                              "② Clic en “Configurar automatización” (arriba) para definir días, horarios y países.  "
                              "③ Clic en “Programar test automático” (al final) para activar la tarea en background.\n"
                              "💡 El test programado lee automáticamente los archivos Excel correspondientes a los dispositivos seleccionados.",
                         font=("Segoe UI", 8, "italic"), bg=SCHED_BG, fg="#C5A9DF", justify="left", anchor="w", wraplength=480)
    prog_hint.pack(anchor="w", pady=(2, 6))

    prog_body = tk.Frame(right_col, bg=SCHED_BG)
    prog_body.pack(fill="x", pady=2)

    prog_desc = tk.Label(prog_body, text="", font=("Segoe UI", 8, "italic"), bg=SCHED_BG, fg="#C5A9DF",
                         justify="left", anchor="w", wraplength=480)
    prog_desc.pack(side="left", fill="x", expand=True)


    prog_state = {"mode": "sin_config"}   # sin_config | configurado | activado
    scheduler_cfg = {"horarios": {}, "paises": [], "dispositivo": "local", "modo_excel": "consecutivo", "navegadores": ["chrome"]}

    sched_mode_btns = {}
    sched_navs_btns = {}

    def select_sched_mode(val):
        scheduler_cfg["modo_excel"] = val
        _refresh_sched_config_ui()
        _refresh_prog()

    def toggle_sched_nav(val):
        cur_disp = scheduler_cfg.get("dispositivo", "local")
        cur_navs = list(scheduler_cfg.get("navegadores", ["chrome"]))

        if val == "lambdatest_mac":
            scheduler_cfg["dispositivo"] = "lambdatest_mac"
            scheduler_cfg["navegadores"] = ["lambdatest_mac"]
        elif val == "lambdatest_android":
            scheduler_cfg["dispositivo"] = "lambdatest_android"
            scheduler_cfg["navegadores"] = ["lambdatest_android"]
        else:  # local chrome/firefox/edge
            if cur_disp in ("lambdatest_android", "lambdatest_mac"):
                scheduler_cfg["dispositivo"] = "local"
                scheduler_cfg["navegadores"] = [val]
            else:
                if val in cur_navs:
                    if len(cur_navs) > 1:
                        cur_navs.remove(val)
                else:
                    cur_navs.append(val)
                scheduler_cfg["navegadores"] = [n for n in cur_navs if n in ("chrome", "firefox", "edge")]

        _refresh_sched_config_ui()
        _refresh_prog()

    def _refresh_sched_config_ui():
        cur_disp = scheduler_cfg.get("dispositivo", "local")
        cur_mode = scheduler_cfg.get("modo_excel", "consecutivo")
        cur_navs = scheduler_cfg.get("navegadores", ["chrome"])

        # 1. Botones de mercados
        is_lt = cur_disp in ("lambdatest_android", "lambdatest_mac")
        for m_key, btn in sched_mode_btns.items():
            btn.config(state="normal", cursor="hand2")
            if m_key == cur_mode:
                btn.config(bg=BUTTON_ACTIVE, fg="white", highlightthickness=1, highlightbackground=ACCENT_COLOR)
            else:
                btn.config(bg=BUTTON_INACTIVE, fg=TEXT_SECONDARY, highlightthickness=0)

        # 2. Botones de navegadores / dispositivos
        for n_key, btn in sched_navs_btns.items():
            if n_key in cur_navs:
                btn.config(bg=BUTTON_ACTIVE, fg="white", highlightthickness=1, highlightbackground=ACCENT_COLOR)
            else:
                btn.config(bg=BUTTON_INACTIVE, fg=TEXT_SECONDARY, highlightthickness=0)

        if is_lt:
            sched_warning_lbl.config(text="⚠️ Si corrés en paralelo con LambdaTest, asegurate de que tu plan soporte suficientes sesiones concurrentes.")
            sched_warning_lbl.pack(fill="x", pady=(4, 0))
        else:
            sched_warning_lbl.pack_forget()

    prog_config = tk.Frame(left_col, bg=SCHED_BG)
    prog_config.pack(fill="x", pady=(4, 6))

    # Columna 1: MERCADOS (Consecutivo / Paralelo)
    col_merc = tk.Frame(prog_config, bg=SCHED_BG)
    col_merc.pack(side="left", padx=(0, 40), anchor="n")
    tk.Label(col_merc, text="MERCADOS", font=("Segoe UI", 8, "bold"), bg=SCHED_BG, fg=TEXT_SECONDARY).pack(anchor="w", pady=(0, 2))
    merc_btn_row = tk.Frame(col_merc, bg=SCHED_BG)
    merc_btn_row.pack(anchor="w")

    # Columna 2: DISPOSITIVOS / NAVEGADORES (Chrome, Firefox, Edge, Mac LT, Android LT)
    col_navs = tk.Frame(prog_config, bg=SCHED_BG)
    col_navs.pack(side="left", anchor="n")
    tk.Label(col_navs, text="DISPOSITIVOS / NAVEGADORES", font=("Segoe UI", 8, "bold"), bg=SCHED_BG, fg=TEXT_SECONDARY).pack(anchor="w", pady=(0, 2))
    navs_btn_row = tk.Frame(col_navs, bg=SCHED_BG)
    navs_btn_row.pack(anchor="w")

    prog_btns.pack(fill="x", pady=(10, 0))

    sched_warning_lbl = tk.Label(right_col, text="", font=("Segoe UI", 8, "italic"), bg=SCHED_BG, fg="#F8C471", justify="left")

    _mode_opts = [
        ("Secuencial", "consecutivo"),
        ("Paralelo", "paralelo")
    ]
    for opt_txt, opt_val in _mode_opts:
        b = tk.Button(merc_btn_row, text=opt_txt, font=("Segoe UI", 8, "bold"), bg=BUTTON_INACTIVE, fg=TEXT_SECONDARY,
                      relief="flat", bd=0, activebackground=BUTTON_HOVER, activeforeground="white",
                      padx=12, pady=4, cursor="hand2")
        b.pack(side="left", padx=2)
        b.config(command=lambda val=opt_val: select_sched_mode(val))
        sched_mode_btns[opt_val] = b

    _nav_opts = [
        ("Chrome", "chrome"),
        ("Firefox", "firefox"),
        ("Edge", "edge"),
        ("Mac LT", "lambdatest_mac"),
        ("Android LT", "lambdatest_android")
    ]
    for opt_txt, opt_val in _nav_opts:
        b = tk.Button(navs_btn_row, text=opt_txt, font=("Segoe UI", 8, "bold"), bg=BUTTON_INACTIVE, fg=TEXT_SECONDARY,
                      relief="flat", bd=0, activebackground=BUTTON_HOVER, activeforeground="white",
                      padx=12, pady=4, cursor="hand2")
        b.pack(side="left", padx=2)
        b.config(command=lambda val=opt_val: toggle_sched_nav(val))
        sched_navs_btns[opt_val] = b



    def _sched_summary():
        dias = len([d for d, h in scheduler_cfg["horarios"].items() if h])
        total = sum(len(v) for v in scheduler_cfg["horarios"].values())
        return dias, total, len(scheduler_cfg["paises"])

    def _mk_prog_btn(text, cmd, bg, fg, hover, icon=None):
        b = tk.Button(prog_btns, text=text, font=("Segoe UI", 8, "bold"), bg=bg, fg=fg, relief="flat", bd=0,
                      activebackground=hover, activeforeground=fg, padx=12, pady=4, cursor="hand2", command=cmd)
        img = get_button_icon(icon) if icon else None
        if img:
            b.config(image=img, compound="left")
        b.bind("<Enter>", lambda e: b.config(bg=hover) if str(b["state"]) == "normal" else None)
        b.bind("<Leave>", lambda e: b.config(bg=bg) if str(b["state"]) == "normal" else None)
        return b
    def _refresh_prog():
        for w in prog_btns.winfo_children():
            w.destroy()
        for w in prog_body.winfo_children():
            w.destroy()
        dias, total, n_p = _sched_summary()
        configured = total > 0 and n_p > 0
        if prog_state["mode"] != "activado":
            prog_state["mode"] = "configurado" if configured else "sin_config"

        if prog_state["mode"] == "activado":
            prog_badge.config(text="✓ Activado", bg=EXECUTE_BG, fg=EXECUTE_FG)
        elif configured:
            prog_badge.config(text="⚙ Configurado", bg=BORDER_COLOR, fg=TEXT_PRIMARY)
        else:
            prog_badge.config(text="🕐 Sin configurar", bg=BUTTON_INACTIVE, fg=TEXT_SECONDARY)

        if configured:
            # Título de estado
            title_txt = "📅 Programado en background:" if prog_state["mode"] == "activado" else "⚙️ Configuración lista:"
            tk.Label(prog_body, text=title_txt, font=("Segoe UI", 8, "bold"), bg=SCHED_BG, fg="white").pack(anchor="w", pady=(0, 4))

            # Contenedor para los horarios individuales (vertical)
            hours_frame = tk.Frame(prog_body, bg=SCHED_BG)
            hours_frame.pack(fill="x", anchor="w", pady=2)

            if total <= 10:
                for day_abbr, day_full in _SCH_DAYS:
                    hours = scheduler_cfg["horarios"].get(day_full) or []
                    if hours:
                        pill = tk.Frame(hours_frame, bg="#4A2666", bd=0, highlightthickness=1, highlightbackground=BORDER_COLOR)
                        pill.pack(anchor="w", pady=2)
                        tk.Label(pill, text=f" {day_abbr.upper()} ", font=("Segoe UI", 8, "bold"), bg="#4A2666", fg=ACCENT_COLOR).pack(side="left", pady=1)
                        hours_str = ", ".join(sorted(hours))
                        tk.Label(pill, text=f"{hours_str} ", font=("Segoe UI", 8), bg="#4A2666", fg="white").pack(side="left", pady=1)
            else:
                pill = tk.Frame(hours_frame, bg="#4A2666", bd=0, highlightthickness=1, highlightbackground=BORDER_COLOR)
                pill.pack(anchor="w", pady=2)
                tk.Label(pill, text=f" 📆 {dias} DÍAS ({total} HORARIOS) ", font=("Segoe UI", 8, "bold"), bg="#4A2666", fg=ACCENT_COLOR).pack(side="left", pady=1)

            # Contenedor de detalles de ejecución (abajo de los horarios)
            details_frame = tk.Frame(prog_body, bg=SCHED_BG)
            details_frame.pack(fill="x", anchor="w", pady=(6, 2))

            # Pill de dispositivo
            cur_disp = scheduler_cfg.get("dispositivo", "local")
            cur_navs = scheduler_cfg.get("navegadores", ["chrome"])
            if cur_disp == "local":
                nav_names = [n.upper() for n in cur_navs]
                disp_txt = f"🖥️ LOCAL ({', '.join(nav_names)})"
            elif cur_disp == "lambdatest_android":
                disp_txt = "🤖 ANDROID LT"
            else:
                disp_txt = "🍏 MAC LT"
            pill_d = tk.Frame(details_frame, bg="#3E2B5A", bd=0, highlightthickness=1, highlightbackground="#9B7FCD")
            pill_d.pack(side="left", padx=(0, 4), pady=2)
            tk.Label(pill_d, text=f" {disp_txt} ", font=("Segoe UI", 8, "bold"), bg="#3E2B5A", fg="#C5A9DF").pack(side="left", pady=1)

            # Pill de modo de ejecución
            cur_mode = scheduler_cfg.get("modo_excel", "consecutivo")
            mode_txt = "SECUENCIAL" if cur_mode == "consecutivo" else "PARALELO"
            pill_m = tk.Frame(details_frame, bg="#2C3E50", bd=0, highlightthickness=1, highlightbackground="#5DADE2")
            pill_m.pack(side="left", padx=4, pady=2)
            tk.Label(pill_m, text=f" {mode_txt} ", font=("Segoe UI", 8, "bold"), bg="#2C3E50", fg="#AED6F1").pack(side="left", pady=1)

            # Pill de países
            paises_abrevs = [_VAL_ABBR.get(p, p) for p in scheduler_cfg["paises"]]
            pstr = ", ".join(paises_abrevs)
            pill_p = tk.Frame(details_frame, bg="#194D33", bd=0, highlightthickness=1, highlightbackground="#82E0AA")
            pill_p.pack(side="left", padx=4, pady=2)
            tk.Label(pill_p, text=" MERCADOS ", font=("Segoe UI", 8, "bold"), bg="#194D33", fg="#82E0AA").pack(side="left", pady=1)
            tk.Label(pill_p, text=f"{pstr} ", font=("Segoe UI", 8), bg="#194D33", fg="white").pack(side="left", pady=1)

            info_txt = " (Se ejecuta en background según los horarios)" if prog_state["mode"] == "activado" else " (Programá el test para activarlo)"
            tk.Label(prog_body, text=info_txt, font=("Segoe UI", 8, "italic"), bg=SCHED_BG, fg="#C5A9DF").pack(anchor="w", pady=2)
        else:
            tk.Label(prog_body, text="Sin configuración. Usá \"Configurar automatización\" para definir días, horarios y países.",
                      font=("Segoe UI", 8, "italic"), bg=SCHED_BG, fg="#C5A9DF").pack(anchor="w", pady=2)

        if prog_state["mode"] == "activado":
            _mk_prog_btn(" Iniciar ahora", _sched_run_now, EXECUTE_BG, EXECUTE_FG, EXECUTE_HOVER, icon="play_green.png").pack(side="left", padx=6)
            _mk_prog_btn(" Desactivar", _sched_deactivate, BUTTON_INACTIVE, TEXT_DELETE, BUTTON_HOVER, icon="stop_coral.png").pack(side="left", padx=6)
        else:
            # Color azul brillante (#3498DB) para resaltar el botón de activar sobre el fondo morado
            bg_color = "#3498DB" if configured else BUTTON_INACTIVE
            fg_color = "white" if configured else TEXT_SECONDARY
            hover_color = "#2980B9" if configured else BUTTON_HOVER
            b = _mk_prog_btn(" Programar test automático", _sched_activate,
                             bg_color, fg_color, hover_color, icon="play_white.png")
            if not configured:
                b.config(state="disabled", cursor="arrow")
            b.pack(side="left", padx=6)

    def _on_sched_save(cfg):
        selected_paises = cfg.get("paises", [])
        selected_navs = scheduler_cfg.get("navegadores", ["chrome"])
        missing_files = []

        for pais in selected_paises:
            for nav in selected_navs:
                if nav == "lambdatest_mac":
                    dev_name = "Mac"
                elif nav == "lambdatest_android":
                    dev_name = "Android"
                elif nav == "firefox":
                    dev_name = "Firefox"
                elif nav == "edge":
                    dev_name = "Edge"
                else:
                    dev_name = "Chrome"

                excel_name = f"Lead_information_Formulario_{pais}_{dev_name}.xlsx"
                excel_path = os.path.join(DATA_DIR, excel_name)
                if not os.path.exists(excel_path):
                    missing_files.append(f"• {pais} ({dev_name}) -> {excel_name}")

        if missing_files:
            missing_str = "\n".join(missing_files)
            messagebox.showerror(
                "Archivos Excel Faltantes",
                f"No se puede guardar la configuración porque no se encuentran los archivos Excel requeridos en la carpeta 'data/':\n\n{missing_str}\n\nPor favor, generá los Excels para estos mercados y navegadores antes de programar la automatización."
            )
            return False

        scheduler_cfg["horarios"] = cfg.get("horarios", {})
        scheduler_cfg["paises"] = selected_paises
        log_message("[INFO] Programación semanal guardada.")
        _refresh_prog()
        return True
    # Panel de configuración (desplegable)
    def _on_config_close():
        config_panel.pack_forget()

    config_panel = _DemoSchedulerPanel(
        scheduler_scroll_frame, 
        on_save=_on_sched_save, 
        initial_config=scheduler_cfg,
        on_close=_on_config_close
    )

    def toggle_scheduler_config():
        if config_panel.winfo_ismapped():
            config_panel.pack_forget()
        else:
            # Sincronizar la config actual
            config_panel._schedule = {k: list(v) for k, v in scheduler_cfg.get("horarios", {}).items()}
            config_panel._countries = list(scheduler_cfg.get("paises", []))
            
            # Actualizar la UI del panel
            config_panel._update_day_buttons()
            config_panel._update_country_count()
            config_panel._update_toggle_all_label()
            for country, var in config_panel._country_vars.items():
                var.set(country in config_panel._countries)
            config_panel._close_hours_panel()
            config_panel._update_footer()
            
            config_panel.pack(fill="x", padx=20, pady=(0, 20))
            
            # Auto-scroll al final de la pestaña para visualizar el panel abierto
            try:
                scheduler_scroll_frame.update_idletasks()
                canvas = scheduler_scroll_frame.master
                canvas.yview_moveto(1.0)
            except Exception:
                pass

    def _sched_activate():
        dias, total, n_p = _sched_summary()
        if not (total > 0 and n_p > 0):
            messagebox.showwarning("Programación", "Configurá días, horarios y países antes de programar.")
            return
        
        disp = scheduler_cfg.get("dispositivo", "local")
        navs = scheduler_cfg.get("navegadores", ["chrome"])

        try:
            guardar_programacion({
                "tipo": "semanal",
                "horarios": {k: v for k, v in scheduler_cfg.get("horarios", {}).items() if v},
                "paises": list(scheduler_cfg.get("paises", [])),
                "navegadores": navs,
                "viewports": ["fullscreen"],
                "dispositivo": disp,
                "modo_excel": scheduler_cfg.get("modo_excel", "consecutivo"),
                "modo_mercados": scheduler_cfg.get("modo_excel", "consecutivo"),
            })
        except Exception as e:
            log_message(f"[ERROR] No se pudo guardar la programación: {e}")
        prog_state["mode"] = "activado"
        log_message("[SUCCESS] Test automático programado y activado (json/programacion_test.json).")
        _refresh_prog()
        messagebox.showinfo("Programación", "✓ Test automático activado. Se ejecutará según los horarios (con la app abierta).")

    def _sched_deactivate():
        try:
            limpiar_programacion()
        except Exception:
            pass
        prog_state["mode"] = "configurado"
        log_message("[INFO] Test automático desactivado.")
        _refresh_prog()

    def _sched_run_now():
        log_message("[INFO] Iniciando test programado ahora...")
        execute_send_leads(scheduled=True)

    # Cargar programación persistida al iniciar
    try:
        _existing = cargar_programacion()
        if _existing and _existing.get("tipo") == "semanal":
            scheduler_cfg["horarios"] = _existing.get("horarios", {})
            scheduler_cfg["paises"] = _existing.get("paises", [])
            scheduler_cfg["dispositivo"] = _existing.get("dispositivo", "local")
            scheduler_cfg["modo_excel"] = _existing.get("modo_mercados") or _existing.get("modo_excel", "consecutivo")
            scheduler_cfg["navegadores"] = _existing.get("navegadores", ["chrome"])
            prog_state["mode"] = "activado"
    except Exception:
        pass

    _refresh_sched_config_ui()
    _refresh_prog()

    # Monitor semanal: dispara la ejecución cuando llega un horario programado (app abierta)
    _SCHED_TRIG_PATH = os.path.join(JSON_DIR, "scheduler_triggered.json")
    _sched_triggered = {}
    _DAYS_ES = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]

    def _sched_load_triggered():
        try:
            import json
            with open(_SCHED_TRIG_PATH, encoding="utf-8") as f:
                raw = json.load(f)
            return {tuple(k.split("|")): _date.fromisoformat(v) for k, v in raw.items()}
        except Exception:
            return {}

    def _sched_save_triggered():
        try:
            import json
            with open(_SCHED_TRIG_PATH, "w", encoding="utf-8") as f:
                json.dump({f"{k[0]}|{k[1]}": v.isoformat() for k, v in _sched_triggered.items()}, f)
        except Exception:
            pass

    def _sched_monitor():
        _sched_triggered.update(_sched_load_triggered())
        first = True
        while True:
            try:
                if prog_state["mode"] == "activado" and scheduler_cfg.get("horarios"):
                    ahora = _dt.now()
                    dia = _DAYS_ES[ahora.weekday()]
                    now_min = ahora.hour * 60 + ahora.minute
                    for hora in list(scheduler_cfg["horarios"].get(dia, [])):
                        try:
                            hh, mm = int(hora[:2]), int(hora[3:5])
                        except Exception:
                            continue
                        start = hh * 60 + mm
                        if start <= now_min < start + 15:
                            key = (dia, hora)
                            if _sched_triggered.get(key) != ahora.date():
                                _sched_triggered[key] = ahora.date()
                                _sched_save_triggered()
                                if not first:
                                    root.after(0, lambda: execute_send_leads(scheduled=True))
                first = False
            except Exception:
                pass
            time.sleep(5)

    threading.Thread(target=_sched_monitor, daemon=True).start()

    # --- 4. DATOS POR PAÍS (Card Panel Table) ---
    datos_card = tk.Frame(leads_scroll_frame, bg=CARD_BG_COLOR, bd=0, highlightthickness=1, highlightbackground=BORDER_COLOR)
    datos_card.pack(fill="x", expand=False, ipady=4)

    d_header = tk.Frame(datos_card, bg=CARD_BG_COLOR)
    d_header.pack(fill="x", padx=15, pady=(6, 4))

    tk.Label(d_header, text="📄 DATOS POR PAÍS", font=("Segoe UI", 10, "bold"), bg=CARD_BG_COLOR, fg=TEXT_PRIMARY).pack(side="left")
    tk.Label(d_header, text="ℹ Acá verás una previsualización del Excel del primer dispositivo que encuentre el sistema.", font=("Segoe UI", 8, "italic"), bg=CARD_BG_COLOR, fg="#C5A9DF").pack(side="left", padx=15)

    p_tabs_row = tk.Frame(datos_card, bg=CARD_BG_COLOR)
    p_tabs_row.pack(fill="x", padx=15, pady=(0, 6))

    p_tab_buttons = {}

    def select_p_tab(btn, name):
        for _n, _b in p_tab_buttons.items():
            _b.config(bg=BUTTON_INACTIVE, fg=TEXT_SECONDARY, highlightthickness=1, highlightbackground=BUTTON_INACTIVE)
        btn.config(bg=BUTTON_ACTIVE, fg="white", highlightthickness=1, highlightbackground=ACCENT_COLOR)
        active_p_tab[0] = name
        update_table_data(name)
        # Forzar actualización en Tab 3 por si cambió el país activo
        try:
            update_excel_calculation()
        except Exception:
            pass

    for idx, pais in enumerate(paises_list):
        has_leads = (pais in ["Argentina", "Bolivia"])
        leads_count = " 3" if pais == "Argentina" else (" 1" if pais == "Bolivia" else "")
        suffix = f" ({leads_count.strip()})" if has_leads else ""
        
        btn_bg = BUTTON_ACTIVE if pais == "Argentina" else BUTTON_INACTIVE
        btn_fg = "white" if pais == "Argentina" else TEXT_SECONDARY
        btn_hb = ACCENT_COLOR if pais == "Argentina" else BUTTON_INACTIVE

        b = tk.Button(p_tabs_row, text=f"{pais}{suffix}", font=("Segoe UI", 8, "bold"), bg=btn_bg, fg=btn_fg,
                      relief="flat", bd=0, activebackground=BUTTON_HOVER, activeforeground="white",
                      highlightthickness=1, highlightbackground=btn_hb,
                      padx=8, pady=3, cursor="hand2")
        b.pack(side="left", padx=1)
        p_tab_buttons[pais] = b
        b.config(command=lambda btn=b, n=pais: select_p_tab(btn, n))

    # Selector de qué Excel (dispositivo) previsualizar dentro del país activo
    current_preview_path = [None]
    _preview_dev_paths = {}  # label visible -> ruta del Excel (del país actual)
    preview_device_var = tk.StringVar()

    def _available_device_excels(pais):
        t3 = bool(var_t3.get())
        t3_tag = "_T3" if t3 else ""
        specs = [("Chrome", "Chrome"), ("Firefox", "Firefox"), ("Edge", "Edge"),
                 ("Mac LT", "Mac"), ("Android LT", "Android"), ("Genérico (compartido)", "Generico")]
        out = []
        for label, suf in specs:
            p = os.path.join(DATA_DIR, f"Lead_information_Formulario_{pais}_{suf}{t3_tag}.xlsx")
            if os.path.exists(p):
                out.append((label, p))
        return out

    tk.Label(p_tabs_row, text="Excel a revisar:", font=("Segoe UI", 8, "bold"),
             bg=CARD_BG_COLOR, fg=TEXT_SECONDARY).pack(side="right", padx=(8, 4))
    preview_device_combo = ttk.Combobox(p_tabs_row, textvariable=preview_device_var,
                                        state="readonly", width=20, style="TCombobox")
    preview_device_combo.pack(side="right")

    def _on_preview_device_change(_e=None):
        label = preview_device_var.get()
        path = _preview_dev_paths.get(label)
        if path:
            update_table_data(active_p_tab[0], path=path)
    preview_device_combo.bind("<<ComboboxSelected>>", _on_preview_device_change)

    # Controles de Tabla
    tbl_ctrl_row = tk.Frame(datos_card, bg=CARD_BG_COLOR)
    tbl_ctrl_row.pack(fill="x", padx=15, pady=(0, 6))

    # Etiqueta para avisos efímeros (Eliminar, Agregar, etc.) al lado de los botones
    tbl_msg_lbl = tk.Label(tbl_ctrl_row, text="", font=("Segoe UI", 8, "bold"), bg=CARD_BG_COLOR)

    # Control de timer para mensaje efímero
    tbl_msg_timer_id = [None]

    def show_table_msg(text, color):
        # Cancelar el timer previo si existe
        if tbl_msg_timer_id[0] is not None:
            try:
                root.after_cancel(tbl_msg_timer_id[0])
            except Exception:
                pass
            tbl_msg_timer_id[0] = None

        tbl_msg_lbl.pack_forget()
        tbl_msg_lbl.config(text=text, fg=color)
        tbl_msg_lbl.pack(side="left", padx=15)
        
        # Desaparecer el mensaje tras 10 segundos (10000 ms)
        def hide_msg():
            tbl_msg_lbl.pack_forget()
            tbl_msg_timer_id[0] = None

        tbl_msg_timer_id[0] = root.after(10000, hide_msg)

    # Comandos (Excel real por país)
    def cmd_agregar():
        _close_cell_editor()
        n = len(tree["columns"]) or 1
        iid = tree.insert("", "end", values=[""] * n)
        _stripe_rows()
        tree.selection_set(iid)
        tree.focus(iid)
        tree.see(iid)
        tree.focus_set()
        show_table_msg("✓ Fila agregada (resaltada abajo). Completá con doble clic y Guardá.", TEXT_ADD)
        log_message("[INFO] Fila vacía agregada a la tabla.")

    def cmd_eliminar():
        _close_cell_editor()
        selected = tree.selection()
        if selected:
            for item in selected:
                tree.delete(item)
            _stripe_rows()
            show_table_msg(f"🗑 {len(selected)} fila(s) eliminada(s). Guardá para persistir.", TEXT_DELETE)
            log_message(f"[INFO] {len(selected)} fila(s) eliminada(s) de la tabla.")
        else:
            show_table_msg("⚠ Seleccioná una o varias filas para eliminar.", WARN_TEXT)

    def cmd_clonar():
        _close_cell_editor()
        selected = tree.selection()
        if selected:
            last = None
            for item in selected:
                last = tree.insert("", "end", values=tree.item(item, "values"))
            _stripe_rows()
            if last:
                tree.see(last)
            show_table_msg(f"📋 {len(selected)} fila(s) clonada(s). Guardá para persistir.", TEXT_SECONDARY)
            log_message(f"[INFO] {len(selected)} fila(s) duplicada(s) en la tabla.")
        else:
            show_table_msg("⚠ Seleccioná una o varias filas para clonar.", WARN_TEXT)

    def cmd_actualizar_tbl():
        update_table_data(active_p_tab[0])
        show_table_msg("🔄 Datos recargados desde el Excel.", TEXT_SAVE)
        log_message(f"[INFO] Datos de {active_p_tab[0]} recargados desde el archivo.")

    def cmd_guardar_tbl():
        pais = active_p_tab[0]
        path = current_preview_path[0] or _excel_path_for(pais)
        cols = list(tree["columns"])
        rows = [list(tree.item(i, "values")) for i in tree.get_children()]
        try:
            import pandas as pd
            os.makedirs(DATA_DIR, exist_ok=True)
            pd.DataFrame(rows, columns=cols).astype(str).to_excel(path, index=False)
            show_table_msg(f"💾 Guardado {len(rows)} fila(s) en {os.path.basename(path)}.", TEXT_SAVE)
            log_message(f"[SUCCESS] Guardadas {len(rows)} filas en {path}.")
            _refresh_tab_count(pais, len(rows))
        except PermissionError:
            messagebox.showerror("Guardar Leads", "El Excel está abierto en otro programa. Cerralo y reintentá.")
        except Exception as e:
            messagebox.showerror("Guardar Leads", f"No se pudo guardar:\n{e}")

    def cmd_abrir_excel():
        pais = active_p_tab[0]
        path = current_preview_path[0] or _excel_path_for(pais)
        try:
            if not os.path.exists(path):
                import pandas as pd
                os.makedirs(DATA_DIR, exist_ok=True)
                pd.DataFrame(columns=build_excel_columns_for_country(pais)).to_excel(path, index=False)
            os.startfile(path)
            show_table_msg("📂 Excel abierto.", TEXT_EXCEL)
            log_message(f"[INFO] Abriendo {path}.")
        except Exception as e:
            messagebox.showerror("Abrir Excel", f"No se pudo abrir el Excel:\n{e}")

    make_icon_btn(tbl_ctrl_row, "+ Agregar", TEXT_ADD, command=cmd_agregar)
    make_icon_btn(tbl_ctrl_row, "🗑 Eliminar", TEXT_DELETE, command=cmd_eliminar)
    make_icon_btn(tbl_ctrl_row, "📋 Clonar", TEXT_SECONDARY, command=cmd_clonar)
    make_icon_btn(tbl_ctrl_row, "🔄 Actualizar", TEXT_SECONDARY, command=cmd_actualizar_tbl)
    make_icon_btn(tbl_ctrl_row, "💾 Guardar", TEXT_SAVE, command=cmd_guardar_tbl)
    make_icon_btn(tbl_ctrl_row, "📂 Abrir Excel", TEXT_EXCEL, command=cmd_abrir_excel)

    # Frame contenedor con barras de scroll (TScrollbar estilizados)
    tree_frame = tk.Frame(datos_card, bg=CARD_BG_COLOR)
    tree_frame.pack(fill="x", expand=False, padx=15, pady=2)

    columns = ("url_landing", "url_form", "modelo", "nombre", "apellido", "email", "telefono", "documento", "direccion", "ciudad", "estado", "cod_postal", "comentarios")
    tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=5, selectmode="extended", style="Treeview")
    
    tree.heading("url_landing", text="URL LANDING")
    tree.heading("url_form", text="URL FORM")
    tree.heading("modelo", text="MODELO")
    tree.heading("nombre", text="NOMBRE")
    tree.heading("apellido", text="APELLIDO")
    tree.heading("email", text="EMAIL")
    tree.heading("telefono", text="TELÉFONO")
    tree.heading("documento", text="DOCUMENTO")
    tree.heading("direccion", text="DIRECCIÓN")
    tree.heading("ciudad", text="CIUDAD")
    tree.heading("estado", text="ESTADO")
    tree.heading("cod_postal", text="CÓDIGO POSTAL")
    tree.heading("comentarios", text="COMENTARIOS")

    tree.column("url_landing", width=320, minwidth=150, stretch=False)
    tree.column("url_form", width=320, minwidth=130, stretch=False)
    tree.column("modelo", width=100, minwidth=60, stretch=False)
    tree.column("nombre", width=120, minwidth=80, stretch=False)
    tree.column("apellido", width=120, minwidth=80, stretch=False)
    tree.column("email", width=180, minwidth=120, stretch=False)
    tree.column("telefono", width=120, minwidth=80, stretch=False)
    tree.column("documento", width=120, minwidth=90, stretch=False)
    tree.column("direccion", width=200, minwidth=110, stretch=False)
    tree.column("ciudad", width=120, minwidth=80, stretch=False)
    tree.column("estado", width=120, minwidth=80, stretch=False)
    tree.column("cod_postal", width=100, minwidth=80, stretch=False)
    tree.column("comentarios", width=250, minwidth=130, stretch=False)

    # Scrollbars Horizontal y Vertical para la tabla Excel (Diseño premium unificado)
    v_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview, style="TScrollbar")
    h_scrollbar = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview, style="TScrollbar")
    tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
    
    h_scrollbar.pack(side="bottom", fill="x")
    v_scrollbar.pack(side="right", fill="y")
    tree.pack(side="top", fill="x", expand=False)

    # Edición inline de celdas (doble clic) + aviso de "editando sin guardar"
    _active_editor = [None]

    def _close_cell_editor(_e=None):
        ed = _active_editor[0]
        _active_editor[0] = None
        if ed is not None:
            try:
                ed.destroy()
            except Exception:
                pass

    def _edit_cell(event):
        _close_cell_editor()  # cerrar cualquier editor previo abierto
        if tree.identify_region(event.x, event.y) != "cell":
            return
        row = tree.identify_row(event.y)
        col = tree.identify_column(event.x)
        if not row or not col:
            return
        bbox = tree.bbox(row, col)
        if not bbox:
            return
        col_index = int(col[1:]) - 1
        x, y, w, h = bbox
        current_val = tree.item(row, "values")[col_index]
        editor = tk.Entry(tree, bg=ENTRY_BG, fg="white", insertbackground="white", bd=0,
                          relief="flat", highlightthickness=1, highlightbackground=ACCENT_COLOR, font=("Segoe UI", 9))
        editor.place(x=x, y=y, width=w, height=h)
        editor.insert(0, current_val)
        editor.focus_set()
        editor.select_range(0, "end")
        _active_editor[0] = editor

        def commit(_e=None):
            if _active_editor[0] is not editor:
                return
            _active_editor[0] = None
            try:
                if tree.exists(row):
                    vals = list(tree.item(row, "values"))
                    if col_index < len(vals) and vals[col_index] != editor.get():
                        vals[col_index] = editor.get()
                        tree.item(row, values=vals)
                        show_table_msg("✏ Estás editando el Excel — recordá Guardar para no perder los cambios.", TEXT_EXCEL)
            except Exception:
                pass
            try:
                editor.destroy()
            except Exception:
                pass

        editor.bind("<Return>", commit)
        editor.bind("<FocusOut>", commit)
        editor.bind("<Escape>", lambda e: _close_cell_editor())

    tree.bind("<Double-1>", _edit_cell)
    # Cerrar el editor si se scrollea o cambia el tamaño (evita que quede flotando)
    tree.bind("<MouseWheel>", lambda e: _close_cell_editor())

    # Rayado alterno para distinguir filas (evita que una fila vacía se pierda con el fondo)
    tree.tag_configure("odd", background="#573083")
    tree.tag_configure("even", background="#3B1C52")

    def _stripe_rows():
        for i, it in enumerate(tree.get_children()):
            tree.item(it, tags=("odd" if i % 2 else "even",))

    def _refresh_tab_count(pais, n):
        btn = p_tab_buttons.get(pais)
        if btn:
            btn.config(text=f"{pais} ({n})" if n else pais)

    # Carga real del Excel del país (columnas dinámicas desde el archivo)
    def update_table_data(pais, path=None):
        _close_cell_editor()

        # Repoblar el selector de dispositivo con los Excels disponibles del país
        avail = _available_device_excels(pais)
        _preview_dev_paths.clear()
        for _lbl, _p in avail:
            _preview_dev_paths[_lbl] = _p

        if not avail:
            preview_device_combo.config(state="disabled", values=[])
            preview_device_var.set("")
        else:
            preview_device_combo.config(state="readonly", values=[_lbl for _lbl, _ in avail])

        if path is None:
            # Mantener el dispositivo elegido si sigue disponible; si no, el primero
            cur = preview_device_var.get()
            if cur in _preview_dev_paths:
                path = _preview_dev_paths[cur]
            elif avail:
                path = avail[0][1]
                preview_device_var.set(avail[0][0])
            else:
                path = _excel_path_for(pais)
        else:
            for _lbl, _p in avail:
                if _p == path:
                    preview_device_var.set(_lbl)
                    break
        current_preview_path[0] = path

        df = None
        try:
            import pandas as pd
            if os.path.exists(path):
                df = pd.read_excel(path, dtype=str).fillna("")
        except Exception as e:
            log_message(f"[ERROR] No se pudo leer el Excel de {pais}: {e}")
            df = None

        if df is not None and len(df.columns) > 0:
            cols = [str(c) for c in df.columns]
        else:
            cols = build_excel_columns_for_country(pais)

        tree.config(columns=cols)
        for c in cols:
            cl = c.lower()
            wide = any(k in cl for k in ("url", "form", "coment", "direcc"))
            tree.heading(c, text=c.upper())
            tree.column(c, width=300 if wide else 120, minwidth=70, stretch=False, anchor="w")

        for item in tree.get_children():
            tree.delete(item)

        n = 0
        if df is not None:
            for _, row in df.iterrows():
                tree.insert("", "end", values=[str(row[c]) for c in df.columns])
                n += 1

        if not avail:
            vals = ["No hay excels creados para este país. Por favor, genéralos en la pestaña 'Generar Excels con Datos'."] + [""] * (len(cols) - 1)
            tree.insert("", "end", values=vals)

        _stripe_rows()
        _refresh_tab_count(pais, n)

    update_table_data("Argentina")

    # Contadores reales en las pestañas de país (según filas del Excel)
    for _pais in paises_list:
        try:
            _p = _excel_path_for(_pais)
            if _pais != "Argentina" and os.path.exists(_p):
                import pandas as pd
                _refresh_tab_count(_pais, len(pd.read_excel(_p, dtype=str)))
        except Exception:
            pass

    # Estado de ejecución (para que la X minimice en vez de cerrar mientras corre).
    _exec_state = {"running": False}

    # Ejecución real de leads (desktop + LambdaTest) con el modal de progreso
    def execute_send_leads(scheduled=False):
        if not BACKEND_OK:
            messagebox.showerror("Ejecutar", f"El backend no está disponible.\n{_BACKEND_IMPORT_ERROR}")
            return

        # Sufijos de dispositivo ↔ tipo de ejecución (para localizar los Excels generados)
        # (suffix Excel, dtype, browser, key en selected_disp)
        _DEVICE_SUFFIX = [
            ("Chrome", "desktop", "chrome", "chrome"),
            ("Firefox", "desktop", "firefox", "firefox"),
            ("Edge", "desktop", "edge", "edge"),
            ("Mac", "mac", None, "mac lt"),
            ("Android", "android", None, "android lt"),
        ]

        # Formularios T3 2.0 (Adobe AEM): usar los Excels con nombre …_T3.xlsx
        t3 = bool(var_t3.get())

        def _sessions_for(pais):
            """Una sesión por dispositivo tildado. Cada dispositivo (desktop y LT)
            usa su propio Excel (…_Chrome/_Firefox/_Edge/_Mac/_Android[_T3].xlsx), que es
            además donde LambdaTest guarda los resultados. No hay fallback: si falta
            el Excel del dispositivo, se detecta luego y no ejecuta."""
            shared = excel_mode_holder[0] == "compartido"
            gpath = _generic_excel_path_for(pais, t3)
            out = []
            for suffix, dtype, browser, key in _DEVICE_SUFFIX:
                if not selected_disp.get(key):
                    continue
                path = gpath if shared else os.path.join(DATA_DIR, _lead_excel_name(pais, suffix, t3))
                out.append({"pais": pais, "dtype": dtype, "browser": browser, "device": suffix, "excel": path})
            return out

        if scheduled:
            disp_sched = scheduler_cfg.get("dispositivo", "local")
            p_mode = scheduler_cfg.get("modo_excel", "consecutivo")
            paises_run = list(scheduler_cfg.get("paises", [])) or [active_p_tab[0]]
            
            market_jobs = []
            for p in paises_run:
                sessions = []
                if disp_sched == "lambdatest_android":
                    sessions.append({
                        "pais": p, "dtype": "android", "browser": None, "device": "Android",
                        "excel": os.path.join(DATA_DIR, _lead_excel_name(p, "Android", t3))
                    })
                elif disp_sched == "lambdatest_mac":
                    sessions.append({
                        "pais": p, "dtype": "mac", "browser": None, "device": "Mac",
                        "excel": os.path.join(DATA_DIR, _lead_excel_name(p, "Mac", t3))
                    })
                else: # local
                    navs = scheduler_cfg.get("navegadores", []) or ["chrome"]
                    for nav in navs:
                        suffix = "Chrome" if nav == "chrome" else "Firefox" if nav == "firefox" else "Edge"
                        sessions.append({
                            "pais": p, "dtype": "desktop", "browser": nav, "device": suffix,
                            "excel": os.path.join(DATA_DIR, _lead_excel_name(p, suffix, t3))
                        })
                market_jobs.append((p, sessions))
                
            mercados_par = (p_mode == "paralelo") and (len(market_jobs) > 1)
            excels_par = (p_mode == "paralelo")
        else:
            if not any(selected_disp.values()):
                messagebox.showwarning("Ejecutar", "Seleccioná al menos un dispositivo / navegador antes de ejecutar.")
                return
            paises_run = [p for p in paises_list if selected_countries.get(p)] or [active_p_tab[0]]
            market_jobs = [(p, _sessions_for(p)) for p in paises_run]
            
            mercados_par = (mercados_mode[0] == "paralelo") and (len(market_jobs) > 1)
            excels_par = (excels_mode[0] == "paralelo")

        total_sessions = sum(len(js) for _, js in market_jobs)
        if total_sessions == 0:
            messagebox.showwarning("Ejecutar", "No hay Excels para ejecutar. Generá datos o seleccioná un dispositivo.")
            return

        mercados_par = (mercados_mode[0] == "paralelo") and (len(market_jobs) > 1)
        excels_par = (excels_mode[0] == "paralelo")

        # Modo "una sesión por URL": expande cada fila de cada Excel en su propia sesión
        url_par = (not scheduled) and bool(var_url_parallel.get())
        try:
            url_max = max(1, min(20, int(url_max_var.get())))
        except Exception:
            url_max = 6

        def _url_sessions():
            import pandas as pd
            tmpdir = os.path.join(_APP_BASE, "temporales")
            os.makedirs(tmpdir, exist_ok=True)
            out = []
            for _pais, sessions in market_jobs:
                for s in sessions:
                    if s["dtype"] != "desktop" or not s["excel"] or not os.path.exists(s["excel"]):
                        out.append(s)  # LT o sin Excel → sesión entera
                        continue
                    try:
                        df = pd.read_excel(s["excel"], dtype=str).fillna("")
                    except Exception:
                        out.append(s)
                        continue
                    if len(df) == 0:
                        continue
                    for i in range(len(df)):
                        tmp = os.path.join(tmpdir, f"_url_{_pais}_{s['device']}_{i + 1}.xlsx")
                        try:
                            df.iloc[[i]].to_excel(tmp, index=False)
                        except Exception:
                            continue
                        out.append({"pais": _pais, "dtype": "desktop", "browser": s["browser"],
                                    "device": f"{s['device']}·URL{i + 1}", "excel": tmp})
            return out

        if url_par:
            flat_sessions = _url_sessions()
            total_sessions = len(flat_sessions)
            if total_sessions == 0:
                messagebox.showwarning("Ejecutar", "No hay URLs para ejecutar.")
                return
            active_sessions_list = flat_sessions
        else:
            active_sessions_list = []
            for _pais, js in market_jobs:
                active_sessions_list.extend(js)

        # Validar Excels ANTES de arrancar: cada dispositivo seleccionado (desktop y LT)
        # debe tener su propio Excel. Si falta alguno, avisar y no ejecutar (sin modal).
        faltantes = []
        faltantes_sessions = []  # sesiones con Excel inexistente (para poder crearlos vacíos)
        for s in active_sessions_list:
            _ex = s["excel"]
            _nombre = os.path.basename(_ex) if _ex else \
                f"Lead_information_Formulario_{s['pais']}_{s['device']}.xlsx"
            if not _ex or not os.path.exists(_ex):
                faltantes.append(f"• {s['pais']} · {s['device']}: {_nombre} (no existe)")
                faltantes_sessions.append(s)
                continue
            # El Excel existe: validar que tenga al menos una fila con datos (no arrancar vacío).
            try:
                import pandas as _pd
                _df = _pd.read_excel(_ex, dtype=str, keep_default_na=False)
                _tiene = (not _df.empty) and any(
                    any(str(v).strip() for v in _r.values) for _, _r in _df.iterrows()
                )
            except Exception:
                _tiene = True  # si no se pudo leer (abierto/corrupto), no bloquear por esto
            if not _tiene:
                faltantes.append(f"• {s['pais']} · {s['device']}: {_nombre} (vacío / sin leads)")
        if faltantes:
            # Dos opciones: crear los Excel vacíos con las columnas del país (respeta el path
            # _T3 si corresponde), o cerrar. Solo se pueden crear los que NO existen.
            _creables = []
            _seen_paths = set()
            for s in faltantes_sessions:
                _p = s.get("excel")
                if _p and _p not in _seen_paths and "temporales" not in _p:
                    _seen_paths.add(_p)
                    _creables.append(s)
            _msg = ("No se encontró un Excel válido (inexistente o vacío) para el/los dispositivo(s) seleccionado(s):\n\n"
                    + "\n".join(faltantes))
            if _creables:
                _msg += ("\n\n¿Querés CREAR el/los Excel vacío(s) con las columnas del país (para completarlos a mano)?\n\n"
                         "Sí = crear los Excel vacíos    ·    No = cerrar")
                if messagebox.askyesno("Error Excel", _msg):
                    _creados = []
                    for s in _creables:
                        try:
                            _cols = build_excel_columns_for_country(s["pais"])
                            import pandas as _pd
                            os.makedirs(os.path.dirname(s["excel"]) or DATA_DIR, exist_ok=True)
                            _pd.DataFrame(columns=_cols).to_excel(s["excel"], index=False)
                            _creados.append(os.path.basename(s["excel"]))
                        except Exception as _ce:
                            log_message(f"[ERROR] No se pudo crear {s.get('excel')}: {_ce}")
                    if _creados:
                        messagebox.showinfo(
                            "Excel creados",
                            "Se crearon vacíos (solo encabezados). Completá al menos un lead "
                            "(o generá datos) y volvé a ejecutar:\n\n"
                            + "\n".join("• " + n for n in _creados))
            else:
                messagebox.showerror("Error Excel", _msg
                                     + "\n\nGenerá los datos (con al menos un lead) antes de ejecutar.")
            return

        import pandas as pd
        total_leads = 0
        for idx, s in enumerate(active_sessions_list):
            s["sess_id"] = idx
            rc = 0
            if s["dtype"] == "desktop" and s["excel"] and os.path.exists(s["excel"]):
                try:
                    rc = len(pd.read_excel(s["excel"]))
                except Exception:
                    rc = 1
            else:
                rc = 1
            s["rows_count"] = max(1, rc)
            total_leads += s["rows_count"]

        # Persistir email + enviar_mail en config_global (lo lee el backend de email)
        enviar_mail = bool(var_enviar_email.get())
        _email_modo = var_modo_email.get()  # "por_pais" | "consolidado"
        dest = email_entry.get().strip()
        try:
            _cfg = cargar_config_global()
            _cfg["email_destinatario"] = dest
            _cfg["enviar_mail"] = enviar_mail
            # Opt-in real: sólo adjunta lo que el usuario tildó (no forzar True por default).
            _cfg["adjuntar_resultados"] = bool(var_adjuntar_res.get())
            _cfg["adjuntar_screenshots"] = bool(var_adjuntar_ss.get())
            guardar_config_global(_cfg)
        except Exception:
            pass

        background = not bool(var_ver_navegador.get())  # ver navegador → visible
        stop_event = threading.Event()

        if not scheduled:
            btn_enviar.config(state="disabled", text=" EN CURSO...", bg=BUTTON_INACTIVE, fg=TEXT_SECONDARY)
        log_message(f"[INFO] Iniciando ejecución {'programada' if scheduled else 'manual'}: "
                    f"{len(market_jobs)} mercado(s) · {total_sessions} sesión(es) · "
                    f"mercados={'paralelo' if mercados_par else 'consecutivo'} · "
                    f"excels={'paralelo' if excels_par else 'consecutivo'}.")

        # Crear modal centrado. El alto se calcula según cuántos mercados hay que mostrar:
        # con alto fijo, a partir del 5º mercado las barras quedaban fuera de la ventana.
        modal = tk.Toplevel(root)
        modal.overrideredirect(True) # Quitar bordes de Windows
        modal_width = 520
        _n_mercados = max(len({_s["pais"] for _s in active_sessions_list}), 1)
        # 260 = encabezado + pastillas + aviso + espacio del resumen final; 46 = alto de cada barra
        modal_height = 260 + 46 * _n_mercados
        modal_height = max(370, min(modal_height, int(root.winfo_screenheight() * 0.85)))
        MODAL_BG = "#231830"
        MODAL_PILL_BG = "#38234D"
        
        # Centrar relativo a la app
        px = root.winfo_rootx() + (root.winfo_width() - modal_width) // 2
        py = root.winfo_rooty() + (root.winfo_height() - modal_height) // 2
        modal.geometry(f"{modal_width}x{modal_height}+{px}+{py}")
        modal.configure(bg=MODAL_BG, bd=1, highlightthickness=1, highlightbackground=BORDER_COLOR)

        # Al cerrar el modal por CUALQUIER motivo, el botón vuelve a su estado original.
        def _reset_exec_btn(_e=None):
            if scheduled:
                return
            try:
                btn_enviar.config(state="normal", text=" EJECUTAR ENVÍO",
                                  bg=EXECUTE_BG, fg=EXECUTE_FG, cursor="hand2")
                refresh_execute_state()
            except Exception:
                pass
        modal.bind("<Destroy>", lambda e: _reset_exec_btn() if str(e.widget) == str(modal) else None, add="+")

        def on_cerrar():
            _reset_exec_btn()
            modal.destroy()  # el handler <Destroy> restaura el botón Ejecutar
            log_message("[INFO] Ventana de ejecución cerrada.")

        # Modal real: bloquea la interacción con la interfaz de atrás mientras se ejecuta
        modal.transient(root)
        modal.attributes("-topmost", False)
        modal.lift()
        # No usamos grab_set() para permitir que el usuario minimice o cierre la ventana principal desde la barra de título de Windows.
        # En su lugar, deshabilitamos la interacción con el área cliente del main window agregando BlockTag a sus widgets.
        try:
            modal.focus_set()
        except Exception:
            pass

        def set_event_blocking(parent, block):
            for child in parent.winfo_children():
                if child == modal or str(child).startswith(str(modal)):
                    continue
                tags = list(child.bindtags())
                if block:
                    if "BlockTag" not in tags:
                        child.bindtags(("BlockTag",) + tuple(tags))
                else:
                    if "BlockTag" in tags:
                        new_tags = tuple(t for t in tags if t != "BlockTag")
                        child.bindtags(new_tags)
                set_event_blocking(child, block)

        def block_evt(e):
            if modal.winfo_exists():
                modal.lift()
                modal.focus_set()
            return "break"

        root.bind_class("BlockTag", "<Button-1>", block_evt)
        root.bind_class("BlockTag", "<ButtonRelease-1>", lambda e: "break")
        root.bind_class("BlockTag", "<Double-Button-1>", lambda e: "break")
        root.bind_class("BlockTag", "<B1-Motion>", lambda e: "break")
        root.bind_class("BlockTag", "<Enter>", lambda e: "break")
        root.bind_class("BlockTag", "<Leave>", lambda e: "break")
        root.bind_class("BlockTag", "<Motion>", lambda e: "break")
        root.bind_class("BlockTag", "<Key>", lambda e: "break")
        root.bind_class("BlockTag", "<FocusIn>", lambda e: block_evt(e))
        set_event_blocking(root, True)

        def on_close_modal():
            if not modal.winfo_exists():
                return
            if "Ejecutando" in title_lbl.cget("text"):
                if messagebox.askyesno("Detener ejecución", "¿Querés detener la ejecución y cerrar la ventana?"):
                    on_detener()
                    _reset_exec_btn()
                    modal.destroy()
            else:
                on_cerrar()

        def on_root_close_request():
            if modal.winfo_exists() and "Ejecutando" in title_lbl.cget("text"):
                if messagebox.askyesno("Salir", "Hay un test en ejecución. ¿Querés detenerlo y salir de la app?"):
                    on_detener()
                    modal.destroy()
                    root.destroy()
            else:
                if modal.winfo_exists():
                    modal.destroy()
                root.destroy()

        orig_close_protocol = root.protocol("WM_DELETE_WINDOW")
        root.protocol("WM_DELETE_WINDOW", on_root_close_request)



        unmap_id = root.bind("<Unmap>", lambda e: modal.withdraw() if (e.widget == root and modal.winfo_exists()) else None, add="+")
        map_id = root.bind("<Map>", lambda e: (modal.deiconify(), modal.lift()) if (e.widget == root and modal.winfo_exists()) else None, add="+")
        focus_id = root.bind("<FocusIn>", lambda e: modal.lift() if (modal.winfo_exists() and e.widget.winfo_toplevel() == root and e.widget != modal and not str(e.widget).startswith(str(modal))) else None, add="+")

        def cleanup_root_binds(e=None):
            if e and str(e.widget) != str(modal):
                return
            try:
                root.unbind("<Unmap>", unmap_id)
                root.unbind("<Map>", map_id)
                root.unbind("<FocusIn>", focus_id)
                root.protocol("WM_DELETE_WINDOW", orig_close_protocol)
                set_event_blocking(root, False)
            except Exception:
                pass
        modal.bind("<Destroy>", cleanup_root_binds)

        # Custom Title Bar for minimizing and closing
        title_bar = tk.Frame(modal, bg=MODAL_BG)
        title_bar.pack(fill="x", side="top", padx=15, pady=(5, 0))
        
        title_lbl_bar = tk.Label(title_bar, text="Ejecución de Test", font=("Segoe UI", 8, "bold"), bg=MODAL_BG, fg="#C5A9DF")
        title_lbl_bar.pack(side="left")

        # Hacer el modal arrastrable/movible
        def _start_drag(event):
            modal._drag_start_x = event.x
            modal._drag_start_y = event.y

        def _drag(event):
            x = modal.winfo_x() - modal._drag_start_x + event.x
            y = modal.winfo_y() - modal._drag_start_y + event.y
            modal.geometry(f"+{x}+{y}")

        title_bar.bind("<Button-1>", _start_drag)
        title_bar.bind("<B1-Motion>", _drag)
        title_lbl_bar.bind("<Button-1>", _start_drag)
        title_lbl_bar.bind("<B1-Motion>", _drag)
        
        btn_cls = tk.Button(title_bar, text="✕", font=("Segoe UI", 8, "bold"), bg=MODAL_BG, fg="#C5A9DF", relief="flat", bd=0, cursor="hand2", padx=6, pady=2, command=on_close_modal)
        btn_cls.pack(side="right")
        btn_cls.bind("<Enter>", lambda e: btn_cls.config(bg="#E74C3C", fg="white"))
        btn_cls.bind("<Leave>", lambda e: btn_cls.config(bg=MODAL_BG, fg="#C5A9DF"))

        btn_min = tk.Button(title_bar, text="—", font=("Segoe UI", 8, "bold"), bg=MODAL_BG, fg="#C5A9DF", relief="flat", bd=0, cursor="hand2", padx=6, pady=2, command=root.iconify)
        btn_min.pack(side="right", padx=2)
        btn_min.bind("<Enter>", lambda e: btn_min.config(bg="#38234D"))
        btn_min.bind("<Leave>", lambda e: btn_min.config(bg=MODAL_BG))

        # 1. Header (Ejecutando / Completo)
        header_frame = tk.Frame(modal, bg=MODAL_BG)
        header_frame.pack(fill="x", padx=20, pady=(5, 10))
        
        icon_lbl = tk.Label(header_frame, text="↻", font=("Segoe UI", 16, "bold"), bg=MODAL_BG, fg="#C5A9DF")
        icon_lbl.pack(side="left")
        
        # Animación de rotación del icono
        rotation_glyphs = ["↻", "➔", "↻", "➔"]
        def rotate_icon(idx=0):
            if modal.winfo_exists() and "Ejecutando" in title_lbl.cget("text"):
                icon_lbl.config(text=rotation_glyphs[idx % len(rotation_glyphs)])
                modal.after(250, lambda: rotate_icon(idx + 1))

        title_info = tk.Frame(header_frame, bg=MODAL_BG)
        title_info.pack(side="left", padx=10)

        title_lbl = tk.Label(title_info, text="Ejecutando...", font=("Segoe UI", 12, "bold"), bg=MODAL_BG, fg="white")
        title_lbl.pack(anchor="w")
        subtitle_lbl = tk.Label(title_info, text=f"0/{total_sessions} mercado(s) completados", font=("Segoe UI", 9), bg=MODAL_BG, fg=TEXT_SECONDARY)
        subtitle_lbl.pack(anchor="w")
        rotate_icon()
        
        # Botón Detener: mismo comportamiento que la X de cerrar — pide confirmación antes
        # de cortar. on_close_modal ya maneja el popup de confirmación mientras se ejecuta.
        def on_detener():
            stop_event.set()
            btn_detener.config(state="disabled", text=" Deteniendo...")

        def on_detener_click():
            on_close_modal()
            try:
                run_note.config(text="Deteniendo… (termina el lead en curso)", fg="#F8C471")
            except Exception:
                pass
            log_message("[WARN] Detención solicitada por el usuario.")

        btn_detener = tk.Button(header_frame, text=" Detener", image=get_button_icon("stop_coral.png"), compound="left",
                                font=("Segoe UI", 9, "bold"),
                                bg="#3D1220", fg="#F1948A", relief="flat", bd=0, highlightthickness=1,
                                highlightbackground="#F1948A", cursor="hand2", command=on_detener_click, padx=12, pady=4)
        btn_detener.pack(side="right")
        btn_detener.bind("<Enter>", lambda e: btn_detener.config(bg="#5E1D31") if btn_detener["state"] == "normal" else None)
        btn_detener.bind("<Leave>", lambda e: btn_detener.config(bg="#3D1220") if btn_detener["state"] == "normal" else None)

        # 2. Badges Row
        badges_row = tk.Frame(modal, bg=MODAL_BG)
        badges_row.pack(fill="x", padx=20, pady=5)

        def make_pill(parent, text, icon=None):
            img = get_button_icon(icon) if icon else None
            lbl = tk.Label(parent, text=text, font=("Segoe UI", 8, "bold"), bg=MODAL_PILL_BG, fg="white", padx=8, pady=3,
                           bd=0, highlightthickness=1, highlightbackground=BORDER_COLOR)
            if img:
                lbl.config(image=img, compound="left")
            lbl.pack(side="left", padx=3)
            return lbl

        if url_par:
            make_pill(badges_row, f" Por URL: Paralelo (máx {url_max})", icon="link_lav.png")
            make_pill(badges_row, " 1 navegador por URL", icon="gear_lav.png")
        else:
            make_pill(badges_row, f" Mercados: {'Paralelo' if mercados_par else 'Secuencial'}", icon="link_lav.png")
            make_pill(badges_row, f" Excels: {'Paralelo' if excels_par else 'Secuencial'}", icon="gear_lav.png")
        make_pill(badges_row, f" {total_sessions} mercado(s)", icon="monitor_lav.png")

        # Aviso durante la ejecución (se quita al completar)
        run_note = tk.Label(modal, text="⚠ No podés cerrar esta ventana mientras se ejecuta. Para correr otro test ahora, abrí otra ventana de la app.",
                            font=("Segoe UI", 8, "italic"), bg=MODAL_BG, fg="#F8C471", wraplength=480, justify="left")
        run_note.pack(anchor="w", padx=20, pady=(4, 8))

        def _ui(fn):
            try:
                root.after(0, fn)
            except Exception:
                pass

        # 3. Una barra de progreso por MERCADO (país). El nombre aparece una sola vez
        #    arriba de su barra; la barra se llena a medida que avanza.
        from collections import OrderedDict as _OrderedDict
        _pais_totals = _OrderedDict()
        _pais_devices = {}
        for _s in active_sessions_list:
            _pais_totals[_s["pais"]] = _pais_totals.get(_s["pais"], 0) + 1
            
            # Formatear el nombre del dispositivo para mostrar
            dev_name = _s.get("device") or ""
            if "·URL" in dev_name:
                dev_name = dev_name.split("·URL")[0]
            if dev_name == "Mac":
                dev_name = "Mac LT"
            elif dev_name == "Android":
                dev_name = "Android LT"
                
            if _s["pais"] not in _pais_devices:
                _pais_devices[_s["pais"]] = []
            if dev_name and dev_name not in _pais_devices[_s["pais"]]:
                _pais_devices[_s["pais"]].append(dev_name)

        # Área de mercados con scroll: si no entran todas las barras (muchos mercados en
        # una pantalla chica), aparece la barra lateral en vez de recortarse.
        markets_wrap = tk.Frame(modal, bg=MODAL_BG)
        markets_wrap.pack(fill="both", expand=True, padx=20, pady=(2, 8))
        markets_canvas = tk.Canvas(markets_wrap, bg=MODAL_BG, highlightthickness=0)
        markets_sb = tk.Scrollbar(markets_wrap, orient="vertical", command=markets_canvas.yview)
        markets_canvas.configure(yscrollcommand=markets_sb.set)
        markets_canvas.pack(side="left", fill="both", expand=True)
        markets_frame = tk.Frame(markets_canvas, bg=MODAL_BG)
        _mk_win = markets_canvas.create_window((0, 0), window=markets_frame, anchor="nw")

        def _sync_markets_scroll(_e=None):
            try:
                markets_canvas.configure(scrollregion=markets_canvas.bbox("all"))
                markets_canvas.itemconfig(_mk_win, width=markets_canvas.winfo_width())
                hace_falta = markets_frame.winfo_reqheight() > markets_canvas.winfo_height()
                if hace_falta and not markets_sb.winfo_ismapped():
                    markets_sb.pack(side="right", fill="y")
                elif not hace_falta and markets_sb.winfo_ismapped():
                    markets_sb.pack_forget()
            except Exception:
                pass

        markets_frame.bind("<Configure>", _sync_markets_scroll)
        markets_canvas.bind("<Configure>", _sync_markets_scroll)
        markets_canvas.bind("<MouseWheel>",
                            lambda e: markets_canvas.yview_scroll(int(-e.delta / 120), "units"))

        _pais_bars = {}
        for _p, _tot in _pais_totals.items():
            _row = tk.Frame(markets_frame, bg=MODAL_BG)
            _row.pack(fill="x", pady=(0, 9))
            _hdr = tk.Frame(_row, bg=MODAL_BG)
            _hdr.pack(fill="x")
            
            # Formatear el texto de dispositivos
            devs_list = _pais_devices.get(_p, [])
            devs_str = " — " + " / ".join(devs_list) if devs_list else ""
            display_name = f"{_p}{devs_str}"
            
            tk.Label(_hdr, text=display_name, font=("Segoe UI", 10, "bold"), bg=MODAL_BG, fg="white").pack(side="left")
            _stx = tk.Label(_hdr, text=f"0/{_tot} lead(s)", font=("Segoe UI", 8), bg=MODAL_BG, fg=TEXT_SECONDARY)
            _stx.pack(side="right")
            _cb = tk.Canvas(_row, height=8, bg="#35164D", highlightthickness=0)
            _cb.pack(fill="x", pady=(3, 0))
            _fl = _cb.create_rectangle(0, 0, 0, 8, fill="#F8C471", width=0)
            _pais_bars[_p] = {"canvas": _cb, "fill": _fl, "status": _stx,
                              "total": _tot, "done": 0, "ok": 0, "fail": 0}

        # Fracción por sesión (para que la barra se llene también con el avance de leads)
        _sess_frac = {}
        _pais_of_sess = {_s["sess_id"]: _s["pais"] for _s in active_sessions_list}

        def _paint_pais(pais):
            b = _pais_bars.get(pais)
            if not b or not b["canvas"].winfo_exists():
                return
            with _lock:
                ids = [sid for sid, pp in _pais_of_sess.items() if pp == pais]
                frac = sum(_sess_frac.get(sid, 0.0) for sid in ids) / max(1, b["total"])
                done, tot, okc, failc = b["done"], b["total"], b["ok"], b["fail"]
            w = int(max(1, b["canvas"].winfo_width()) * max(0.0, min(1.0, frac)))
            b["canvas"].coords(b["fill"], 0, 0, w, 8)
            if done >= tot:
                col = "#F1948A" if failc else "#82E0AA"
                b["canvas"].itemconfig(b["fill"], fill=col)
                b["status"].config(text=f"✓ {okc} OK · {failc} error(es)", fg=col)
            else:
                b["status"].config(text=f"{done}/{tot} lead(s) listos", fg=TEXT_SECONDARY)

        _dev_label = {"desktop": None, "mac": "Mac LT (Safari)", "android": "Android LT"}

        def show_completed(ok_total, fail_total, detenido, err_msg):
            _exec_state["running"] = False
            if not modal.winfo_exists():
                return
            if detenido:
                icon_lbl.config(text="■", fg="#F1948A", font=("Segoe UI", 16, "bold"))
                title_lbl.config(text="Ejecución detenida")
            elif err_msg:
                icon_lbl.config(text="✕", fg="#F1948A", font=("Segoe UI", 16, "bold"))
                title_lbl.config(text="Ejecución con error")
            else:
                icon_lbl.config(text="✓", fg="#82E0AA", font=("Segoe UI", 18, "bold"))
                title_lbl.config(text="Ejecución completada")
            try:
                btn_detener.pack_forget()
            except Exception:
                pass
            try:
                run_note.destroy()
            except Exception:
                pass

            # Completar todas las barras por mercado
            for _p, b in _pais_bars.items():
                if not b["canvas"].winfo_exists():
                    continue
                _col = "#F1948A" if (b["fail"] or err_msg or detenido) else "#82E0AA"
                _wfull = max(1, b["canvas"].winfo_width())
                b["canvas"].itemconfig(b["fill"], fill=_col)
                b["canvas"].coords(b["fill"], 0, 0, _wfull, 8)
                b["status"].config(text=f"✓ {b['ok']} OK · {b['fail']} error(es)", fg=_col)

            if err_msg:
                tk.Label(modal, text=f"✕ {err_msg}", font=("Segoe UI", 8), bg=MODAL_BG, fg="#F1948A",
                         wraplength=480, justify="left").pack(anchor="w", padx=20, pady=(2, 0))

            if scheduled:
                tk.Label(modal, text="✓ Ya podés cerrar esta ventana. Los tests programados posteriores se ejecutarán igual.",
                         font=("Segoe UI", 8, "italic"), bg=MODAL_BG, fg="#82E0AA", wraplength=480, justify="left").pack(anchor="w", padx=20, pady=(2, 0))

            # Banner de email (el backend encola el envío si "Enviar mail" está activo)
            if enviar_mail and not detenido:
                if dest:
                    _bg, _fg, _tx = "#1F3A30", "#82E0AA", f"✉  Email de resultados encolado a: {dest}"
                else:
                    _bg, _fg, _tx = "#3A1F22", "#F1948A", "⚠  Falta el destinatario: no se envió email."
            else:
                _bg, _fg, _tx = BUTTON_INACTIVE, TEXT_SECONDARY, "✉  Envío de email desactivado."
            _eb = tk.Frame(modal, bg=_bg, bd=0, highlightthickness=1, highlightbackground=_fg)
            _eb.pack(fill="x", padx=20, pady=5)
            tk.Label(_eb, text=_tx, font=("Segoe UI", 9, "bold"), bg=_bg, fg=_fg, pady=4, wraplength=475, justify="center").pack(anchor="center")

            summary_row = tk.Frame(modal, bg=MODAL_BG)
            summary_row.pack(fill="x", padx=20, pady=5)
            tk.Label(summary_row, text=f"🟢 {ok_total} OK      🔴 {fail_total} con error", font=("Segoe UI", 9, "bold"),
                     bg=MODAL_BG, fg="white").pack(side="left")



            btn_close = tk.Button(modal, text="Cerrar resultados", font=("Segoe UI", 10, "bold"), bg="#AED6F1", fg="#110518",
                                  relief="flat", bd=0, cursor="hand2", command=on_cerrar, pady=6)
            btn_close.pack(fill="x", padx=20, pady=(15, 10))
            btn_close.bind("<Enter>", lambda e: btn_close.config(bg="#D4E6F1"))
            btn_close.bind("<Leave>", lambda e: btn_close.config(bg="#AED6F1"))

        # Estado compartido entre sesiones (thread-safe)
        _lock = threading.Lock()
        _st = {"ok": 0, "fail": 0, "done": 0, "err": ""}
        _email_results = []  # entradas para el email consolidado / por país
        _email_lock = threading.Lock()
        # LambdaTest Android suele permitir 1 sesión concurrente (device real): serializamos
        # SOLO las sesiones Android para que TODOS los países se ejecuten (en cola), sin que
        # una quede afuera. El resto (desktop / Mac) sigue en paralelo.
        _android_sem = threading.Semaphore(1)

        def _bump(pais, sess_id=None, ok=0, fail=0, err=""):
            with _lock:
                _st["ok"] += ok
                _st["fail"] += fail
                _st["done"] += 1
                if err:
                    _st["err"] = err
                if sess_id is not None:
                    _sess_frac[sess_id] = 1.0  # sesión terminada = barra de esa sesión llena
                b = _pais_bars.get(pais)
                if b:
                    b["done"] += 1
                    b["ok"] += ok
                    b["fail"] += fail
                done, total = _st["done"], total_sessions
            def _u():
                if modal.winfo_exists():
                    subtitle_lbl.config(text=f"{done}/{total} mercado(s) completados")
                    _paint_pais(pais)
            _ui(_u)

        def _collect_lt_email(pais, navegador, viewport, summary):
            """Registra el resultado LT para email y, si es modo por país, lo envía ya."""
            if not (enviar_mail and not stop_event.is_set()):
                return
            _rp = summary.get("results_excel") if summary else None
            if not _rp:
                return
            _entry = {"pais": pais, "navegador": navegador, "viewport": viewport,
                      "estado": "completado", "excel_path": _rp, "screenshots_dir": None}
            with _email_lock:
                _email_results.append(_entry)
            if _email_modo == "por_pais":
                try:
                    from interface.helpers_interface import enviar_email_resultados_consolidados
                    enviar_email_resultados_consolidados([_entry])
                except Exception as _e:
                    log_message(f"[ERROR] email LT {pais}: {_e}")

        def _run_session(sess):
            """Corre una sesión (un Excel de un mercado en un dispositivo)."""
            if stop_event.is_set():
                return
            pais, dtype, browser, device, excel = sess["pais"], sess["dtype"], sess["browser"], sess["device"], sess["excel"]

            _sid = sess["sess_id"]

            def _set_cur():
                if modal.winfo_exists():
                    title_lbl.config(text="Ejecutando...")
            _ui(_set_cur)

            try:
                if dtype == "desktop":
                    from core.generic_country_base import GenericCountryBase
                    _pausar = var_pausar_autenticacion.get()
                    form = GenericCountryBase(pais, browser=browser, viewport="fullscreen",
                                              headless=False, background=background, is_scheduled=scheduled,
                                              pausar_autenticacion=_pausar)
                    if excel:
                        form.EXCEL_PATH = excel  # ← una sesión por Excel generado
                    def _pcb(done, total):
                        with _lock:
                            _sess_frac[_sid] = (done / total) if total else 0.0
                        _ui(lambda: _paint_pais(pais) if modal.winfo_exists() else None)
                    form.run(progress_callback=_pcb)
                    if enviar_mail and not stop_event.is_set():
                        rp = getattr(form, "RESULTADOS_PATH", None)
                        sd = getattr(form, "SCREENSHOT_DIR", None)
                        if rp:
                            _entry = {"pais": pais, "navegador": browser, "viewport": "fullscreen",
                                      "estado": "completado", "excel_path": rp, "screenshots_dir": sd}
                            with _email_lock:
                                _email_results.append(_entry)
                            if _email_modo == "por_pais":
                                from interface.helpers_interface import enviar_email_resultados
                                enviar_email_resultados(pais, rp, sd, browser=browser, viewport="fullscreen")
                    _bump(pais, _sid, ok=1)
                elif dtype == "mac":
                    sys.path.insert(0, os.path.join(_APP_BASE, "lambdatest_mac"))
                    import lt_controller  # type: ignore
                    b_name = f"Osocio Automatizado LT MAC - {pais}" if scheduled else f"Osocio LT Mac - Envío Manual - {pais}"
                    summary = lt_controller.run(pais=pais, build_name=b_name, excel_path=excel, log_fn=log_message) or {}
                    _collect_lt_email(pais, "lambdatest_mac", "mac", summary)
                    _err = summary.get("error")
                    _ok, _fail = int(summary.get("ok", 0)), int(summary.get("failed", 0))
                    if _err and _ok == 0 and _fail == 0:
                        _bump(pais, _sid, fail=1, err=str(_err)[:200])
                    else:
                        _bump(pais, _sid, ok=_ok, fail=_fail)
                elif dtype == "android":
                    sys.path.insert(0, os.path.join(_APP_BASE, "lambdatest_android"))
                    import lt_android_controller  # type: ignore
                    with _android_sem:  # 1 sesión Android a la vez → todos los países corren
                        if stop_event.is_set():
                            return
                        b_name = f"Osocio Automatizado LT ANDROID - {pais}" if scheduled else f"Osocio LT Android - Envío Manual - {pais}"
                        summary = lt_android_controller.run(pais=pais, build_name=b_name, excel_path=excel, log_fn=log_message) or {}
                    _collect_lt_email(pais, "lambdatest_android", "android", summary)
                    _err = summary.get("error")
                    _ok, _fail = int(summary.get("ok", 0)), int(summary.get("failed", 0))
                    if _err and _ok == 0 and _fail == 0:
                        _bump(pais, _sid, fail=1, err=str(_err)[:200])
                    else:
                        _bump(pais, _sid, ok=_ok, fail=_fail)
            except Exception as e:
                log_message(f"[ERROR] {pais}/{device}: {e}")
                _bump(pais, _sid, fail=1, err=str(e)[:200])

        def _run_market(sessions):
            """Corre las sesiones de un mercado. Si mercados_par o excels_par están
            activos, lanza todos los dispositivos/Excels del mercado en paralelo
            (Chrome local y LambdaTest no compiten, corren en máquinas distintas)."""
            _parallel = (excels_par or mercados_par) and len(sessions) > 1
            if _parallel:
                ts = [threading.Thread(target=_run_session, args=(s,), daemon=True) for s in sessions]
                for t in ts:
                    t.start()
                for t in ts:
                    t.join()
            else:
                for s in sessions:
                    if stop_event.is_set():
                        break
                    _run_session(s)

        def _worker():
            _ensure_serialized_setup()  # evita choque de resultados en paralelo
            if url_par:
                # Una sesión por URL, todas en paralelo con tope de concurrencia (url_max)
                sem = threading.Semaphore(url_max)

                def _guarded(s):
                    if stop_event.is_set():
                        return
                    with sem:
                        _run_session(s)

                ts = [threading.Thread(target=_guarded, args=(s,), daemon=True) for s in flat_sessions]
                for t in ts:
                    t.start()
                for t in ts:
                    t.join()
            elif mercados_par:
                ts = [threading.Thread(target=_run_market, args=(js,), daemon=True) for _, js in market_jobs]
                for t in ts:
                    t.start()
                for t in ts:
                    t.join()
            else:
                for _pais, js in market_jobs:
                    if stop_event.is_set():
                        break
                    _run_market(js)
            _detenido = stop_event.is_set()
            # Modo consolidado: un único email al terminar TODOS los mercados/dispositivos.
            if enviar_mail and _email_modo == "consolidado" and _email_results and not _detenido:
                try:
                    from interface.helpers_interface import enviar_email_resultados_consolidados
                    enviar_email_resultados_consolidados(list(_email_results))
                except Exception as _e:
                    log_message(f"[ERROR] email consolidado: {_e}")
            _ui(lambda: show_completed(_st["ok"], _st["fail"], _detenido, _st["err"]))

        _exec_state["running"] = True
        threading.Thread(target=_worker, daemon=True).start()

    def view_results_dialog():
        carpeta = os.path.join(BASE_DIR, "resultados")
        try:
            os.makedirs(carpeta, exist_ok=True)
            os.startfile(carpeta)
            log_message("[INFO] Abriendo carpeta de resultados.")
        except Exception as e:
            messagebox.showinfo("Resultados", f"Carpeta de resultados:\n{carpeta}\n\n({e})")

    # CTAs en la cabecera de "DATOS POR PAÍS" (junto a la fila de países)
    btn_resultados = tk.Button(d_header, text=" Ver Resultados", image=get_button_icon("report_white.png"), compound="left",
                               font=("Segoe UI", 9, "bold"), bg=BUTTON_INACTIVE, fg="white",
                               relief="flat", bd=0, activebackground=BUTTON_HOVER, activeforeground="white",
                               padx=14, pady=4, cursor="hand2", command=view_results_dialog)
    btn_resultados.pack(side="right", padx=(6, 0))
    btn_resultados.bind("<Enter>", lambda e: btn_resultados.config(bg=BUTTON_HOVER))
    btn_resultados.bind("<Leave>", lambda e: btn_resultados.config(bg=BUTTON_INACTIVE))

    btn_enviar = tk.Button(d_header, text=" EJECUTAR ENVÍO", image=get_button_icon("play_green.png"), compound="left",
                           font=("Segoe UI", 10, "bold"), bg=EXECUTE_BG, fg=EXECUTE_FG,
                           relief="flat", bd=0, activebackground=EXECUTE_HOVER, activeforeground=EXECUTE_FG,
                           padx=22, pady=7, cursor="hand2", command=execute_send_leads)
    btn_enviar.pack(side="right", padx=(6, 0))
    btn_enviar.bind("<Enter>", lambda e: btn_enviar.config(bg=EXECUTE_HOVER) if btn_enviar['state'] == "normal" else None)
    btn_enviar.bind("<Leave>", lambda e: btn_enviar.config(bg=EXECUTE_BG) if btn_enviar['state'] == "normal" else None)

    def refresh_execute_state():
        if any(selected_countries.values()):
            btn_enviar.config(state="normal", bg=EXECUTE_BG, fg=EXECUTE_FG, cursor="hand2")
        else:
            btn_enviar.config(state="disabled", bg=BUTTON_INACTIVE, fg="#9B86B5", cursor="arrow")
    refresh_execute_state()


    # ==========================================
    # TAB 2: VALIDACIÓN DE CAMPOS (validation)
    # ==========================================
    # Funcionalidad completa e idéntica al run original (interface/field_validation_ui.py).
    if BACKEND_OK and build_field_validation_tab is not None:
        build_field_validation_tab(
            make_scrollable_tab_container(tabs["validation"]),
            {
                "app_bg": APP_BG_COLOR,
                "container_bg": CARD_BG_COLOR,
                "section_bg": CARD_BG_COLOR,
                "text_color": TEXT_PRIMARY,
                "button_bg": BUTTON_ACTIVE,
                "button_fg": TEXT_PRIMARY,
                "entry_bg": ENTRY_BG,
                "entry_fg": "#FFFFFF",
                "tree_bg": CARD_BG_COLOR,
                "tree_fg": "#FFFFFF",
                "heading_bg": BUTTON_INACTIVE,
            },
        )
    else:
        tk.Label(tabs["validation"],
                 text="Validación no disponible (backend no cargado).",
                 bg=APP_BG_COLOR, fg=TEXT_PRIMARY, font=("Segoe UI", 10)).pack(padx=20, pady=20)


    # ==========================================
    # TAB: COMPARADOR DEALERS (dealers)
    # ==========================================
    # Deshabilitada temporalmente (todavía no está terminada): se deja el tab_button
    # disabled y esta pestaña muestra un placeholder "Próximamente". El armado real
    # de la pestaña (build_dealer_comparator_tab) queda listo para reactivar cambiando
    # DEALER_COMPARATOR_ENABLED a True cuando esté terminada.
    DEALER_COMPARATOR_ENABLED = True
    if DEALER_COMPARATOR_ENABLED and BACKEND_OK and build_dealer_comparator_tab is not None:
        build_dealer_comparator_tab(
            tabs["dealers"],
            {
                "root": root,
                "APP_BG_COLOR": APP_BG_COLOR,
                "CARD_BG_COLOR": CARD_BG_COLOR,
                "BORDER_COLOR": BORDER_COLOR,
                "ACCENT_COLOR": ACCENT_COLOR,
                "TEXT_PRIMARY": TEXT_PRIMARY,
                "TEXT_SECONDARY": TEXT_SECONDARY,
                "BUTTON_INACTIVE": BUTTON_INACTIVE,
                "BUTTON_ACTIVE": BUTTON_ACTIVE,
                "BUTTON_HOVER": BUTTON_HOVER,
                "VALIDATE_BG": VALIDATE_BG,
                "VALIDATE_FG": VALIDATE_FG,
                "VALIDATE_HOVER": VALIDATE_HOVER,
                "EXECUTE_BG": EXECUTE_BG,
                "EXECUTE_FG": EXECUTE_FG,
                "EXECUTE_HOVER": EXECUTE_HOVER,
                "ENTRY_BG": ENTRY_BG,
                "TEXT_DELETE": TEXT_DELETE,
                "get_button_icon": get_button_icon,
                "make_scrollable_tab_container": make_scrollable_tab_container,
            },
        )
    else:
        tk.Label(tabs["dealers"],
                 text="🚧 Comparador Dealers — Próximamente",
                 bg=APP_BG_COLOR, fg=TEXT_PRIMARY, font=("Segoe UI", 14, "bold")).pack(padx=20, pady=40)


    # ==========================================
    # TAB 3: GENERAR EXCELS CON DATOS (excel)
    # ==========================================
    # Barra de acciones fija (Generar/Regenerar/Borrar), siempre visible sin scrollear
    excel_actions_bar = tk.Frame(tabs["excel"], bg=CARD_BG_COLOR, bd=0, highlightthickness=1, highlightbackground=BORDER_COLOR)
    excel_actions_bar.pack(side="bottom", fill="x", pady=(6, 0))
    excel_footer_btns = tk.Frame(excel_actions_bar, bg=CARD_BG_COLOR)
    excel_footer_btns.pack(fill="x", padx=15, pady=8)

    excel_scroll_frame = make_scrollable_tab_container(tabs["excel"])

    # Variables de control específicas de la pestaña de Generación
    excel_url_mode = tk.StringVar(value="landing_form")
    excel_selected_disp = {d.lower(): False for d in dispositivos}
    excel_selected_disp["chrome"] = True
    excel_disp_btns = {}
    excel_pais_var = tk.StringVar(value="Argentina")
    excel_warn_var = tk.StringVar(value="")

    # Detección de país desde las URLs (igual criterio que el original)
    _EXCEL_COUNTRY_KW = {
        "argentina": "Argentina", "bolivia": "Bolivia", "brasil": "Brasil", "brazil": "Brasil",
        "chile": "Chile", "colombia": "Colombia", "ecuador": "Ecuador",
        "paraguay": "Paraguay", "peru": "Peru", "uruguay": "Uruguay",
        ".com.ar": "Argentina", ".com.bo": "Bolivia", ".com.br": "Brasil", ".com.co": "Colombia",
        ".com.ec": "Ecuador", ".com.py": "Paraguay", ".com.pe": "Peru", ".com.uy": "Uruguay",
    }

    def _detect_excel_country(text):
        low = text.lower()
        for kw, pais in _EXCEL_COUNTRY_KW.items():
            if kw in low:
                return pais
        return None

    # ── 0. CARD: MERCADO A GENERAR (un Excel por mercado a la vez) ──
    mercado_card = tk.Frame(excel_scroll_frame, bg=CARD_BG_COLOR, bd=0, highlightthickness=1, highlightbackground=BORDER_COLOR)
    mercado_card.pack(fill="x", pady=(0, 8), ipady=5)

    m_header = tk.Frame(mercado_card, bg=CARD_BG_COLOR)
    m_header.pack(fill="x", padx=15, pady=(6, 4))
    tk.Label(m_header, text="🌐 MERCADO A GENERAR", font=("Segoe UI", 9, "bold"), bg=CARD_BG_COLOR, fg=TEXT_SECONDARY).pack(side="left")
    tk.Label(m_header, text="Se genera un Excel por mercado a la vez (podés incluir varios dispositivos del mismo mercado).",
             font=("Segoe UI", 8, "italic"), bg=CARD_BG_COLOR, fg="#C5A9DF").pack(side="left", padx=12)

    excel_pais_cards = {}
    excel_pais_labels = {}
    m_grid = tk.Frame(mercado_card, bg=CARD_BG_COLOR)
    m_grid.pack(fill="x", padx=15, pady=2)

    def select_excel_pais(name):
        excel_pais_var.set(name)
        for p, card in excel_pais_cards.items():
            code_lbl, name_lbl = excel_pais_labels[p]
            if p == name:
                card.config(highlightbackground=ACCENT_COLOR, bg=BUTTON_INACTIVE)
                code_lbl.config(fg=ACCENT_COLOR, bg=BUTTON_INACTIVE)
                name_lbl.config(fg=ACCENT_COLOR, bg=BUTTON_INACTIVE)
            else:
                card.config(highlightbackground=BORDER_COLOR, bg=CARD_BG_COLOR)
                code_lbl.config(fg=TEXT_PRIMARY, bg=CARD_BG_COLOR)
                name_lbl.config(fg=TEXT_SECONDARY, bg=CARD_BG_COLOR)
        _on_excel_url_change()

    for idx, pais in enumerate(paises_list):
        code = p_codes[pais]
        sel0 = (pais == "Argentina")
        c_bg = BUTTON_INACTIVE if sel0 else CARD_BG_COLOR
        card = tk.Frame(m_grid, bg=c_bg, bd=0, highlightthickness=1,
                        highlightbackground=ACCENT_COLOR if sel0 else BORDER_COLOR, cursor="hand2")
        card.grid(row=idx // 9, column=idx % 9, padx=3, pady=3, sticky="nsew")
        m_grid.columnconfigure(idx % 9, weight=1)
        code_lbl = tk.Label(card, text=code, font=("Segoe UI", 11, "bold"), bg=c_bg,
                            fg=ACCENT_COLOR if sel0 else TEXT_PRIMARY, cursor="hand2")
        code_lbl.pack(pady=(5, 1))
        name_lbl = tk.Label(card, text=pais, font=("Segoe UI", 8), bg=c_bg,
                            fg=ACCENT_COLOR if sel0 else TEXT_SECONDARY, cursor="hand2")
        name_lbl.pack(pady=(0, 5))
        excel_pais_cards[pais] = card
        excel_pais_labels[pais] = (code_lbl, name_lbl)
        for w in (card, code_lbl, name_lbl):
            w.bind("<Button-1>", lambda e, p=pais: select_excel_pais(p))

    # ── 1. CARD: URLS A PROCESAR ──
    urls_card = tk.Frame(excel_scroll_frame, bg=CARD_BG_COLOR, bd=0, highlightthickness=1, highlightbackground=BORDER_COLOR)
    urls_card.pack(fill="x", pady=(0, 8), ipady=6)

    # Cabecera con Título y Toggle Pills a la derecha
    urls_header = tk.Frame(urls_card, bg=CARD_BG_COLOR)
    urls_header.pack(fill="x", padx=15, pady=(6, 4))

    tk.Label(urls_header, text="🔗 URLS A PROCESAR", font=("Segoe UI", 9, "bold"), bg=CARD_BG_COLOR, fg=TEXT_SECONDARY).pack(side="left")

    mode_btn_frame = tk.Frame(urls_header, bg=CARD_BG_COLOR)
    mode_btn_frame.pack(side="right")

    mode_btns = {}

    def switch_url_mode(mode):
        excel_url_mode.set(mode)
        for m, btn in mode_btns.items():
            if m == mode:
                btn.config(bg=BUTTON_ACTIVE, fg="white", highlightthickness=1, highlightbackground=ACCENT_COLOR)
            else:
                btn.config(bg=BUTTON_INACTIVE, fg=TEXT_SECONDARY, highlightthickness=1, highlightbackground=BUTTON_INACTIVE)
        
        # Actualizar formato de texto descriptivo
        if mode == "landing_form":
            fmt_val_lbl.config(text="FORMATO: url landing  •  url form  •  url landing  •  url form  •  ...")
        else:
            fmt_val_lbl.config(text="FORMATO: url form  •  url form  •  url form  •  url form  •  ...")
        update_excel_calculation()

    for m_val, m_txt in [("landing_form", "URL Landing + URL Form"), ("solo_forms", "Solo URL Form")]:
        b = tk.Button(mode_btn_frame, text=m_txt, font=("Segoe UI", 8, "bold"), bg=BUTTON_INACTIVE, fg=TEXT_SECONDARY,
                      relief="flat", bd=0, activebackground=BUTTON_HOVER, activeforeground="white",
                      highlightthickness=1, highlightbackground=BUTTON_INACTIVE,
                      padx=10, pady=3, cursor="hand2")
        b.pack(side="left", padx=1)
        mode_btns[m_val] = b
        b.config(command=lambda m=m_val: switch_url_mode(m))

        def make_mode_hover(btn=b, val=m_val):
            btn.bind("<Enter>", lambda e: btn.config(bg=BUTTON_HOVER) if excel_url_mode.get() != val else None)
            btn.bind("<Leave>", lambda e: btn.config(bg=BUTTON_INACTIVE) if excel_url_mode.get() != val else None)
        make_mode_hover()

    # Formato e Instrucción
    fmt_row = tk.Frame(urls_card, bg=CARD_BG_COLOR)
    fmt_row.pack(fill="x", padx=15, pady=(4, 4))
    
    fmt_val_lbl = tk.Label(fmt_row, text="FORMATO: url landing  •  url form  •  url landing  •  url form  •  ...",
                           font=("Segoe UI", 8, "bold"), bg=CARD_BG_COLOR, fg="#C5A9DF")
    fmt_val_lbl.pack(side="left")

    # Aviso de discrepancia país detectado vs mercado seleccionado
    excel_warn_lbl = tk.Label(urls_card, textvariable=excel_warn_var, font=("Segoe UI", 8, "italic"),
                              bg=CARD_BG_COLOR, fg="#F8C471", wraplength=900, justify="left")
    excel_warn_lbl.pack(anchor="w", padx=15, pady=(0, 2))

    # Caja de texto para pegar URLs con barra de scroll vertical
    excel_text_border = tk.Frame(urls_card, bg=BORDER_COLOR, padx=1, pady=1)
    excel_text_border.pack(fill="x", padx=15, pady=4)

    v_scroll_text = ttk.Scrollbar(excel_text_border, orient="vertical", style="TScrollbar")
    excel_text_area = tk.Text(excel_text_border, bg=ENTRY_BG, fg="white", insertbackground="white",
                              bd=0, relief="flat", height=5, font=("Consolas", 9),
                              yscrollcommand=v_scroll_text.set)
    v_scroll_text.config(command=excel_text_area.yview)
    v_scroll_text.pack(side="right", fill="y")
    excel_text_area.pack(fill="both", expand=True, padx=(3, 0), pady=3)
    
    # Rellenar con datos de prueba iniciales
    initial_urls = (
        "https://www.ejemplo.com/landing-de-prueba-1\n"
        "https://www.ejemplo.com/formulario-de-prueba-1\n"
        "https://www.ejemplo.com/landing-de-prueba-2\n"
        "https://www.ejemplo.com/formulario-de-prueba-2\n"
    )
    excel_text_area.insert("1.0", initial_urls)

    # ── 2. CARD: DISPOSITIVOS PARA EL EXCEL ──
    excel_devices_card = tk.Frame(excel_scroll_frame, bg=CARD_BG_COLOR, bd=0, highlightthickness=1, highlightbackground=BORDER_COLOR)
    excel_devices_card.pack(fill="x", pady=(0, 8), ipady=6, before=urls_card)

    tk.Label(excel_devices_card, text="🖥 DISPOSITIVOS PARA EL EXCEL", font=("Segoe UI", 9, "bold"), bg=CARD_BG_COLOR, fg=TEXT_SECONDARY).pack(anchor="w", padx=15, pady=(6, 2))
    tk.Label(excel_devices_card, text="Seleccioná en qué dispositivos vas a correr este form. El Excel generado incluirá una columna \"Dispositivo\" con esta info.",
             font=("Segoe UI", 8), bg=CARD_BG_COLOR, fg=TEXT_SECONDARY).pack(anchor="w", padx=15, pady=(0, 6))

    # Formularios T3 2.0 (Adobe AEM): mismos datos, nombre …_T3.xlsx para diferenciar
    var_gen_t3 = tk.BooleanVar(value=False)
    tk.Checkbutton(excel_devices_card, text="🧩 Es formulario T3 2.0 (genera los Excels como …_T3.xlsx)", variable=var_gen_t3,
                   bg=CARD_BG_COLOR, fg=TEXT_SECONDARY, selectcolor=ENTRY_BG, bd=0,
                   activebackground=CARD_BG_COLOR, activeforeground="white",
                   font=("Segoe UI", 8), cursor="hand2",
                   command=lambda: update_excel_calculation()).pack(anchor="w", padx=15, pady=(0, 6))

    # Documentos a generar (solo países con múltiples campos de documento, ej. Brasil).
    # Cada tipo tildado se genera en su columna (CPF/CNPJ/CEP); destildado → columna vacía.
    try:
        from utils.data_generator import DOC_TYPES_BY_COUNTRY as _DOC_TYPES
    except Exception:
        _DOC_TYPES = {"Brasil": ["CPF", "CNPJ", "CEP"]}
    excel_doc_vars = {}
    excel_docs_frame = tk.Frame(excel_devices_card, bg=CARD_BG_COLOR)
    # Label creado UNA sola vez (no dentro del refresh, si no se acumulan copias)
    tk.Label(excel_docs_frame, text="📄 DOCUMENTOS A GENERAR", font=("Segoe UI", 8, "bold"),
             bg=CARD_BG_COLOR, fg=TEXT_SECONDARY).pack(anchor="w", padx=15, pady=(2, 0))
    _docs_cb_row = tk.Frame(excel_docs_frame, bg=CARD_BG_COLOR)
    _docs_cb_row.pack(anchor="w", pady=(0, 4))

    def refresh_doc_types_section(*_):
        tipos = _DOC_TYPES.get(excel_pais_var.get())
        for w in _docs_cb_row.winfo_children():
            w.destroy()
        if not tipos:
            excel_docs_frame.pack_forget()
            return
        for t in tipos:
            excel_doc_vars.setdefault(t, tk.BooleanVar(value=True))
            tk.Checkbutton(_docs_cb_row, text=t, variable=excel_doc_vars[t],
                           bg=CARD_BG_COLOR, fg=TEXT_SECONDARY, selectcolor=ENTRY_BG, bd=0,
                           activebackground=CARD_BG_COLOR, activeforeground="white",
                           font=("Segoe UI", 8), cursor="hand2").pack(side="left", padx=(15 if t == tipos[0] else 8, 0))
        excel_docs_frame.pack(fill="x")

    def _selected_doc_types():
        """dict {tipo: bool} para el país actual, o None si el país no usa multi-doc."""
        tipos = _DOC_TYPES.get(excel_pais_var.get())
        if not tipos:
            return None
        return {t: bool(excel_doc_vars.get(t, tk.BooleanVar(value=True)).get()) for t in tipos}

    # Contenedor horizontal para botones de dispositivos a la izquierda y mensajes al costado
    excel_content_row = tk.Frame(excel_devices_card, bg=CARD_BG_COLOR)
    excel_content_row.pack(fill="x", padx=15, pady=4)

    excel_disp_btn_row = tk.Frame(excel_content_row, bg=CARD_BG_COLOR)
    excel_disp_btn_row.pack(side="left", anchor="nw")

    # Panel de cálculo dinámico para múltiples archivos
    calc_lbl = tk.Label(excel_content_row, text="", font=("Segoe UI", 8, "italic"), bg=CARD_BG_COLOR, fg="#C5A9DF", justify="left", anchor="nw")
    calc_lbl.pack(side="left", fill="both", expand=True, padx=(20, 0), anchor="nw")

    def update_excel_calculation(*_):
        # 1. Contar URLs válidas ingresadas
        raw_text = excel_text_area.get("1.0", "end-1c")
        lines = [ln.strip() for ln in raw_text.splitlines() if ln.strip()]
        num_urls = len(lines)
        
        # Si es modo par (Landing+Form), dividimos por 2
        is_pair_mode = (excel_url_mode.get() == "landing_form")
        effective_urls = num_urls // 2 if is_pair_mode else num_urls

        t3 = bool(var_gen_t3.get())

        # Modo Excel compartido: un único Excel genérico para todos los dispositivos
        if excel_mode_holder[0] == "compartido":
            pais_actual = excel_pais_var.get()
            calc_lbl.config(fg="#F8C471",
                            text=f"Modo Excel compartido: se generará 1 solo Excel con {effective_urls} filas para TODOS los dispositivos.\n"
                                 f"Archivo: {_lead_excel_name(pais_actual, 'Generico', t3)}\n"
                                 f"⚠ Mismos datos para todos → posibles duplicados.")
            return

        # 2. Obtener lista de dispositivos seleccionados
        selected_list = [d.capitalize() for d in dispositivos if excel_selected_disp[d.lower()]]
        num_devices = len(selected_list)
        
        if num_devices == 0:
            calc_lbl.config(text="⚠ Sin selección: se elegirá un dispositivo aleatorio para enviar y generar un solo Excel (ej. Chrome).", fg="#F8C471")
            return
            
        calc_lbl.config(fg="#C5A9DF")
        pais_actual = excel_pais_var.get()
        
        if num_devices == 1:
            device_name = selected_list[0]
            calc_lbl.config(text=f"El Excel tendrá {effective_urls} filas. Columna Dispositivo: {device_name}.\n"
                                 f"Archivo a generar: {_lead_excel_name(pais_actual, _device_excel_suffix(device_name), t3)}")
        else:
            files_lines = []
            for dev in selected_list:
                files_lines.append(f"• {_lead_excel_name(pais_actual, _device_excel_suffix(dev), t3)} ({effective_urls} filas con datos aleatorios independientes)")
            files_str = "\n".join(files_lines)
            calc_lbl.config(text=f"Se generarán {num_devices} Excels independientes (uno por cada dispositivo):\n{files_str}")

    def toggle_excel_disp(name):
        key = name.lower()
        excel_selected_disp[key] = not excel_selected_disp[key]
        btn = excel_disp_btns[key]
        if excel_selected_disp[key]:
            btn.config(bg=BUTTON_ACTIVE, fg="white", highlightthickness=1, highlightbackground=ACCENT_COLOR)
        else:
            btn.config(bg=BUTTON_INACTIVE, fg=TEXT_SECONDARY, highlightthickness=1, highlightbackground=BUTTON_INACTIVE)
        update_excel_calculation()

    for disp in dispositivos:
        d_key = disp.lower()
        init_bg = BUTTON_ACTIVE if d_key == "chrome" else BUTTON_INACTIVE
        init_fg = "white" if d_key == "chrome" else TEXT_SECONDARY
        init_hb = ACCENT_COLOR if d_key == "chrome" else BUTTON_INACTIVE
        
        b = tk.Button(excel_disp_btn_row, text=disp, font=("Segoe UI", 8, "bold"), bg=init_bg, fg=init_fg,
                      relief="flat", bd=0, activebackground=BUTTON_HOVER, activeforeground="white",
                      highlightthickness=1, highlightbackground=init_hb,
                      padx=10, pady=4, cursor="hand2")
        b.pack(side="left", padx=2)
        excel_disp_btns[d_key] = b
        b.config(command=lambda n=disp: toggle_excel_disp(n))

        def make_excel_disp_hover(btn=b, k=d_key):
            btn.bind("<Enter>", lambda e: btn.config(bg=BUTTON_HOVER) if not excel_selected_disp[k] else None)
            btn.bind("<Leave>", lambda e: btn.config(bg=BUTTON_INACTIVE) if not excel_selected_disp[k] else None)
        make_excel_disp_hover()

    def _on_excel_url_change(*_):
        detected = _detect_excel_country(excel_text_area.get("1.0", "end-1c"))
        sel = excel_pais_var.get()
        if detected and detected != sel:
            excel_warn_var.set(f"⚠ Las URLs parecen de {detected} pero tenés seleccionado {sel}. Seleccioná el mercado correcto antes de generar.")
        else:
            excel_warn_var.set("")
        refresh_doc_types_section()
        update_excel_calculation()

    excel_text_area.bind("<KeyRelease>", _on_excel_url_change)
    refresh_doc_types_section()  # estado inicial (Argentina → oculto)



    # Comandos de Generación de Excels (reales, con datos aleatorios por dispositivo)
    def _build_excel_pares():
        raw = excel_text_area.get("1.0", "end-1c")
        lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
        if not lines:
            return None, "Ingresá al menos una URL."
        if excel_url_mode.get() == "solo_forms":
            return [("", u) for u in lines], None
        if len(lines) % 2 != 0:
            return None, f"En modo Landing+Form las URLs deben ir de a pares. Tenés {len(lines)} línea(s)."
        return [(lines[i], lines[i + 1]) for i in range(0, len(lines), 2)], None

    def _email_device_token(dev):
        """Token de dispositivo para el email: chrome/firefox/edge/ltmac/ltandroid."""
        suf = _device_excel_suffix(dev)
        return {"Mac": "ltmac", "Android": "ltandroid"}.get(suf, suf.lower())

    def _rows_for_pais(pais, pares, columnas, device=None, doc_types=None):
        try:
            from utils.data_generator import fixed_values_for_url
        except Exception:
            fixed_values_for_url = lambda _u: {}
        rows = []
        for landing, form in pares:
            datos = generar_fila_datos(pais, device=device, doc_types=doc_types)
            # Valores fijos por form (ej. clubemyev: Modelo Spark EUV + VIN real)
            fijos = fixed_values_for_url(form) or fixed_values_for_url(landing)
            if fijos:
                datos.update(fijos)
            fila = []
            for col in columnas:
                if col == "URL":
                    fila.append(landing)
                elif col == "Formulario":
                    fila.append(form)
                else:
                    fila.append(datos.get(col, ""))
            rows.append(fila)
        return rows

    def _do_generar(title):
        pais = excel_pais_var.get()
        shared = excel_mode_holder[0] == "compartido"
        selected_list = [d.capitalize() for d in dispositivos if excel_selected_disp[d.lower()]]
        if not shared and not selected_list:
            messagebox.showwarning(title, "⚠ Seleccioná al menos un dispositivo.")
            return
        pares, err = _build_excel_pares()
        if err:
            messagebox.showwarning(title, "⚠ " + err)
            return
        detected = _detect_excel_country(excel_text_area.get("1.0", "end-1c"))
        if detected and detected != pais:
            if not messagebox.askyesno("Verificá el mercado",
                                       f"Las URLs parecen de {detected} pero el mercado seleccionado es {pais}.\n\n¿Generar igual para {pais}?"):
                return
        columnas = build_excel_columns_for_country(pais)
        created = []
        try:
            import pandas as pd
            os.makedirs(DATA_DIR, exist_ok=True)

            t3 = bool(var_gen_t3.get())
            doc_types = _selected_doc_types()  # None para países sin multi-documento

            # Modo compartido: un único Excel genérico con los mismos datos para todos.
            if shared:
                rows = _rows_for_pais(pais, pares, columnas, doc_types=doc_types)
                df = pd.DataFrame(rows, columns=columnas).astype(str)
                df.to_excel(_generic_excel_path_for(pais, t3), index=False)
                messagebox.showinfo(title,
                                    f"✓ Generado en data/:\n\n• {_lead_excel_name(pais, 'Generico', t3)} ({len(rows)} filas)\n\n"
                                    "ℹ Modo Excel compartido: TODOS los dispositivos usarán este mismo Excel.\n"
                                    "⚠ Los mismos datos pueden generar leads duplicados o rechazados.")
                log_message(f"[SUCCESS] Generado Excel genérico (compartido) para {pais}.")
                try:
                    if active_p_tab[0] == pais:
                        update_table_data(pais)
                except Exception:
                    pass
                return

            for dev in selected_list:
                rows = _rows_for_pais(pais, pares, columnas, device=_email_device_token(dev), doc_types=doc_types)  # datos aleatorios + email con dispositivo
                fname = _lead_excel_name(pais, _device_excel_suffix(dev), t3)
                pd.DataFrame(rows, columns=columnas).astype(str).to_excel(os.path.join(DATA_DIR, fname), index=False)
                created.append(f"• {fname} ({len(rows)} filas)")
            messagebox.showinfo(title,
                                "✓ Generado(s) en data/:\n\n" + "\n".join(created) +
                                "\n\nℹ En \"Datos por País\" se previsualiza el primer Excel (dispositivo).")
            log_message(f"[SUCCESS] Generados {len(created)} Excel(s) para {pais}.")
            try:
                if active_p_tab[0] == pais:
                    update_table_data(pais)
            except Exception:
                pass
        except PermissionError:
            messagebox.showerror(title, "Cerrá los Excel abiertos y volvé a intentar.")
        except Exception as e:
            messagebox.showerror(title, f"No se pudo generar:\n{e}")

    def cmd_generar_excels():
        _do_generar("Generar Excels")

    def cmd_regen_datos():
        _do_generar("Regenerar Datos")

    def cmd_borrar_urls():
        excel_text_area.delete("1.0", "end")
        excel_warn_var.set("")
        excel_text_area.focus_set()
        update_excel_calculation()
        log_message("[INFO] URLs borradas. Podés ingresar nuevas.")



    btn_exec_excel = tk.Button(excel_footer_btns, text=" GENERAR EXCELS", image=get_button_icon("download_blue.png"), compound="left",
                               font=("Segoe UI", 9, "bold"), bg=VALIDATE_BG, fg=VALIDATE_FG,
                               relief="flat", bd=0, activebackground=VALIDATE_HOVER, activeforeground=VALIDATE_FG,
                               padx=18, pady=6, cursor="hand2", command=cmd_generar_excels)
    btn_exec_excel.pack(side="left", padx=(0, 6))
    btn_exec_excel.bind("<Enter>", lambda e: btn_exec_excel.config(bg=VALIDATE_HOVER))
    btn_exec_excel.bind("<Leave>", lambda e: btn_exec_excel.config(bg=VALIDATE_BG))

    btn_regen_excel = tk.Button(excel_footer_btns, text=" REGENERAR DATOS", image=get_button_icon("bolt_yellow.png"), compound="left",
                                font=("Segoe UI", 9, "bold"), bg=BUTTON_INACTIVE, fg=TEXT_EXCEL,
                                relief="flat", bd=0, activebackground=BUTTON_HOVER, activeforeground=TEXT_EXCEL,
                                padx=18, pady=6, cursor="hand2", command=cmd_regen_datos)
    btn_regen_excel.pack(side="left", padx=6)
    btn_regen_excel.bind("<Enter>", lambda e: btn_regen_excel.config(bg=BUTTON_HOVER))
    btn_regen_excel.bind("<Leave>", lambda e: btn_regen_excel.config(bg=BUTTON_INACTIVE))

    btn_borrar_urls = tk.Button(excel_footer_btns, text=" Borrar URLs", image=get_button_icon("trash_coral.png"), compound="left",
                                font=("Segoe UI", 9, "bold"), bg=BUTTON_INACTIVE, fg=TEXT_DELETE,
                                relief="flat", bd=0, activebackground=BUTTON_HOVER, activeforeground=TEXT_DELETE,
                                padx=18, pady=6, cursor="hand2", command=cmd_borrar_urls)
    btn_borrar_urls.pack(side="left", padx=6)
    btn_borrar_urls.bind("<Enter>", lambda e: btn_borrar_urls.config(bg=BUTTON_HOVER))
    btn_borrar_urls.bind("<Leave>", lambda e: btn_borrar_urls.config(bg=BUTTON_INACTIVE))



    # Inicializar toggle de modo inicial en la pestaña de Excels
    switch_url_mode("landing_form")


    # ==========================================
    # CONFIGURACIÓN DEL MOUSEWHEEL GLOBAL Y SEGURO (Scroll suave y rápido)
    # ==========================================
    def _on_global_mousewheel(event):
        for canvas in scrollable_canvases:
            if not canvas.winfo_viewable():
                continue
            
            # Obtener coordenadas del cursor y del canvas para ver si el mouse está sobre él
            x, y = root.winfo_pointerxy()
            wx = canvas.winfo_rootx()
            wy = canvas.winfo_rooty()
            ww = canvas.winfo_width()
            wh = canvas.winfo_height()
            
            if wx <= x <= wx + ww and wy <= y <= wy + wh:
                scroll_amount = 3  # Multiplicador de velocidad de scroll para fluidez
                if event.delta:
                    # Windows y MacOS
                    direction = -1 if event.delta > 0 else 1
                    canvas.yview_scroll(direction * scroll_amount, "units")
                elif event.num == 4:
                    # Linux scroll up
                    canvas.yview_scroll(-1 * scroll_amount, "units")
                elif event.num == 5:
                    # Linux scroll down
                    canvas.yview_scroll(1 * scroll_amount, "units")
                break

    # Bindear a nivel root global
    root.bind_all("<MouseWheel>", _on_global_mousewheel)
    root.bind_all("<Button-4>", _on_global_mousewheel)
    root.bind_all("<Button-5>", _on_global_mousewheel)

    # Créditos (footer discreto)
    footer_credit = tk.Frame(root, bg=APP_BG_COLOR)
    footer_credit.pack(side="bottom", fill="x", pady=(0, 2))
    tk.Label(footer_credit, text="Some Updates by Elian Zás", font=("Segoe UI", 7), bg=APP_BG_COLOR, fg="#8A6DB0").pack(side="right", padx=(0, 20))
    tk.Label(footer_credit, text="Made by Ariel Melgratti", font=("Segoe UI", 8), bg=APP_BG_COLOR, fg="#D8B4FE").pack(expand=True)

    # Inicializar y mostrar pestaña por defecto
    switch_tab("leads")

    def _restore_from_tray():
        root.deiconify()
        root.state("normal")
        root.focus_force()
        global _tray_instance
        if _tray_instance:
            try:
                _tray_instance.destroy()
            except Exception:
                pass
            _tray_instance = None

    def _force_close():
        try:
            from core.browser_manager import kill_active_drivers
            kill_active_drivers()
        except Exception:
            pass
        try:
            global _tray_instance
            if _tray_instance:
                _tray_instance.destroy()
                _tray_instance = None
        except Exception:
            pass
        try:
            root.destroy()
        except Exception:
            pass
        os._exit(0)

    def _menu_tray(x, y):
        """
        Menú del click derecho sobre el icono de la bandeja.

        Se dibuja con Tk (no con TrackPopupMenu): la ventana propietaria del icono es una
        ventana ctypes de 0x0 invisible, SetForegroundWindow sobre ella falla y Windows 11
        terminaba mostrando el menú como un recuadro blanco vacío.

        OJO: esto lo llama wnd_proc, que es un callback de ctypes. Tocar Tk ahí adentro
        crashea el intérprete ("PyEval_RestoreThread ... GIL"), así que se difiere con
        after() para que el menú se arme dentro del loop de Tk.
        """
        def _popup():
            # La ventana principal está withdrawn, así que no puede tomar el foco y el menú
            # quedaba abierto para siempre al clickear afuera. Se usa un Toplevel de 1x1
            # transparente como dueño del menú: ese sí toma el foco, y cuando lo pierde
            # (click en cualquier otro lado) cerramos el menú.
            owner = tk.Toplevel(root)
            owner.overrideredirect(True)
            owner.geometry(f"1x1+{x}+{y}")
            owner.attributes("-alpha", 0.0)
            owner.attributes("-topmost", True)
            owner.focus_force()

            menu = tk.Menu(owner, tearoff=0)

            def _cerrar(*_):
                try:
                    menu.unpost()
                except Exception:
                    pass
                try:
                    owner.destroy()
                except Exception:
                    pass

            def _elegir(accion):
                _cerrar()
                accion()

            menu.add_command(label="Restaurar", command=lambda: _elegir(_restore_from_tray))
            menu.add_separator()
            menu.add_command(label="Salir", command=lambda: _elegir(_force_close))
            owner.bind("<FocusOut>", _cerrar)
            menu.bind("<Unmap>", _cerrar)

            try:
                menu.tk_popup(x, y)
            finally:
                menu.grab_release()

        try:
            root.after(0, _popup)
        except Exception:
            pass

    def _bombear_eventos_tray():
        # El icono corre en su propio hilo (ver SysTrayIcon): los clicks llegan por una cola
        # y se ejecutan acá, en el hilo de Tk, porque Tk no es thread-safe.
        global _tray_instance
        if _tray_instance is None:
            return
        try:
            _tray_instance.procesar_eventos()
        except Exception:
            pass
        try:
            root.after(80, _bombear_eventos_tray)
        except Exception:
            pass

    def _on_close():
        global _tray_instance
        # Si hay una ejecución en curso (modal abierto), no cerrar: minimizar para
        # que la corrida continúe. El usuario puede restaurar la ventana luego.
        if _exec_state.get("running"):
            try:
                root.iconify()
            except Exception:
                pass
            return
        if os.name == 'nt' and var_minimizar_a_bandeja.get():
            root.withdraw()
            if _tray_instance is None:
                icon_path = os.path.join(ASSET_DIR, "icon.ico")
                try:
                    _tray_instance = SysTrayIcon(
                        icon_path=icon_path,
                        hover_text="Osocio - Form Automation",
                        on_quit=_force_close,
                        on_double_click=_restore_from_tray,
                        on_right_click=_menu_tray,
                    )
                    _bombear_eventos_tray()
                except Exception:
                    pass
        else:
            _force_close()

    root.protocol("WM_DELETE_WINDOW", _on_close)
    root.mainloop()

def create_demo_interface():
    iniciar_interfaz()

if __name__ == "__main__":
    iniciar_interfaz()
