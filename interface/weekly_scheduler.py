"""
weekly_scheduler.py — Panel y diálogo de programación semanal recurrente.
WeeklySchedulerPanel: tarjeta de estado embebida en la pestaña de testing.
WeeklySchedulerDialog: modal de configuración (días, horarios, países).
"""
import json
import os
import threading
import time
from datetime import datetime, date as _date
from tkinter import *
from tkinter import messagebox

from utils.scheduling import guardar_programacion, cargar_programacion, limpiar_programacion
from .helpers_interface import cargar_config_global, guardar_config_global

# ── Colores del tema Figma (purple dark) ──────────────────────────────────────
SCH_BG      = "#5B1A87"
SCH_CARD    = "#7230A0"
SCH_PRIMARY = "#C084FC"
SCH_PFG     = "#2D0060"
SCH_MUTED   = "#D8B4FE"
SCH_BORDER  = "#8B44B8"
SCH_HOVER   = "#6B2A97"
SCH_AMBER   = "#F59E0B"
SCH_GREEN   = "#10B981"
SCH_RED     = "#F87171"
SCH_WHITE   = "#FFFFFF"

DAYS_OF_WEEK = [
    ("Lun", "Lunes"),
    ("Mar", "Martes"),
    ("Mié", "Miércoles"),
    ("Jue", "Jueves"),
    ("Vie", "Viernes"),
    ("Sáb", "Sábado"),
    ("Dom", "Domingo"),
]

HOURS = [f"{h:02d}:{m:02d}" for h in range(24) for m in (0, 15, 30, 45)]

COUNTRIES = [
    "Argentina", "Bolivia", "Brasil",
    "Chile",     "Colombia", "Ecuador",
    "Paraguay",  "Peru",     "Uruguay",
]

DAYS_ES = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]

COUNTRY_FLAGS = {
    "Argentina": "🇦🇷", "Bolivia": "🇧🇴", "Brasil": "🇧🇷",
    "Chile": "🇨🇱",     "Colombia": "🇨🇴", "Ecuador": "🇪🇨",
    "Paraguay": "🇵🇾",  "Peru": "🇵🇪",     "Uruguay": "🇺🇾",
}


# ──────────────────────────────────────────────────────────────────────────────
# DIÁLOGO DE CONFIGURACIÓN
# ──────────────────────────────────────────────────────────────────────────────

class WeeklySchedulerDialog(Toplevel):
    """Modal de configuración semanal: días, horarios y países."""

    def __init__(self, parent, on_save, initial_config=None):
        super().__init__(parent)
        self.title("Configurar automatización")
        self.configure(bg=SCH_BG)
        self.geometry("820x700")
        self.minsize(700, 580)
        self.resizable(True, True)
        self.grab_set()

        self._on_save = on_save
        self._selected_day = None
        self._copy_open_state = False
        self._copy_selected_days = []
        self._hour_btns = {}
        self._badges_frame = None
        self._copy_frame_inner = None
        self._val_lbl = None
        self._save_btn = None
        self._edit_all_days = False

        cfg = initial_config or {}
        self._schedule = {k: list(v) for k, v in cfg.get("horarios", {}).items()}
        self._countries = list(cfg.get("paises", []))

        self._build_ui()
        self._center_on(parent)

    def _center_on(self, parent):
        self.update_idletasks()
        px = parent.winfo_rootx() + (parent.winfo_width()  - self.winfo_width())  // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{max(0, px)}+{max(0, py)}")

    # ── Build ──────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Header (top)
        header = Frame(self, bg=SCH_CARD)
        header.pack(fill="x")
        Label(header, text="⚙  Configurar automatización",
              font=("Segoe UI", 13, "bold"), bg=SCH_CARD, fg=SCH_WHITE).pack(side=LEFT, padx=16, pady=12)
        btn_x = Button(header, text="✕", font=("Segoe UI", 12), bg=SCH_CARD, fg=SCH_MUTED,
                       relief=FLAT, cursor="hand2", activebackground=SCH_BG,
                       activeforeground=SCH_WHITE, bd=0, command=self.destroy)
        btn_x.pack(side=RIGHT, padx=12, pady=8)
        Frame(self, bg=SCH_BORDER, height=1).pack(fill="x")

        # Footer fixed at bottom (packed before canvas so it stays pinned)
        self._build_footer(self, dict(padx=14, pady=10))
        Frame(self, bg=SCH_BORDER, height=1).pack(fill="x", side=BOTTOM)

        # Scrollable body (fills remaining space between header and footer)
        canvas = Canvas(self, bg=SCH_BG, highlightthickness=0)
        self._scroll_canvas = canvas  # kept for child widgets to bind scroll
        vsb = Scrollbar(self, orient=VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side=RIGHT, fill=Y)
        canvas.pack(side=LEFT, fill=BOTH, expand=True)

        body = Frame(canvas, bg=SCH_BG)
        win_id = canvas.create_window((0, 0), window=body, anchor="nw")

        body.bind("<Configure>",
                  lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfig(win_id, width=e.width))

        def _scroll(e):
            canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _scroll)
        self.bind("<Destroy>",
                  lambda e, c=canvas: c.unbind_all("<MouseWheel>") if e.widget is self else None)

        p = dict(padx=14, pady=8)
        self._build_days_section(body, p)
        self._build_countries_section(body, p)

    def _card_frame(self, parent):
        f = Frame(parent, bg=SCH_CARD, highlightbackground=SCH_BORDER,
                  highlightthickness=1)
        return f

    # ── Days ───────────────────────────────────────────────────────────────────

    def _build_days_section(self, parent, pad):
        card = self._card_frame(parent)
        card.pack(fill="x", **pad)

        hrow = Frame(card, bg=SCH_CARD)
        hrow.pack(fill="x", padx=12, pady=(12, 6))
        Label(hrow, text="DÍAS DE LA SEMANA", font=("Segoe UI", 9, "bold"),
              bg=SCH_CARD, fg=SCH_MUTED).pack(side=LEFT)

        self._total_badge = Label(hrow, text="", font=("Segoe UI", 8, "bold"),
                                   bg=SCH_BG, fg=SCH_PRIMARY, padx=8, pady=2)
        self._total_badge.pack(side=LEFT, padx=6)

        self._clear_all_lbl = Label(hrow, text="Desmarcar todos",
                                     font=("Segoe UI", 9, "underline"),
                                     bg=SCH_CARD, fg=SCH_PRIMARY, cursor="hand2")
        self._clear_all_lbl.pack(side=RIGHT)
        self._clear_all_lbl.bind("<Button-1>", lambda e: self._clear_all())

        # Day buttons row
        day_row = Frame(card, bg=SCH_CARD)
        day_row.pack(fill="x", padx=12, pady=(4, 8))
        self._day_btns = {}
        for short, full in DAYS_OF_WEEK:
            col_f = Frame(day_row, bg=SCH_CARD)
            col_f.pack(side=LEFT, expand=True, fill="x", padx=2)
            btn = Button(col_f, text=f"{short}\n—",
                         font=("Segoe UI", 8, "bold"), relief=FLAT,
                         cursor="hand2", pady=8, wraplength=60)
            btn.pack(fill="x")
            btn.config(command=lambda d=full: self._select_day(d))
            self._day_btns[full] = btn

        # Hour panel placeholder
        self._hours_outer = Frame(card, bg=SCH_BG)
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

        # Header
        h_hdr = Frame(self._hours_outer, bg=SCH_BG)
        h_hdr.pack(fill="x", pady=(8, 4))
        Label(h_hdr, text=f"🕐  Horarios para el {self._selected_day}",
              font=("Segoe UI", 10, "bold"), bg=SCH_BG, fg=SCH_WHITE).pack(side=LEFT)
        Button(h_hdr, text="✕", font=("Segoe UI", 10), bg=SCH_BG, fg=SCH_MUTED,
               relief=FLAT, cursor="hand2", bd=0,
               command=self._close_hours_panel).pack(side=RIGHT)
        Button(h_hdr, text="✓ Listo", font=("Segoe UI", 8, "bold"),
               bg=SCH_GREEN, fg=SCH_WHITE, relief=FLAT, cursor="hand2",
               padx=8, pady=3,
               command=self._close_hours_panel).pack(side=RIGHT, padx=(0, 6))

        # Edit mode toggle: Solo este día / Todos los días
        mode_row = Frame(self._hours_outer, bg=SCH_BG)
        mode_row.pack(fill="x", pady=(0, 4))
        Label(mode_row, text="Modo:", font=("Segoe UI", 8),
              bg=SCH_BG, fg=SCH_MUTED).pack(side=LEFT)
        Button(mode_row, text="Solo este día",
               font=("Segoe UI", 8, "bold"),
               bg=SCH_PRIMARY if not self._edit_all_days else SCH_HOVER,
               fg=SCH_PFG if not self._edit_all_days else SCH_MUTED,
               relief=FLAT, cursor="hand2", padx=6, pady=2,
               command=lambda: self._set_edit_mode(False)).pack(side=LEFT, padx=(4, 2))
        Button(mode_row, text="Todos los días",
               font=("Segoe UI", 8, "bold"),
               bg=SCH_AMBER if self._edit_all_days else SCH_HOVER,
               fg=SCH_WHITE if self._edit_all_days else SCH_MUTED,
               relief=FLAT, cursor="hand2", padx=6, pady=2,
               command=lambda: self._set_edit_mode(True)).pack(side=LEFT, padx=2)
        if self._edit_all_days:
            Label(mode_row, text="⚠ Cambios aplican a TODOS los días",
                  font=("Segoe UI", 8), bg=SCH_BG, fg=SCH_AMBER).pack(side=LEFT, padx=8)

        # Hour grid (6 cols × 16 rows: comfortable size, readable font)
        grid = Frame(self._hours_outer, bg=SCH_BG)
        grid.pack(fill="x", pady=4)
        current = self._schedule.get(self._selected_day, [])
        self._hour_btns = {}
        for c in range(6):
            grid.columnconfigure(c, weight=1)

        def _scroll_fwd(e):
            if hasattr(self, "_scroll_canvas"):
                self._scroll_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")

        for idx, hour in enumerate(HOURS):
            r, c = divmod(idx, 6)
            picked = hour in current
            btn = Button(grid, text=hour, font=("Segoe UI", 9),
                         bg=SCH_PRIMARY if picked else SCH_HOVER,
                         fg=SCH_PFG if picked else SCH_MUTED,
                         relief=RAISED if picked else FLAT,
                         cursor="hand2", padx=4, pady=6, width=5)
            btn.grid(row=r, column=c, padx=2, pady=2, sticky="ew")
            btn.config(command=lambda h=hour: self._toggle_hour(h))
            btn.bind("<MouseWheel>", _scroll_fwd)
            self._hour_btns[hour] = btn

        # Badges row
        self._badges_frame = Frame(self._hours_outer, bg=SCH_BG)
        self._badges_frame.pack(fill="x", pady=(4, 0))
        self._refresh_badges()

        # Copy section
        Frame(self._hours_outer, bg=SCH_BORDER, height=1).pack(fill="x", pady=(8, 0))
        self._copy_frame_inner = Frame(self._hours_outer, bg=SCH_BG)
        self._copy_frame_inner.pack(fill="x", pady=4)
        self._build_copy_ui()

    def _refresh_badges(self):
        if not self._badges_frame:
            return
        for w in self._badges_frame.winfo_children():
            w.destroy()
        selected = sorted(self._schedule.get(self._selected_day, []))
        MAX_SHOW = 12
        for h in selected[:MAX_SHOW]:
            Label(self._badges_frame, text=h, font=("Segoe UI", 9, "bold"),
                  bg=SCH_BG, fg=SCH_PRIMARY, padx=6, pady=3,
                  relief=SOLID, bd=1).pack(side=LEFT, padx=2)
        if len(selected) > MAX_SHOW:
            Label(self._badges_frame, text=f"+{len(selected) - MAX_SHOW} más",
                  font=("Segoe UI", 9), bg=SCH_BG, fg=SCH_MUTED).pack(side=LEFT, padx=4)

    def _toggle_hour(self, hour):
        day = self._selected_day
        currently_picked = hour in self._schedule.get(day, [])

        # Always modify the selected day first
        dh = self._schedule.setdefault(day, [])
        if currently_picked:
            dh.remove(hour)
        else:
            dh.append(hour)
            dh.sort()

        if self._edit_all_days:
            # Replace ALL other days with the exact schedule of the selected day
            new_hours = list(self._schedule.get(day, []))
            for _, full in DAYS_OF_WEEK:
                if full == day:
                    continue
                if new_hours:
                    self._schedule[full] = list(new_hours)
                else:
                    self._schedule.pop(full, None)

        # Clean empty selected day
        if not self._schedule.get(day):
            self._schedule.pop(day, None)

        has_hours_now = bool(self._schedule.get(day))
        # True when visibility of copy section must change (0→1 or N→0 hours)
        copy_visibility_changed = (
            (currently_picked and not has_hours_now) or          # removed last hour
            (not currently_picked and len(self._schedule.get(day, [])) == 1)  # added first hour
        )

        picked = hour in self._schedule.get(day, [])
        btn = self._hour_btns.get(hour)
        if btn:
            btn.config(bg=SCH_PRIMARY if picked else SCH_HOVER,
                       fg=SCH_PFG if picked else SCH_MUTED,
                       relief=RAISED if picked else FLAT)
        self._refresh_badges()
        self._update_day_buttons()
        self._update_footer()
        if copy_visibility_changed:
            self._build_copy_ui()

    def _update_day_buttons(self):
        total = sum(len(v) for v in self._schedule.values())
        if total > 0:
            self._total_badge.config(text=f"{total} horario{'s' if total != 1 else ''}")
            self._total_badge.pack(side=LEFT, padx=6)
            self._clear_all_lbl.pack(side=RIGHT)
        else:
            self._total_badge.pack_forget()
            self._clear_all_lbl.pack_forget()

        for short, full in DAYS_OF_WEEK:
            count = len(self._schedule.get(full, []))
            is_open = self._selected_day == full
            btn = self._day_btns[full]
            if is_open:
                btn.config(bg=SCH_PRIMARY, fg=SCH_PFG,
                           text=f"{short}\n● abierto")
            elif count > 0:
                btn.config(bg=SCH_HOVER, fg=SCH_PRIMARY,
                           text=f"{short}\n{count} sel.")
            else:
                btn.config(bg=SCH_HOVER, fg=SCH_MUTED,
                           text=f"{short}\n—")

    def _clear_all(self):
        self._schedule = {}
        self._close_hours_panel()
        self._update_day_buttons()
        self._update_footer()

    # ── Copy to other days ─────────────────────────────────────────────────────

    def _build_copy_ui(self):
        if not self._copy_frame_inner:
            return
        for w in self._copy_frame_inner.winfo_children():
            w.destroy()

        current = self._schedule.get(self._selected_day, [])
        if not current:
            return

        if not self._copy_open_state:
            quick_row = Frame(self._copy_frame_inner, bg=SCH_BG)
            quick_row.pack(fill="x", pady=2)
            Button(quick_row,
                   text="📅  Aplicar a otros días",
                   font=("Segoe UI", 9, "bold"), bg=SCH_BG, fg=SCH_PRIMARY,
                   relief=FLAT, cursor="hand2", pady=6, bd=1,
                   highlightbackground=SCH_PRIMARY, highlightthickness=1,
                   command=self._open_copy_ui).pack(side=LEFT, fill="x", expand=True, padx=(0, 4))
            Button(quick_row,
                   text="⚡ TODOS",
                   font=("Segoe UI", 9, "bold"), bg=SCH_AMBER, fg=SCH_WHITE,
                   relief=FLAT, cursor="hand2", pady=6,
                   command=self._apply_copy_all).pack(side=LEFT)
        else:
            Label(self._copy_frame_inner,
                  text=f"Copiar horarios de {self._selected_day} a:",
                  font=("Segoe UI", 8, "bold"), bg=SCH_BG, fg=SCH_MUTED).pack(anchor="w")

            days_row = Frame(self._copy_frame_inner, bg=SCH_BG)
            days_row.pack(fill="x", pady=4)
            other_days = [d for _, d in DAYS_OF_WEEK if d != self._selected_day]
            for d in other_days:
                is_sel = d in self._copy_selected_days
                has = bool(self._schedule.get(d))
                label = f"{d} (ya tiene)" if has and not is_sel else d
                b = Button(days_row, text=label, font=("Segoe UI", 8),
                           bg=SCH_PRIMARY if is_sel else SCH_HOVER,
                           fg=SCH_PFG if is_sel else SCH_MUTED,
                           relief=FLAT, cursor="hand2", padx=6, pady=4)
                b.pack(side=LEFT, padx=2)
                b.config(command=lambda dd=d: self._toggle_copy_day(dd))

            act = Frame(self._copy_frame_inner, bg=SCH_BG)
            act.pack(fill="x", pady=4)
            n = len(self._copy_selected_days)
            lbl = f"Aplicar a {n} día{'s' if n!=1 else ''}" if n > 0 else "Aplicar a ..."
            Button(act, text=f"✓ {lbl}", font=("Segoe UI", 8, "bold"),
                   bg=SCH_PRIMARY, fg=SCH_PFG, relief=FLAT, cursor="hand2",
                   padx=8, pady=4, state=NORMAL if n > 0 else DISABLED,
                   command=self._apply_copy).pack(side=LEFT, padx=(0, 10))
            lnk = Label(act, text="Cancelar", font=("Segoe UI", 8, "underline"),
                        bg=SCH_BG, fg=SCH_MUTED, cursor="hand2")
            lnk.pack(side=LEFT)
            lnk.bind("<Button-1>", lambda e: self._close_copy_ui())

    def _set_edit_mode(self, all_days):
        self._edit_all_days = all_days
        self._build_hours_panel()

    def _apply_copy_all(self):
        src = self._schedule.get(self._selected_day, [])
        if not src:
            return
        for _, full in DAYS_OF_WEEK:
            if full != self._selected_day:
                merged = sorted(set(self._schedule.get(full, [])) | set(src))
                self._schedule[full] = merged
        self._update_day_buttons()
        self._update_footer()
        messagebox.showinfo("Copiado",
                            f"Horarios de {self._selected_day} copiados a todos los días.",
                            parent=self)

    def _open_copy_ui(self):
        self._copy_open_state = True
        self._copy_selected_days = []
        self._build_copy_ui()

    def _close_copy_ui(self):
        self._copy_open_state = False
        self._copy_selected_days = []
        self._build_copy_ui()

    def _toggle_copy_day(self, day):
        if day in self._copy_selected_days:
            self._copy_selected_days.remove(day)
        else:
            self._copy_selected_days.append(day)
        self._build_copy_ui()

    def _apply_copy(self):
        if not self._copy_selected_days:
            return
        src = self._schedule.get(self._selected_day, [])
        for d in self._copy_selected_days:
            merged = sorted(set(self._schedule.get(d, [])) | set(src))
            self._schedule[d] = merged
        n = len(self._copy_selected_days)
        names = ", ".join(self._copy_selected_days)
        self._close_copy_ui()
        self._update_day_buttons()
        messagebox.showinfo("Copiado",
                            f"Horarios copiados a {n} día{'s' if n!=1 else ''}:\n{names}",
                            parent=self)

    # ── Countries ──────────────────────────────────────────────────────────────

    def _build_countries_section(self, parent, pad):
        card = self._card_frame(parent)
        card.pack(fill="x", **pad)

        hrow = Frame(card, bg=SCH_CARD)
        hrow.pack(fill="x", padx=12, pady=(12, 6))
        Label(hrow, text="🌎  PAÍSES A TESTEAR",
              font=("Segoe UI", 9, "bold"), bg=SCH_CARD, fg=SCH_MUTED).pack(side=LEFT)
        self._toggle_all_lbl = Label(hrow, text="Seleccionar todos",
                                      font=("Segoe UI", 9, "underline"),
                                      bg=SCH_CARD, fg=SCH_PRIMARY, cursor="hand2")
        self._toggle_all_lbl.pack(side=RIGHT)
        self._toggle_all_lbl.bind("<Button-1>", lambda e: self._toggle_all())

        grid = Frame(card, bg=SCH_CARD)
        grid.pack(fill="x", padx=12, pady=(0, 4))
        for c in range(3):
            grid.columnconfigure(c, weight=1)

        self._country_vars = {}
        self._country_items = {}  # frame per country for re-styling

        for idx, country in enumerate(COUNTRIES):
            r, c = divmod(idx, 3)
            checked = country in self._countries
            var = BooleanVar(value=checked)
            self._country_vars[country] = var

            bg = SCH_PRIMARY if checked else SCH_HOVER
            border = SCH_PRIMARY if checked else SCH_BORDER
            item = Frame(grid, bg=bg, highlightbackground=border, highlightthickness=1)
            item.grid(row=r, column=c, padx=3, pady=3, sticky="ew")
            self._country_items[country] = item

            cb = Checkbutton(item, text=country, variable=var,
                             font=("Segoe UI", 9), bg=bg,
                             fg=SCH_PFG if checked else SCH_WHITE,
                             activebackground=SCH_PRIMARY,
                             selectcolor=SCH_PRIMARY,
                             padx=8, pady=6)
            cb.pack(fill="x")

            var.trace("w", lambda *a, co=country, it=item, cb_=cb, v=var:
                      self._on_country_toggle(co, it, cb_, v))

        self._count_lbl = Label(card, text="", font=("Segoe UI", 8),
                                 bg=SCH_CARD, fg=SCH_MUTED)
        self._count_lbl.pack(anchor="w", padx=12, pady=(2, 10))
        self._update_country_count()

    def _on_country_toggle(self, country, item, cb, var):
        checked = var.get()
        bg = SCH_PRIMARY if checked else SCH_HOVER
        border = SCH_PRIMARY if checked else SCH_BORDER
        item.config(bg=bg, highlightbackground=border)
        cb.config(bg=bg, fg=SCH_PFG if checked else SCH_WHITE)
        if checked and country not in self._countries:
            self._countries.append(country)
        elif not checked and country in self._countries:
            self._countries.remove(country)
        self._update_country_count()
        self._update_toggle_all_label()
        self._update_footer()

    def _toggle_all(self):
        all_sel = len(self._countries) == len(COUNTRIES)
        for var in self._country_vars.values():
            var.set(not all_sel)

    def _update_toggle_all_label(self):
        if len(self._countries) == len(COUNTRIES):
            self._toggle_all_lbl.config(text="Desmarcar todos")
        else:
            self._toggle_all_lbl.config(text="Seleccionar todos")

    def _update_country_count(self):
        n = len(self._countries)
        self._count_lbl.config(
            text=f"{n} de {len(COUNTRIES)} países seleccionados" if n > 0 else "")

    # ── Footer ─────────────────────────────────────────────────────────────────

    def _build_footer(self, parent, pad):
        footer = Frame(parent, bg=SCH_BG)
        footer.pack(fill="x", side=BOTTOM, **pad)

        self._val_lbl = Label(footer, text="", font=("Segoe UI", 9),
                               bg=SCH_BG, fg=SCH_AMBER,
                               padx=10, pady=6, justify="left", anchor="w",
                               wraplength=500, relief=FLAT)

        btn_row = Frame(footer, bg=SCH_BG)
        btn_row.pack(fill="x")
        self._save_btn = Button(btn_row, text="💾  Guardar configuración",
                                font=("Segoe UI", 10, "bold"),
                                bg=SCH_PRIMARY, fg=SCH_PFG,
                                relief=FLAT, cursor="hand2", padx=16, pady=8,
                                command=self._save)
        self._save_btn.pack(side=RIGHT)
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
        self._save_btn.config(state=NORMAL if can else DISABLED,
                               bg=SCH_PRIMARY if can else SCH_HOVER,
                               cursor="hand2" if can else "arrow")

    def _save(self):
        total = sum(len(v) for v in self._schedule.values())
        if total == 0 or not self._countries:
            return
        config = {
            "horarios": {k: v for k, v in self._schedule.items() if v},
            "paises": list(self._countries),
        }
        self._on_save(config)
        self.destroy()


# ──────────────────────────────────────────────────────────────────────────────
# PANEL DE ESTADO (embebido en la UI principal)
# ──────────────────────────────────────────────────────────────────────────────

class WeeklySchedulerPanel(Frame):
    """
    Tarjeta de estado para la programación semanal recurrente.
    Se embebe en testing_tab reemplazando el viejo panel de fecha/hora.
    """

    _STATUS_META = {
        "idle":      ("🕐", "Sin configurar",  SCH_HOVER),
        "scheduled": ("📅", "Programado",      "#5B2A9A"),
        "running":   ("⚙",  "Ejecutando...",   "#7A5A00"),
        "completed": ("✓",  "Completado",      "#0A5A3A"),
        "stopped":   ("■",  "Detenido",        "#7A1A1A"),
    }

    def __init__(self, parent, get_navegadores_cb, get_viewports_cb,
                 on_scheduling_change=None, execute_cb=None,
                 send_email_cb=None, root=None, **kwargs):
        super().__init__(parent, bg=SCH_CARD, **kwargs)

        self._get_navegadores = get_navegadores_cb
        self._get_viewports   = get_viewports_cb
        self._on_change       = on_scheduling_change
        self._execute_cb      = execute_cb
        self._send_email_cb   = send_email_cb
        self._root            = root

        self._config    = None   # saved weekly config (horarios, paises)
        self._status    = "idle"
        self._is_active = False
        self._last_triggered = self._load_triggered()  # {(day_str, hour_str): date}
        self._stop_event = threading.Event()

        # Load saved config draft from config_global
        cfg_g = cargar_config_global()
        draft = cfg_g.get("scheduler_config")
        if draft and draft.get("horarios") and draft.get("paises"):
            self._config = draft

        # Check if there's an active weekly schedule in JSON
        existing = cargar_programacion()
        if existing and existing.get("tipo") == "semanal":
            self._config    = existing
            self._is_active = True
            self._status    = "scheduled"

        self._build_ui()
        self._refresh_ui()

        threading.Thread(target=self._monitor_loop, daemon=True).start()

    # ── Build ──────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Title row
        title_row = Frame(self, bg=SCH_CARD)
        title_row.pack(fill="x", padx=16, pady=(14, 2))
        Label(title_row, text="Test Automático", font=("Segoe UI", 12, "bold"),
              bg=SCH_CARD, fg=SCH_WHITE).pack(side=LEFT)
        self._badge = Label(title_row, text="", font=("Segoe UI", 8, "bold"),
                             bg=SCH_HOVER, fg=SCH_WHITE, padx=10, pady=3)
        self._badge.pack(side=RIGHT)

        Label(self, text="Automatización de leads por país y horario",
              font=("Segoe UI", 9), bg=SCH_CARD, fg=SCH_MUTED).pack(anchor="w", padx=16)

        self._content = Frame(self, bg=SCH_CARD)
        self._content.pack(fill="x", padx=16, pady=(10, 0))

        self._progress = Frame(self, bg=SCH_CARD)

        Frame(self, bg=SCH_BORDER, height=1).pack(fill="x", pady=(10, 0))

        self._actions = Frame(self, bg=SCH_CARD)
        self._actions.pack(fill="x", padx=16, pady=10)

    # ── Refresh ────────────────────────────────────────────────────────────────

    def _refresh_ui(self):
        for w in self._content.winfo_children():
            w.destroy()
        for w in self._actions.winfo_children():
            w.destroy()
        for w in self._progress.winfo_children():
            w.destroy()
        self._progress.pack_forget()

        icon, label, bg = self._STATUS_META.get(self._status, self._STATUS_META["idle"])
        self._badge.config(text=f"{icon}  {label}", bg=bg)

        if self._config:
            self._render_summary()
        else:
            self._render_empty()

        if self._status == "running":
            self._progress.pack(fill="x", padx=16, pady=(0, 6))
            Label(self._progress, text="⚙  Procesando países...",
                  font=("Segoe UI", 9), bg=SCH_CARD, fg=SCH_AMBER).pack(anchor="w")
        elif self._status == "completed":
            self._progress.pack(fill="x", padx=16, pady=(0, 6))
            Label(self._progress, text="✓  Todos los países procesados correctamente.",
                  font=("Segoe UI", 9), bg=SCH_CARD, fg=SCH_GREEN).pack(anchor="w")
        elif self._status == "stopped":
            self._progress.pack(fill="x", padx=16, pady=(0, 6))
            Label(self._progress, text="■  El test fue detenido manualmente.",
                  font=("Segoe UI", 9), bg=SCH_CARD, fg=SCH_RED).pack(anchor="w")

        self._render_actions()

    def _render_empty(self):
        c = Frame(self._content, bg=SCH_CARD)
        c.pack(expand=True, pady=14)
        Label(c, text="⚙", font=("Segoe UI", 22), bg=SCH_CARD, fg=SCH_MUTED).pack()
        Label(c, text="Sin configuración",
              font=("Segoe UI", 11, "bold"), bg=SCH_CARD, fg=SCH_WHITE).pack(pady=(4, 0))
        Label(c, text="Configurá los días, horarios y países\npara poder programar el test.",
              font=("Segoe UI", 9), bg=SCH_CARD, fg=SCH_MUTED, justify="center").pack()

    def _render_summary(self):
        cfg = self._config

        Label(self._content, text="🌎  PAÍSES A TESTEAR",
              font=("Segoe UI", 8, "bold"), bg=SCH_CARD, fg=SCH_MUTED).pack(anchor="w", pady=(0, 4))
        c_row = Frame(self._content, bg=SCH_CARD)
        c_row.pack(fill="x", pady=(0, 10))
        for country in cfg.get("paises", []):
            flag = COUNTRY_FLAGS.get(country, "")
            Label(c_row, text=f"{flag} {country}",
                  font=("Segoe UI", 9, "bold"), bg=SCH_HOVER, fg=SCH_WHITE,
                  padx=8, pady=3).pack(side=LEFT, padx=2)

        horarios = {k: v for k, v in cfg.get("horarios", {}).items() if v}
        total = sum(len(v) for v in horarios.values())
        Label(self._content,
              text=f"🕐  HORARIO CONFIGURADO  ({total} slot{'s' if total != 1 else ''})",
              font=("Segoe UI", 8, "bold"), bg=SCH_CARD, fg=SCH_MUTED).pack(anchor="w", pady=(0, 4))
        for day, hours in horarios.items():
            row = Frame(self._content, bg=SCH_CARD)
            row.pack(fill="x", pady=1)
            Label(row, text=f"{day}:", font=("Segoe UI", 9, "bold"),
                  bg=SCH_CARD, fg=SCH_WHITE, width=12, anchor="w").pack(side=LEFT)
            for h in hours:
                Label(row, text=h, font=("Segoe UI", 8),
                      bg=SCH_BG, fg=SCH_PRIMARY, padx=5, pady=1).pack(side=LEFT, padx=2)

    def _render_actions(self):
        locked = self._status in ("running", "scheduled")

        if self._status != "running":
            cfg_text = "⚙  Editar configuración" if self._config else "⚙  Configurar automatización"
            Button(self._actions, text=cfg_text,
                   font=("Segoe UI", 9, "bold"), bg=SCH_HOVER, fg=SCH_WHITE,
                   relief=FLAT, cursor="hand2" if not locked else "arrow",
                   padx=12, pady=6,
                   state=DISABLED if locked else NORMAL,
                   command=self._open_dialog).pack(side=LEFT, padx=(0, 8))

        if self._status == "idle":
            Button(self._actions, text="▶  Programar test automático",
                   font=("Segoe UI", 9, "bold"), bg=SCH_PRIMARY, fg=SCH_PFG,
                   relief=FLAT, cursor="hand2" if self._config else "arrow",
                   padx=12, pady=6,
                   state=NORMAL if self._config else DISABLED,
                   command=self._activate).pack(side=LEFT)

        elif self._status in ("completed", "stopped"):
            Button(self._actions, text="■  Desactivar",
                   font=("Segoe UI", 9, "bold"), bg=SCH_HOVER, fg=SCH_RED,
                   relief=FLAT, cursor="hand2", padx=12, pady=6,
                   command=self._deactivate).pack(side=LEFT)

        elif self._status == "scheduled":
            Button(self._actions, text="▶  Iniciar ahora",
                   font=("Segoe UI", 9, "bold"), bg=SCH_AMBER, fg=SCH_WHITE,
                   relief=FLAT, cursor="hand2", padx=12, pady=6,
                   command=self._run_now).pack(side=LEFT, padx=(0, 8))
            Button(self._actions, text="■  Desactivar",
                   font=("Segoe UI", 9, "bold"), bg=SCH_HOVER, fg=SCH_RED,
                   relief=FLAT, cursor="hand2", padx=12, pady=6,
                   command=self._deactivate).pack(side=LEFT)

        elif self._status == "running":
            Button(self._actions, text="■  Detener",
                   font=("Segoe UI", 9, "bold"), bg=SCH_RED, fg=SCH_WHITE,
                   relief=FLAT, cursor="hand2", padx=12, pady=6,
                   command=self._stop).pack(side=LEFT)

    # ── State management ───────────────────────────────────────────────────────

    def is_active(self):
        return self._is_active

    def set_status(self, status):
        self._status = status
        if status in ("running",):
            self._is_active = True
        elif status in ("completed", "stopped"):
            pass  # keep is_active as-is (scheduled remains active for next week)
        root = self._root or self.winfo_toplevel()
        root.after(0, self._refresh_ui)

    def _activate(self):
        if not self._config:
            return
        navs = self._get_navegadores()
        vps  = self._get_viewports()
        if not navs:
            messagebox.showwarning("Configuración", "Seleccioná al menos un navegador en Configuración Global.")
            return
        full_config = dict(self._config)
        full_config["tipo"]       = "semanal"
        full_config["navegadores"] = navs
        full_config["viewports"]   = vps
        guardar_programacion(full_config)
        self._config    = full_config
        self._is_active = True
        self._status    = "scheduled"
        self._refresh_ui()
        if self._on_change:
            self._on_change(True)

    def _deactivate(self):
        limpiar_programacion()
        self._is_active = False
        self._status    = "idle"
        self._refresh_ui()
        if self._on_change:
            self._on_change(False)

    def _stop(self):
        self._stop_event.set()  # signals _execute_scheduled to abort
        self._is_active = False
        self._status    = "stopped"
        self._refresh_ui()
        if self._on_change:
            self._on_change(False)

    def _open_dialog(self):
        root = self._root or self.winfo_toplevel()
        WeeklySchedulerDialog(root, on_save=self._on_dialog_save,
                               initial_config=self._config)

    def _on_dialog_save(self, new_config):
        self._config = new_config
        # Persist draft in config_global so it survives restarts
        cfg_g = cargar_config_global()
        cfg_g["scheduler_config"] = new_config
        guardar_config_global(cfg_g)
        # If currently active, also update programacion_test.json
        if self._is_active and self._status == "scheduled":
            navs = self._get_navegadores()
            vps  = self._get_viewports()
            full = dict(new_config)
            full["tipo"] = "semanal"
            full["navegadores"] = navs
            full["viewports"]   = vps
            guardar_programacion(full)
            self._config = full
        self._refresh_ui()

    def _run_now(self):
        self._stop_event.clear()
        threading.Thread(target=self._execute_scheduled, daemon=True).start()

    # ── Execution ──────────────────────────────────────────────────────────────

    def _execute_scheduled(self):
        if not self._config or not self._execute_cb:
            return
        self._stop_event.clear()
        root = self._root or self.winfo_toplevel()
        root.after(0, lambda: self.set_status("running"))
        try:
            resultados = self._execute_cb(self._config, self._stop_event)
            if self._stop_event.is_set():
                print("⛔ Ejecución detenida por el usuario.")
                return
            if resultados and self._send_email_cb:
                try:
                    self._send_email_cb(resultados)
                except Exception as e:
                    print(f"❌ Error enviando email consolidado: {e}")
        except Exception as e:
            print(f"❌ Error durante ejecución programada: {e}")
        if not self._stop_event.is_set():
            root.after(0, lambda: self.set_status("completed"))

    # ── Triggered persistence ──────────────────────────────────────────────────

    _TRIGGERED_FILE = os.path.join(os.path.dirname(__file__), '..', 'json', 'scheduler_triggered.json')

    def _load_triggered(self):
        try:
            with open(self._TRIGGERED_FILE, 'r', encoding='utf-8') as f:
                raw = json.load(f)
            return {tuple(k.split('|')): _date.fromisoformat(v) for k, v in raw.items()}
        except Exception:
            return {}

    def _save_triggered(self):
        try:
            raw = {f"{k[0]}|{k[1]}": v.isoformat() for k, v in self._last_triggered.items()}
            with open(self._TRIGGERED_FILE, 'w', encoding='utf-8') as f:
                json.dump(raw, f)
        except Exception:
            pass

    # ── Monitor loop ───────────────────────────────────────────────────────────

    def _monitor_loop(self):
        startup = True  # primera iteración: marcar slots actuales sin ejecutar
        while True:
            try:
                if self._is_active and self._status == "scheduled" and self._config:
                    ahora      = datetime.now()
                    dia        = DAYS_ES[ahora.weekday()]
                    slot_min   = (ahora.minute // 15) * 15
                    hora       = f"{ahora.hour:02d}:{slot_min:02d}"
                    programados = self._config.get("horarios", {}).get(dia, [])
                    if hora in programados:
                        key = (dia, hora)
                        if self._last_triggered.get(key) != ahora.date():
                            self._last_triggered[key] = ahora.date()
                            self._save_triggered()
                            if not startup:
                                self._stop_event.clear()
                                self._execute_scheduled()
            except Exception as e:
                print(f"⚠️  Error en monitor semanal: {e}")
            startup = False
            time.sleep(60)
