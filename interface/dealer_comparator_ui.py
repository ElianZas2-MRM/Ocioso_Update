# -*- coding: utf-8 -*-
"""
dealer_comparator_ui.py — Pestaña "Comparador Dealers".

Permite cargar un Excel de dealers esperados (fila de encabezado y columnas configurables,
porque el Excel que mandan varía de país a país y no siempre arranca en la fila 1), abrir
un formulario real (landing+form o solo form, igual criterio que "Envío de Leads"), navegar
región→ciudad→dealer y comparar contra el Excel filtrado, generando un reporte
PASS/FAIL/EXTRA/DUPLICADO/NOTA (Excel, y opcionalmente capturas de pantalla) — un reporte y
una carpeta propia por cada formulario (landing+form o solo form) del Excel de URLs.

Visualmente sigue el mismo lenguaje que "Envío de Leads" y "Generar Excels con Datos"
(cards con código de país, bloque de URLs con pills + textarea igual al de Generar Excels,
pills de dispositivo, barra de acciones fija abajo con solo 2 botones: Borrar URLs y
Ejecutar, y un modal de ejecución igual al de Envío de Leads).
"""
import json
import logging
import os
import sys
import threading
from datetime import datetime
from tkinter import (
    BooleanVar, Button, Canvas, Checkbutton, Entry, Frame, Label, StringVar, Text,
    Toplevel, filedialog, messagebox, ttk,
)

from core.browser_manager import BrowserManager
from core.dealer_comparator_runner import (
    DEFAULT_SELECT_IDS,
    StopRequested,
    _locate_form_iframe,
    advance_to_selects,
    banner_lines_for_result,
    capture_dropdown_evidence,
    capture_result_screenshot,
    compare_dealers,
    detect_hidden_columns,
    detect_hidden_rows,
    evidence_level_for_result,
    export_results_excel,
    filter_rows,
    find_extra_dealers,
    get_country_level_defaults,
    list_model_options,
    open_target,
    read_excel_rows,
    resolve_column,
)

LOGGER = logging.getLogger(__name__)

PAISES_LIST = ["Argentina", "Bolivia", "Brasil", "Chile", "Colombia", "Ecuador", "Paraguay", "Peru", "Uruguay"]
P_CODES = {"Argentina": "AR", "Bolivia": "BO", "Brasil": "BR", "Chile": "CL", "Colombia": "CO",
           "Ecuador": "EC", "Paraguay": "PY", "Peru": "PE", "Uruguay": "UY"}
BROWSER_OPTIONS = [("chrome", "Chrome"), ("firefox", "Firefox"), ("edge", "Edge")]

# Mismo criterio de detección de país por URL que usa "Generar Excels con Datos"
_URL_COUNTRY_KW = {
    "argentina": "Argentina", "bolivia": "Bolivia", "brasil": "Brasil", "brazil": "Brasil",
    "chile": "Chile", "colombia": "Colombia", "ecuador": "Ecuador",
    "paraguay": "Paraguay", "peru": "Peru", "uruguay": "Uruguay",
    ".com.ar": "Argentina", ".com.bo": "Bolivia", ".com.br": "Brasil", ".com.co": "Colombia",
    ".com.ec": "Ecuador", ".com.py": "Paraguay", ".com.pe": "Peru", ".com.uy": "Uruguay",
}


def _detect_url_country(text):
    low = (text or "").lower()
    for kw, pais in _URL_COUNTRY_KW.items():
        if kw in low:
            return pais
    return None


MODAL_BG = "#231830"
MODAL_PILL_BG = "#38234D"


def _get_base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _get_settings_path():
    json_dir = os.path.join(_get_base_dir(), "json")
    os.makedirs(json_dir, exist_ok=True)
    return os.path.join(json_dir, "dealer_comparator_settings.json")


def _get_results_dir():
    return os.path.join(_get_base_dir(), "Dealerscheck_resultados")


def _load_all_settings():
    path = _get_settings_path()
    if not os.path.exists(path):
        return {"paises": {}, "presets": {}}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
            data.setdefault("paises", {})
            data.setdefault("presets", {})
            return data
    except Exception:
        return {"paises": {}, "presets": {}}


def _save_all_settings(all_settings):
    path = _get_settings_path()
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(all_settings, f, indent=2, ensure_ascii=False)
    except Exception as e:  # noqa: BLE001
        LOGGER.warning("No se pudo guardar dealer_comparator_settings.json: %s", e)


def _default_country_settings(pais):
    levels = get_country_level_defaults(pais)
    return {
        "excel_path": "",
        "header_row": 1,
        "filter_col": "",
        "filter_value": "",
        "condition_mode": "include",
        "has_region": levels.get("has_region", True),
        "has_city": levels.get("has_city", True),
        "col_region": "REGION",
        "col_city": "CIUDAD",
        "col_dealer": "NOMBRE",
        "col_bac": "BAC",
        "chk_bac": False,
        "find_extras": True,
        "field_checks": [],
        "has_models": False,
        "models_field_id": "models",
        "models_mode": "all",
        "models_list": "",
        "url_mode": "landing_form",
        "url_excel_path": "",
        "browser": "chrome",
        "viewport": "fullscreen",
        "output_mode": "excel",
    }


def build_dealer_comparator_tab(tab_frame, ctx):
    """tab_frame: el Frame crudo de la pestaña (tabs['dealers']), sin scroll todavía.
    ctx: dict con colores/helpers ya definidos en main_interface.py, para que esta
    pestaña se vea igual que 'Envío de Leads' / 'Generar Excels con Datos':
      root, CARD_BG_COLOR, BORDER_COLOR, APP_BG_COLOR, ACCENT_COLOR, TEXT_PRIMARY, TEXT_SECONDARY,
      BUTTON_INACTIVE, BUTTON_ACTIVE, BUTTON_HOVER, VALIDATE_BG, VALIDATE_FG, VALIDATE_HOVER,
      ENTRY_BG, TEXT_DELETE, get_button_icon, make_scrollable_tab_container.
    """
    root = ctx["root"]
    CARD_BG = ctx["CARD_BG_COLOR"]
    BORDER = ctx["BORDER_COLOR"]
    ACCENT = ctx["ACCENT_COLOR"]
    TEXT_P = ctx["TEXT_PRIMARY"]
    TEXT_S = ctx["TEXT_SECONDARY"]
    BTN_INACTIVE = ctx["BUTTON_INACTIVE"]
    BTN_ACTIVE = ctx["BUTTON_ACTIVE"]
    BTN_HOVER = ctx["BUTTON_HOVER"]
    VALIDATE_BG = ctx["VALIDATE_BG"]
    VALIDATE_FG = ctx["VALIDATE_FG"]
    VALIDATE_HOVER = ctx["VALIDATE_HOVER"]
    EXECUTE_BG = ctx.get("EXECUTE_BG", VALIDATE_BG)
    EXECUTE_FG = ctx.get("EXECUTE_FG", VALIDATE_FG)
    EXECUTE_HOVER = ctx.get("EXECUTE_HOVER", VALIDATE_HOVER)
    ENTRY_BG = ctx["ENTRY_BG"]
    TEXT_DELETE = ctx["TEXT_DELETE"]
    get_button_icon = ctx["get_button_icon"]
    make_scrollable_tab_container = ctx["make_scrollable_tab_container"]

    # Estilo propio para los combobox (dropdown "Configuraciones guardadas" y "Condición")
    style = ttk.Style()
    style.configure("Dealer.TCombobox", fieldbackground=ENTRY_BG, background=BTN_ACTIVE,
                     foreground="white", arrowcolor="white", bordercolor=BORDER,
                     lightcolor=BORDER, darkcolor=BORDER, padding=7)
    style.map("Dealer.TCombobox",
              fieldbackground=[("readonly", ENTRY_BG), ("disabled", BTN_INACTIVE)],
              foreground=[("readonly", "white")],
              background=[("active", BTN_HOVER), ("readonly", BTN_ACTIVE)])
    root.option_add("*TCombobox*Listbox.background", ENTRY_BG)
    root.option_add("*TCombobox*Listbox.foreground", "white")
    root.option_add("*TCombobox*Listbox.selectBackground", BTN_ACTIVE)
    root.option_add("*TCombobox*Listbox.selectForeground", "white")
    root.option_add("*TCombobox*Listbox.font", ("Segoe UI", 9))

    all_settings = _load_all_settings()
    state = {
        "pais": None,
        "headers": [], "rows": [], "filtered_rows": [], "results": [],
        "driver": None, "running": False,
        "stop_event": threading.Event(),
        "field_check_widgets": [],
        "modal": None,
        "expect_absent": False,
    }

    # ── Barra de acciones fija (solo Borrar URLs + Ejecutar), siempre visible ──
    actions_bar = Frame(tab_frame, bg=CARD_BG, bd=0, highlightthickness=1, highlightbackground=BORDER)
    actions_bar.pack(side="bottom", fill="x", pady=(6, 0))
    footer_btns = Frame(actions_bar, bg=CARD_BG)
    footer_btns.pack(fill="x", padx=15, pady=8)

    root_frame = make_scrollable_tab_container(tab_frame)

    # ── Guía rápida (breve — el detalle está en la mini-guía de cada bloque) ─────
    guide_card = Frame(root_frame, bg=CARD_BG, bd=0, highlightthickness=1, highlightbackground=BORDER)
    guide_card.pack(fill="x", pady=(0, 8), ipady=4)
    Label(guide_card, text="Seguí los pasos ①→④ de cada bloque. EJECUTAR se habilita cuando está todo completo.",
          font=("Segoe UI", 9, "bold"), bg=CARD_BG, fg=TEXT_P).pack(anchor="w", padx=15, pady=(6, 6))

    def card(title, subtitle=None, parent=None, side=None, expand=False, subtitle_wrap=750):
        """Crea una card. Por defecto se apila a lo ancho en root_frame; si se pasa
        `parent`+`side` se coloca como columna dentro de un contenedor horizontal."""
        holder = parent if parent is not None else root_frame
        block = Frame(holder, bg=CARD_BG, bd=0, highlightthickness=1, highlightbackground=BORDER)
        if side:
            block.pack(side=side, fill="both", expand=expand, padx=(0, 8), pady=(0, 8), ipady=5)
        else:
            block.pack(fill="x", pady=(0, 8), ipady=5)
        head = Frame(block, bg=CARD_BG)
        head.pack(fill="x", padx=15, pady=(6, 4))
        Label(head, text=title, font=("Segoe UI", 9, "bold"), bg=CARD_BG, fg=TEXT_S).pack(side="left")
        if subtitle:
            Label(head, text=subtitle, font=("Segoe UI", 8, "italic"), bg=CARD_BG, fg="#C5A9DF",
                  wraplength=subtitle_wrap, justify="left").pack(side="left", padx=12)
        return block

    def labeled_entry(parent_frame, label_text, var, width=16):
        col = Frame(parent_frame, bg=CARD_BG)
        col.pack(side="left", padx=6, pady=3)
        Label(col, text=label_text, font=("Segoe UI", 8), bg=CARD_BG, fg=TEXT_S).pack(anchor="w")
        ent = Entry(col, textvariable=var, width=width, bg=ENTRY_BG, fg="white", font=("Segoe UI", 10),
                    insertbackground="white", relief="flat",
                    disabledbackground="#2E2140", disabledforeground="#8A7E9E")
        ent.pack(ipady=4, fill="x")
        return ent

    def make_single_select(parent, options, var, default=None, on_select=None):
        """options: lista de (valor, etiqueta). Pills de selección única (estilo Excel/Leads)."""
        btns = {}

        def _select(val):
            var.set(val)
            for v, b in btns.items():
                if v == val:
                    b.config(bg=BTN_ACTIVE, fg="white", highlightthickness=1, highlightbackground=ACCENT)
                else:
                    b.config(bg=BTN_INACTIVE, fg=TEXT_S, highlightthickness=1, highlightbackground=BTN_INACTIVE)
            if on_select:
                on_select(val)

        for val, label_text in options:
            b = Button(parent, text=label_text, font=("Segoe UI", 9, "bold"), bg=BTN_INACTIVE, fg=TEXT_S,
                       relief="flat", bd=0, activebackground=BTN_HOVER, activeforeground="white",
                       highlightthickness=1, highlightbackground=BTN_INACTIVE,
                       padx=16, pady=6, cursor="hand2", command=lambda v=val: _select(v))
            b.pack(side="left", padx=3)
            btns[val] = b
        if default is not None:
            _select(default)
        return _select

    def make_toggle_pill(parent, label_text, var, locked=False):
        """Pill tipo checkbox (se prende/apaga independientemente). Si locked=True, queda
        siempre prendida y no se puede tocar (ej. "Dealer", que siempre es obligatorio)."""

        def _refresh():
            if var.get():
                btn.config(bg=BTN_ACTIVE, fg="white", highlightthickness=1, highlightbackground=ACCENT)
            else:
                btn.config(bg=BTN_INACTIVE, fg=TEXT_S, highlightthickness=1, highlightbackground=BTN_INACTIVE)

        def _toggle(event=None):
            if locked:
                return "break"
            var.set(not var.get())
            _refresh()
            return "break"

        btn = Button(parent, text=label_text, font=("Segoe UI", 9, "bold"), bg=BTN_INACTIVE, fg=TEXT_S,
                     relief="flat", bd=0, activebackground=BTN_INACTIVE, activeforeground=TEXT_S,
                     highlightthickness=1, highlightbackground=BTN_INACTIVE,
                     padx=16, pady=6, cursor=("arrow" if locked else "hand2"))
        # Usar bind en lugar de command para evitar que activebackground "trague" el primer click
        btn.bind("<Button-1>", _toggle)
        btn.pack(side="left", padx=3)
        _refresh()
        return btn, _refresh

    def ui_log(msg, level="info"):
        LOGGER.info("[%s] %s", level, msg)

    def mini_guide(parent, text):
        """Mini instructivo dentro de un bloque (una línea, estilo del resto de la app)."""
        Label(parent, text=text, font=("Segoe UI", 8, "italic"), bg=CARD_BG, fg="#C5A9DF",
              wraplength=340, justify="left").pack(anchor="w", padx=15, pady=(0, 4))

    # ── 1. Mercado ────────────────────────────────────────────────────────────
    mercado_card = card("🌐 MERCADO A CHEQUEAR")
    mini_guide(mercado_card, "① Elegí el país cuyo formulario vas a chequear.")
    m_grid = Frame(mercado_card, bg=CARD_BG)
    m_grid.pack(fill="x", padx=15, pady=2)

    pais_cards = {}
    pais_labels = {}

    def _refresh_pais_cards():
        for p, c in pais_cards.items():
            code_lbl, name_lbl = pais_labels[p]
            sel = (p == state["pais"])
            c.config(highlightbackground=ACCENT if sel else BORDER, bg=BTN_INACTIVE if sel else CARD_BG)
            code_lbl.config(fg=ACCENT if sel else TEXT_P, bg=BTN_INACTIVE if sel else CARD_BG)
            name_lbl.config(fg=ACCENT if sel else TEXT_S, bg=BTN_INACTIVE if sel else CARD_BG)

    for idx, pais in enumerate(PAISES_LIST):
        code = P_CODES[pais]
        c = Frame(m_grid, bg=CARD_BG, bd=0, highlightthickness=1, highlightbackground=BORDER, cursor="hand2")
        c.grid(row=idx // 9, column=idx % 9, padx=3, pady=3, sticky="nsew")
        m_grid.columnconfigure(idx % 9, weight=1)
        code_lbl = Label(c, text=code, font=("Segoe UI", 11, "bold"), bg=CARD_BG, fg=TEXT_P, cursor="hand2")
        code_lbl.pack(pady=(5, 1))
        name_lbl = Label(c, text=pais, font=("Segoe UI", 8), bg=CARD_BG, fg=TEXT_S, cursor="hand2")
        name_lbl.pack(pady=(0, 5))
        pais_cards[pais] = c
        pais_labels[pais] = (code_lbl, name_lbl)
        for w in (c, code_lbl, name_lbl):
            w.bind("<Button-1>", lambda e, p=pais: select_country(p))

    # ── 2. Bloque de 3 columnas: Navegador · Excel de URLs · Modelos ─────────────
    blockA = Frame(root_frame, bg=root_frame.cget("bg"))
    blockA.pack(fill="x", pady=(0, 0))

    # Col 1 — Dispositivos / Navegadores
    dev_col = card("🖥 NAVEGADOR", parent=blockA, side="left", expand=True, subtitle_wrap=200)
    mini_guide(dev_col, "② Elegí el navegador con el que se abre el form.")
    browser_var = StringVar(value="chrome")
    dev_row = Frame(dev_col, bg=CARD_BG)
    dev_row.pack(anchor="w", padx=15, pady=(0, 2))
    make_single_select(dev_row, BROWSER_OPTIONS, browser_var, default="chrome")
    lt_row = Frame(dev_col, bg=CARD_BG)
    lt_row.pack(anchor="w", padx=15, pady=(2, 2))
    for _lt in ("Mac LT", "Android LT"):
        Button(lt_row, text=_lt, font=("Segoe UI", 9, "bold"), bg=BTN_INACTIVE, fg="#6E5A86",
               relief="flat", bd=0, state="disabled", disabledforeground="#6E5A86",
               highlightthickness=1, highlightbackground=BTN_INACTIVE,
               padx=14, pady=6).pack(side="left", padx=3)
    Label(lt_row, text="(próx.)", font=("Segoe UI", 8, "italic"),
          bg=CARD_BG, fg="#6E5A86").pack(side="left", padx=4)
    ver_navegador_var = BooleanVar(value=False)
    viewport_var = StringVar(value="fullscreen")  # siempre escritorio

    # Col 2 — Excel de URLs del form
    urls_col = card("🔗 EXCEL DE URLs", parent=blockA, side="left", expand=True, subtitle_wrap=200)
    mini_guide(urls_col, "③ Cargá un Excel con columnas 'URL' y 'Formulario' (como el de Envío de Leads).")
    url_excel_path_var = StringVar()

    def _pick_url_excel():
        path = filedialog.askopenfilename(title="Seleccionar Excel de URLs del form",
                                          filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
        if path:
            url_excel_path_var.set(path)
            _refresh_url_excel_preview()
            refresh_execute_state()

    Button(urls_col, text=" Seleccionar Excel de URLs", image=get_button_icon("folder_yellow.png"),
           compound="left", font=("Segoe UI", 9, "bold"), bg=BTN_ACTIVE, fg="white", relief="flat", bd=0,
           padx=16, pady=6, cursor="hand2", command=_pick_url_excel).pack(anchor="w", padx=15, pady=(0, 2))
    Label(urls_col, textvariable=url_excel_path_var, font=("Segoe UI", 8), bg=CARD_BG,
          fg=TEXT_S, wraplength=320, justify="left").pack(anchor="w", padx=15)

    # Modo fijo: por cada fila decide solo — landing+form si la columna 'URL' tiene valor,
    # o solo form si viene vacía. (Ya no hay toggle manual.)
    url_mode_var = StringVar(value="landing_form")

    fmt_val_lbl = Label(urls_col, text="Toma 'URL' (landing) + 'Formulario' por fila; si 'URL' está vacía, usa solo el form.",
                         font=("Segoe UI", 8, "bold"), bg=CARD_BG, fg="#C5A9DF", wraplength=320, justify="left")
    fmt_val_lbl.pack(anchor="w", padx=15, pady=(4, 0))
    url_warn_var = StringVar(value="")
    Label(urls_col, textvariable=url_warn_var, font=("Segoe UI", 8, "italic"),
          bg=CARD_BG, fg="#F8C471", wraplength=320, justify="left").pack(anchor="w", padx=15, pady=(0, 6))

    def switch_url_mode(_mode=None):
        # Compatibilidad: ya no hay toggle; el modo se decide por fila.
        _refresh_url_excel_preview()

    # Col 3 — se rellena más abajo con el bloque Modelos (models_card usa este parent)

    def _read_url_pairs_from_excel():
        """Lee el Excel de URLs y devuelve pares (landing_url, form_url) — uno por fila.
        Resuelve las columnas por nombre ('URL'→landing, 'Formulario'→form) con fallback a
        A/B. Por fila: si 'URL' (landing) tiene valor → (landing, form); si está vacía →
        ('', form) (solo form). Así se decide automáticamente sin toggle."""
        path = url_excel_path_var.get().strip()
        if not path or not os.path.exists(path):
            return []
        try:
            headers, rows = read_excel_rows(path, header_row=1)
        except Exception:
            return []
        landing_key = resolve_column(headers, "URL") or "_c0"
        form_key = resolve_column(headers, "Formulario") or "_c1"

        def _es_url(txt):
            return txt.lower().startswith(("http://", "https://"))

        pairs = []
        for r in rows:
            form_url = (r.get(form_key, "") or "").strip()
            landing_url = (r.get(landing_key, "") or "").strip()
            # Sólo filas con URLs reales: si el usuario elige por error otro Excel (ej. el
            # de dealers), sus celdas de texto no deben tratarse como formularios.
            if not _es_url(form_url):
                continue
            if landing_url and not _es_url(landing_url):
                landing_url = ""
            pairs.append((landing_url, form_url))  # landing "" si la columna viene vacía
        return pairs

    def _refresh_url_excel_preview(*_a):
        pairs = _read_url_pairs_from_excel()
        if not url_excel_path_var.get().strip():
            url_warn_var.set("")
            return
        if not pairs:
            url_warn_var.set("⚠ No se encontraron URLs de formulario en el Excel (columna 'Formulario').")
            return
        detected = _detect_url_country(" ".join(f for _l, f in pairs))
        if detected and detected != state["pais"]:
            url_warn_var.set(f"⚠ Las URLs parecen de {detected} pero tenés seleccionado {state['pais']}. "
                             "Verificá el mercado antes de ejecutar.")
        else:
            url_warn_var.set(f"✓ {len(pairs)} form(s) detectados en el Excel.")

    # Alias para el resto del código que ya llamaba a _parse_urls_text()
    def _parse_urls_text():
        return _read_url_pairs_from_excel()

    switch_url_mode("landing_form")

    # ── 3. Bloque de 2 columnas: Excel de dealers · Columnas del Excel ──────────
    blockB = Frame(root_frame, bg=root_frame.cget("bg"))
    blockB.pack(fill="x", pady=(0, 0))

    excel_card = card("📄 EXCEL DE DEALERS A CHEQUEAR", parent=blockB, side="left", expand=False)
    mini_guide(excel_card, "④ Cargá el Excel de dealers esperados, la fila de encabezados y (opcional) el filtro.")
    excel_path_var = StringVar()
    excel_row = Frame(excel_card, bg=CARD_BG)
    excel_row.pack(fill="x", padx=15, pady=(0, 4))

    def _pick_excel():
        path = filedialog.askopenfilename(title="Seleccionar Excel de dealers",
                                           filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
        if path:
            excel_path_var.set(path)

    Button(excel_row, text=" Seleccionar Excel", image=get_button_icon("folder_yellow.png"), compound="left",
           font=("Segoe UI", 8, "bold"), bg=BTN_ACTIVE, fg="white", relief="flat", bd=0,
           cursor="hand2", command=_pick_excel).pack(side="left")
    Label(excel_row, textvariable=excel_path_var, font=("Segoe UI", 8), bg=CARD_BG,
          fg=TEXT_S, wraplength=300, justify="left").pack(side="left", padx=10)

    header_row_row = Frame(excel_card, bg=CARD_BG)
    header_row_row.pack(fill="x", padx=9, pady=(4, 0))
    header_row_var = StringVar(value="1")
    labeled_entry(header_row_row, "Fila encabezados", header_row_var, width=8)

    filter_mode_row = Frame(excel_card, bg=CARD_BG)
    filter_mode_row.pack(fill="x", padx=15, pady=(4, 2))
    Label(filter_mode_row, text="Este Excel:", font=("Segoe UI", 8), bg=CARD_BG, fg=TEXT_S).pack(side="left")
    excel_filter_mode_var = StringVar(value="with_filter")

    filter_row = Frame(excel_card, bg=CARD_BG)
    filter_col_var = StringVar()
    filter_value_var = StringVar()
    labeled_entry(filter_row, "Columna filtro", filter_col_var, width=16)
    labeled_entry(filter_row, "Valor filtro", filter_value_var, width=12)

    def _refresh_filter_mode_visibility(*_a):
        if excel_filter_mode_var.get() == "with_filter":
            filter_row.pack(fill="x", padx=9, pady=4, after=filter_mode_row)
        else:
            filter_row.pack_forget()

    select_excel_filter_mode = make_single_select(
        filter_mode_row,
        [("with_filter", "Tiene columna de filtro"), ("no_filter", "No tiene filtro (usar todas las filas)")],
        excel_filter_mode_var, default="with_filter", on_select=_refresh_filter_mode_visibility,
    )

    CONDITION_LABELS = ["Incluir", "Excluir"]
    condition_mode_var = StringVar(value="Incluir")
    cond_col = Frame(filter_row, bg=CARD_BG)
    cond_col.pack(side="left", padx=6, pady=3)
    Label(cond_col, text="Condición", font=("Segoe UI", 8), bg=CARD_BG, fg=TEXT_S).pack(anchor="w")
    ttk.Combobox(cond_col, textvariable=condition_mode_var, state="readonly", width=22,
                 style="Dealer.TCombobox", values=CONDITION_LABELS, font=("Segoe UI", 10)).pack(fill="x")

    # ── 4. Columnas del Excel (columna derecha del bloque B) ────────────────────
    columns_card = card("🧭 COLUMNAS DEL EXCEL", parent=blockB, side="left", expand=True)
    mini_guide(columns_card, "⑤ Mapeá qué columna del Excel es cada dato (Dealer es obligatoria).")
    levels_label_row = Frame(columns_card, bg=CARD_BG)
    levels_label_row.pack(fill="x", padx=15, pady=(0, 2))
    Label(levels_label_row, text="Campos que tiene el form (id HTML del select):", font=("Segoe UI", 8),
          bg=CARD_BG, fg=TEXT_S).pack(anchor="w")

    levels_row = Frame(columns_card, bg=CARD_BG)
    levels_row.pack(fill="x", padx=15, pady=(0, 2))
    has_region_var = BooleanVar(value=True)
    has_city_var = BooleanVar(value=True)
    has_dealer_var = BooleanVar(value=True)
    make_toggle_pill(levels_row, "region", has_region_var)
    make_toggle_pill(levels_row, "city", has_city_var)
    make_toggle_pill(levels_row, "dealer", has_dealer_var, locked=False)

    Label(columns_card,
          text="ℹ Estos son los id tal cual aparecen en el HTML del <select> del form (region / city / "
               "dealer) — no cambian de país a país aunque el texto que ve el usuario sí (ej. en Argentina "
               "se ve como \"Provincia\" pero el id sigue siendo \"region\"). Desactivá acá el que tu form "
               "no tenga.",
          font=("Segoe UI", 8, "italic"), bg=CARD_BG, fg="#C5A9DF", wraplength=340,
          justify="left").pack(anchor="w", padx=15, pady=(2, 6))

    cols_row = Frame(columns_card, bg=CARD_BG)
    cols_row.pack(fill="x", padx=9, pady=4)
    col_region_var = StringVar(value="REGION")
    col_city_var = StringVar(value="CIUDAD")
    col_dealer_var = StringVar(value="NOMBRE")
    col_bac_var = StringVar(value="BAC")
    _e_region = labeled_entry(cols_row, "Columna Región", col_region_var)
    _e_city = labeled_entry(cols_row, "Columna Ciudad", col_city_var)
    _e_dealer = labeled_entry(cols_row, "Columna Dealer", col_dealer_var)
    _e_bac = labeled_entry(cols_row, "Columna BAC", col_bac_var)

    # Cuando un nivel se DESELECCIONA, su "Columna X" se bloquea y muestra "No aplica"
    # (así no se pide ni valida esa columna). Al reactivarlo, se restaura el valor real.
    _NO_APLICA = "No aplica"
    _col_shadow = {}  # nivel -> último valor real escrito por el usuario

    def _sync_col_entry(nivel, var, entry, activo):
        if activo:
            entry.config(state="normal")
            if var.get() == _NO_APLICA:
                var.set(_col_shadow.get(nivel, ""))
        else:
            if var.get() != _NO_APLICA:
                _col_shadow[nivel] = var.get()
            var.set(_NO_APLICA)
            entry.config(state="disabled")

    has_region_var.trace_add("write", lambda *_a: _sync_col_entry("region", col_region_var, _e_region, has_region_var.get()))
    has_city_var.trace_add("write", lambda *_a: _sync_col_entry("city", col_city_var, _e_city, has_city_var.get()))
    has_dealer_var.trace_add("write", lambda *_a: _sync_col_entry("dealer", col_dealer_var, _e_dealer, has_dealer_var.get()))
    # Estado inicial según las píldoras actuales
    _sync_col_entry("region", col_region_var, _e_region, has_region_var.get())
    _sync_col_entry("city", col_city_var, _e_city, has_city_var.get())
    _sync_col_entry("dealer", col_dealer_var, _e_dealer, has_dealer_var.get())

    opts_row = Frame(columns_card, bg=CARD_BG)
    opts_row.pack(fill="x", padx=15, pady=(0, 4))
    chk_bac_var = BooleanVar(value=False)
    Checkbutton(opts_row, text="Verificar BAC (opcional — muchos forms no exponen data-bac en el HTML)",
                variable=chk_bac_var, bg=CARD_BG, fg=TEXT_S, selectcolor=ENTRY_BG,
                activebackground=CARD_BG, wraplength=340, justify="left").pack(anchor="w")
    # La "Columna BAC" se bloquea a "No aplica" mientras "Verificar BAC" esté destildado.
    chk_bac_var.trace_add("write", lambda *_a: _sync_col_entry("bac", col_bac_var, _e_bac, chk_bac_var.get()))
    _sync_col_entry("bac", col_bac_var, _e_bac, chk_bac_var.get())
    # Extras y duplicados se buscan SIEMPRE (integrado en el flujo normal)
    find_extras_var = BooleanVar(value=True)

    field_check_block = Frame(columns_card, bg=CARD_BG)
    field_check_block.pack(fill="x", padx=15, pady=(4, 6))
    Label(field_check_block, text="Columnas adicionales a comprobar en el form",
          font=("Segoe UI", 8, "bold"), bg=CARD_BG, fg=TEXT_S).pack(anchor="w")
    Label(field_check_block,
          text="Agregá cualquier otro campo del form que quieras validar: el nombre de la columna del "
               "Excel y el id del campo tal cual aparece en el HTML del form (ej. columna \"CEP\" → id "
               "\"customer-cep\").",
          font=("Segoe UI", 8), bg=CARD_BG, fg=TEXT_S, wraplength=340, justify="left").pack(anchor="w", pady=(0, 4))

    field_check_header = Frame(field_check_block, bg=CARD_BG)
    field_check_header.pack(fill="x")
    Label(field_check_header, text="Columna del Excel", font=("Segoe UI", 8, "bold"), bg=CARD_BG,
          fg="#C5A9DF", width=22, anchor="w").pack(side="left", padx=2)
    Label(field_check_header, text="ID del campo en el form", font=("Segoe UI", 8, "bold"), bg=CARD_BG,
          fg="#C5A9DF", width=22, anchor="w").pack(side="left", padx=2)

    field_check_rows_container = Frame(field_check_block, bg=CARD_BG)
    field_check_rows_container.pack(fill="x", pady=(2, 4))

    def _add_field_check_row(column="", field_id="", active=True):
        row = Frame(field_check_rows_container, bg=CARD_BG)
        row.pack(fill="x", pady=2)
        column_var = StringVar(value=column)
        field_id_var = StringVar(value=field_id)
        Entry(row, textvariable=column_var, width=22, bg=ENTRY_BG, fg="white", font=("Segoe UI", 10),
              insertbackground="white", relief="flat").pack(side="left", padx=2, ipady=4)
        Entry(row, textvariable=field_id_var, width=22, bg=ENTRY_BG, fg="white", font=("Segoe UI", 10),
              insertbackground="white", relief="flat").pack(side="left", padx=2, ipady=4)

        widgets = {
            "frame": row,
            "column_var": column_var,
            "field_id_var": field_id_var,
            "active_var": BooleanVar(value=active),
            "pill_btn": None
        }

        def _add_to_form_fields():
            f_id = field_id_var.get().strip()
            if not f_id:
                messagebox.showwarning("Comparador Dealers", "Ingresá un ID de campo válido antes de agregarlo.")
                return
            if widgets["pill_btn"]:
                widgets["pill_btn"].destroy()
                widgets["pill_btn"] = None
            btn, _ = make_toggle_pill(levels_row, f_id, widgets["active_var"])
            widgets["pill_btn"] = btn

        if field_id.strip():
            _add_to_form_fields()

        btn_add = Button(row, text="➕ Agregar a campos del form", font=("Segoe UI", 8, "bold"), bg="#1a3d24", fg="#39c55a",
                         relief="flat", bd=0, cursor="hand2", padx=6, command=_add_to_form_fields)
        btn_add.pack(side="left", padx=4)

        def _remove():
            row.destroy()
            if widgets["pill_btn"]:
                widgets["pill_btn"].destroy()
            state["field_check_widgets"].remove(widgets)

        Button(row, text="×", font=("Segoe UI", 9, "bold"), bg="#3d1414", fg="#f85149",
               relief="flat", bd=0, width=2, cursor="hand2", command=_remove).pack(side="left", padx=4)
        state["field_check_widgets"].append(widgets)

    Button(field_check_block, text="+ Agregar columna a comprobar",
           font=("Segoe UI", 8, "bold"), bg=BTN_ACTIVE, fg="white", relief="flat", bd=0,
           cursor="hand2", command=lambda: _add_field_check_row()).pack(anchor="w", pady=2)

    # ── 4b. Modelos (3ra columna del bloque A: Navegador · URLs · Modelos) ──────
    models_card = card("🚗 MODELOS", parent=blockA, side="left", expand=True, subtitle_wrap=200)
    mini_guide(models_card, "Opcional: si el form tiene selector de Modelo (id \"models\").")
    has_models_var = BooleanVar(value=False)
    models_toggle_row = Frame(models_card, bg=CARD_BG)
    models_toggle_row.pack(fill="x", padx=15, pady=(0, 4))
    models_toggle_btn, _models_refresh = make_toggle_pill(models_toggle_row, "🚗  Tiene selector de Modelo", has_models_var)
    Label(models_card, text="⚠ Solo forms T1 (id \"models\" estándar). T2/T3 con selector distinto no soportados aún.",
          font=("Segoe UI", 8, "italic"), bg=CARD_BG, fg="#F8C471", wraplength=320,
          justify="left").pack(anchor="w", padx=15, pady=(0, 4))

    models_body = Frame(models_card, bg=CARD_BG)
    models_body.pack(fill="x", padx=15, pady=(0, 6))

    # ID del campo fijo (no editable): todos los forms de este proyecto usan "models".
    models_field_id_var = StringVar(value="models")

    models_mode_var = StringVar(value="all")
    models_mode_row = Frame(models_body, bg=CARD_BG)
    models_mode_row.pack(fill="x", pady=(0, 4))
    Label(models_mode_row, text="Modo:", font=("Segoe UI", 8), bg=CARD_BG, fg=TEXT_S).pack(side="left")

    def _refresh_models_mode_visibility(*_a):
        models_list_row.pack(fill="x", pady=(0, 2)) if models_mode_var.get() == "specific" else models_list_row.pack_forget()

    _select_models_mode = make_single_select(
        models_mode_row, [("all", "Todos los modelos"), ("specific", "Modelo(s) específico(s)")],
        models_mode_var, default="all",
    )
    # Re-wrappear _select_models_mode para además refrescar la visibilidad del input de texto
    _orig_select_models_mode = _select_models_mode

    def select_models_mode(val):
        if not has_models_var.get():
            return
        _orig_select_models_mode(val)
        _refresh_models_mode_visibility()

    for child in models_mode_row.winfo_children():
        if isinstance(child, Button):
            child.config(command=lambda v=("all" if child.cget("text") == "Todos los modelos" else "specific"):
                          select_models_mode(v))

    models_list_var = StringVar()
    models_list_row = Frame(models_body, bg=CARD_BG)
    Label(models_list_row, text="Modelos (separados por coma):", font=("Segoe UI", 8), bg=CARD_BG,
          fg=TEXT_S).pack(anchor="w")
    Entry(models_list_row, textvariable=models_list_var, bg=ENTRY_BG, fg="white", font=("Segoe UI", 10),
          insertbackground="white", relief="flat", width=60).pack(anchor="w", pady=(2, 0), ipady=4, fill="x")
    Label(models_list_row, text="ej: Onix, Tracker, S10", font=("Segoe UI", 8, "italic"), bg=CARD_BG,
          fg="#C5A9DF").pack(anchor="w")

    def _update_models_section_state(*_a):
        is_enabled = has_models_var.get()
        for child in models_mode_row.winfo_children():
            if isinstance(child, Button):
                if is_enabled:
                    child.config(state="normal", cursor="hand2")
                    val = "all" if child.cget("text") == "Todos los modelos" else "specific"
                    if models_mode_var.get() == val:
                        child.config(bg=BTN_ACTIVE, fg="white")
                    else:
                        child.config(bg=BTN_INACTIVE, fg=TEXT_S)
                else:
                    child.config(state="disabled", cursor="arrow", bg=BTN_INACTIVE, fg="#7D6695")
        
        if is_enabled:
            _refresh_models_mode_visibility()
        else:
            models_list_row.pack_forget()

    has_models_var.trace_add("write", _update_models_section_state)
    _update_models_section_state()

    # ── 5. Modo de salida ─────────────────────────────────────────────────────
    output_card = card("📦 MODO DE SALIDA")
    output_mode_var = StringVar(value="excel")
    output_row = Frame(output_card, bg=CARD_BG)
    output_row.pack(fill="x", padx=15, pady=(0, 6))
    for val, txt in (("excel", "Solo Excel"), ("caps", "Excel + Capturas")):
        ttk.Radiobutton(output_row, text=txt, value=val, variable=output_mode_var).pack(side="left", padx=8)

    # El envío por email se controla SOLO desde la barra superior de la app
    # ("Email destinatario" + "Enviar mail"), común a todas las pestañas. Antes había
    # acá un card duplicado que escribía en esa misma config global; se quitó para no
    # tener dos lugares donde configurar lo mismo. El valor se lee al ejecutar.

    # ── 6. Configuraciones guardadas ─────────────────────────────────────────
    presets_card = card("💾 CONFIGURACIONES GUARDADAS",
                         "Guardá el mapeo de columnas con el nombre que quieras para reusarlo después.")
    preset_name_var = StringVar()
    preset_selected_var = StringVar()

    presets_row1 = Frame(presets_card, bg=CARD_BG)
    presets_row1.pack(fill="x", padx=15, pady=(0, 4))
    Label(presets_row1, text="Nombre:", font=("Segoe UI", 8), bg=CARD_BG, fg=TEXT_S).pack(side="left")
    Entry(presets_row1, textvariable=preset_name_var, width=24, bg=ENTRY_BG, fg="white",
          insertbackground="white", relief="flat").pack(side="left", padx=(4, 10))

    def _save_preset():
        name = preset_name_var.get().strip()
        if not name:
            messagebox.showwarning("Comparador Dealers", "Ingresá un nombre para la configuración.")
            return
        if name in all_settings["presets"]:
            if not messagebox.askyesno(
                "Comparador Dealers",
                f"Ya existe una configuración llamada '{name}'.\n\n¿Querés sobrescribirla con los valores actuales?",
            ):
                return
        all_settings["presets"][name] = _collect_current_settings()
        _save_all_settings(all_settings)
        _refresh_preset_combo()
        preset_selected_var.set(name)
        ui_log(f"Configuración guardada como '{name}'.", "ok")

    Button(presets_row1, text="💾 Guardar configuración", font=("Segoe UI", 8, "bold"),
           bg=BTN_ACTIVE, fg="white", relief="flat", bd=0, cursor="hand2",
           command=_save_preset).pack(side="left")

    presets_row2 = Frame(presets_card, bg=CARD_BG)
    presets_row2.pack(fill="x", padx=15, pady=(0, 8))
    Label(presets_row2, text="Cargar:", font=("Segoe UI", 8), bg=CARD_BG, fg=TEXT_S).pack(side="left")
    preset_combo = ttk.Combobox(presets_row2, textvariable=preset_selected_var, state="readonly", width=30,
                                 style="Dealer.TCombobox", font=("Segoe UI", 10))
    preset_combo.pack(side="left", padx=(4, 10))

    def _refresh_preset_combo():
        preset_combo["values"] = sorted(all_settings["presets"].keys())

    def _load_preset():
        name = preset_selected_var.get().strip()
        cfg = all_settings["presets"].get(name)
        if not cfg:
            messagebox.showwarning("Comparador Dealers", "Seleccioná una configuración guardada.")
            return
        _apply_settings_dict(cfg)
        # Dejar el nombre en el campo Nombre para poder editarlo/renombrarlo y actualizar.
        preset_name_var.set(name)
        ui_log(f"Configuración '{name}' cargada. Modificá lo que quieras y usá 'Editar'.", "ok")

    def _update_preset():
        """Edita el preset seleccionado en el combo: reescribe sus valores con los actuales.
        Si el campo Nombre tiene otro texto, lo renombra (borra la clave vieja)."""
        original = preset_selected_var.get().strip()
        if not original or original not in all_settings["presets"]:
            messagebox.showwarning("Comparador Dealers",
                                   "Elegí en 'Cargar' la configuración que querés editar.")
            return
        nuevo_nombre = preset_name_var.get().strip() or original
        if nuevo_nombre != original and nuevo_nombre in all_settings["presets"]:
            if not messagebox.askyesno(
                "Comparador Dealers",
                f"Ya existe otra configuración llamada '{nuevo_nombre}'.\n\n¿Sobrescribirla?",
            ):
                return
        if nuevo_nombre != original:
            del all_settings["presets"][original]
        all_settings["presets"][nuevo_nombre] = _collect_current_settings()
        _save_all_settings(all_settings)
        _refresh_preset_combo()
        preset_selected_var.set(nuevo_nombre)
        if nuevo_nombre != original:
            ui_log(f"Configuración '{original}' actualizada y renombrada a '{nuevo_nombre}'.", "ok")
        else:
            ui_log(f"Configuración '{original}' actualizada.", "ok")

    def _delete_preset():
        name = preset_selected_var.get().strip()
        if not name or name not in all_settings["presets"]:
            messagebox.showwarning("Comparador Dealers", "Seleccioná una configuración guardada.")
            return
        if not messagebox.askyesno("Comparador Dealers", f"¿Eliminar la configuración '{name}'?"):
            return
        del all_settings["presets"][name]
        _save_all_settings(all_settings)
        _refresh_preset_combo()
        preset_selected_var.set("")
        ui_log(f"Configuración '{name}' eliminada.", "warn")

    Button(presets_row2, text="📂 Cargar", font=("Segoe UI", 8, "bold"),
           bg=BTN_ACTIVE, fg="white", relief="flat", bd=0, cursor="hand2",
           command=_load_preset).pack(side="left", padx=4)
    Button(presets_row2, text="✏ Editar", font=("Segoe UI", 8, "bold"),
           bg=BTN_ACTIVE, fg="white", relief="flat", bd=0, cursor="hand2",
           command=_update_preset).pack(side="left", padx=4)
    Button(presets_row2, text="🗑 Eliminar", font=("Segoe UI", 8, "bold"),
           bg="#3d1414", fg="#f85149", relief="flat", bd=0, cursor="hand2",
           command=_delete_preset).pack(side="left", padx=4)

    # ── Settings persistidos por país / presets ─────────────────────────────
    _LEGACY_CONDITION_MAP = {"include": "Incluir", "exclude": "Excluir"}

    def _apply_settings_dict(cfg):
        excel_path_var.set(cfg.get("excel_path", ""))
        header_row_var.set(str(cfg.get("header_row", 1)))
        select_excel_filter_mode(cfg.get("excel_filter_mode", "with_filter"))
        filter_col_var.set(cfg.get("filter_col", ""))
        filter_value_var.set(cfg.get("filter_value", ""))
        raw_condition = cfg.get("condition_mode", "Incluir")
        norm_condition = _LEGACY_CONDITION_MAP.get(raw_condition, raw_condition)
        # "Buscar extras" ya no es una condición (los extras se buscan siempre): si una
        # config vieja lo tenía guardado, se cae a "Incluir".
        if norm_condition not in CONDITION_LABELS:
            norm_condition = "Incluir"
        condition_mode_var.set(norm_condition)
        # Primero las columnas (valores reales) y DESPUÉS las píldoras: así el sync corre al
        # final y deja "No aplica" (bloqueado) en los niveles que queden deseleccionados,
        # guardando el valor real en el shadow por si se reactivan.
        col_region_var.set(cfg.get("col_region", "REGION"))
        col_city_var.set(cfg.get("col_city", "CIUDAD"))
        col_dealer_var.set(cfg.get("col_dealer", "NOMBRE"))
        col_bac_var.set(cfg.get("col_bac", "BAC"))
        _col_shadow.update({"region": col_region_var.get(), "city": col_city_var.get(),
                            "dealer": col_dealer_var.get(), "bac": col_bac_var.get()})
        has_region_var.set(bool(cfg.get("has_region", True)))
        has_city_var.set(bool(cfg.get("has_city", True)))
        has_dealer_var.set(bool(cfg.get("has_dealer", True)))
        chk_bac_var.set(bool(cfg.get("chk_bac", False)))
        find_extras_var.set(True)  # Siempre True (extras se buscan siempre)
        has_models_var.set(bool(cfg.get("has_models", False)))
        models_field_id_var.set(cfg.get("models_field_id", "models"))
        select_models_mode(cfg.get("models_mode", "all"))
        models_list_var.set(cfg.get("models_list", ""))
        url_mode_var.set(cfg.get("url_mode", "landing_form"))
        switch_url_mode(url_mode_var.get())
        url_excel_path_var.set(cfg.get("url_excel_path", ""))
        _refresh_url_excel_preview()
        browser_var.set(cfg.get("browser", "chrome"))
        viewport_var.set(cfg.get("viewport", "fullscreen"))
        ver_navegador_var.set(bool(cfg.get("ver_navegador", False)))
        output_mode_var.set(cfg.get("output_mode", "excel"))

        for row_widgets in list(state["field_check_widgets"]):
            row_widgets["frame"].destroy()
        state["field_check_widgets"] = []
        for check in cfg.get("field_checks", []):
            _add_field_check_row(check.get("column", ""), check.get("field_id", ""), active=bool(check.get("active", True)))
        refresh_execute_state()

    def _apply_country_settings(pais):
        cfg = all_settings["paises"].get(pais) or _default_country_settings(pais)
        _apply_settings_dict(cfg)

    def _col_real(nivel, var, activo):
        """Valor real de la columna: si el nivel está deseleccionado la var dice 'No aplica',
        así que se guarda el valor del shadow (lo último que escribió el usuario)."""
        return var.get() if activo else _col_shadow.get(nivel, "")

    def _collect_current_settings():
        return {
            "excel_path": excel_path_var.get(),
            "header_row": int(header_row_var.get() or 1),
            "excel_filter_mode": excel_filter_mode_var.get(),
            "filter_col": filter_col_var.get(),
            "filter_value": filter_value_var.get(),
            "condition_mode": condition_mode_var.get(),
            "has_region": has_region_var.get(),
            "has_city": has_city_var.get(),
            "has_dealer": has_dealer_var.get(),
            "col_region": _col_real("region", col_region_var, has_region_var.get()),
            "col_city": _col_real("city", col_city_var, has_city_var.get()),
            "col_dealer": _col_real("dealer", col_dealer_var, has_dealer_var.get()),
            "col_bac": _col_real("bac", col_bac_var, chk_bac_var.get()),
            "chk_bac": chk_bac_var.get(),
            "find_extras": find_extras_var.get(),
            "has_models": has_models_var.get(),
            "models_field_id": models_field_id_var.get(),
            "models_mode": models_mode_var.get(),
            "models_list": models_list_var.get(),
            "url_mode": url_mode_var.get(),
            "url_excel_path": url_excel_path_var.get(),
            "browser": browser_var.get(),
            "viewport": viewport_var.get(),
            "ver_navegador": ver_navegador_var.get(),
            "output_mode": output_mode_var.get(),
            "field_checks": [
                {
                    "column": w["column_var"].get().strip(),
                    "field_id": w["field_id_var"].get().strip(),
                    "active": bool(w["active_var"].get())
                }
                for w in state["field_check_widgets"]
                if w["column_var"].get().strip() and w["field_id_var"].get().strip()
            ],
        }

    def _persist_current_settings():
        all_settings["paises"][state["pais"]] = _collect_current_settings()
        _save_all_settings(all_settings)

    def select_country(pais):
        if state["running"]:
            return
        if state["pais"] and state["pais"] != pais:
            _persist_current_settings()
        state["pais"] = pais
        _refresh_pais_cards()
        _apply_country_settings(pais)

    _refresh_preset_combo()

    # ── Lógica de ejecución ──────────────────────────────────────────────────
    def _column_map():
        headers = state["headers"]
        return {
            "region": resolve_column(headers, col_region_var.get()) if has_region_var.get() else None,
            "city": resolve_column(headers, col_city_var.get()) if has_city_var.get() else None,
            "dealer": resolve_column(headers, col_dealer_var.get()) if has_dealer_var.get() else None,
            "bac": resolve_column(headers, col_bac_var.get()) if chk_bac_var.get() else None,
        }

    def _field_checks_resolved():
        headers = state["headers"]
        out = []
        for w in state["field_check_widgets"]:
            if not w["active_var"].get():
                continue
            field_id = w["field_id_var"].get().strip()
            col_key = resolve_column(headers, w["column_var"].get())
            if field_id and col_key:
                out.append({"field_id": field_id, "column": col_key})
        return out

    def _load_and_filter_excel():
        excel_path = excel_path_var.get().strip()
        if not excel_path or not os.path.exists(excel_path):
            raise ValueError("Seleccioná un Excel de dealers válido.")
        try:
            header_row = int(header_row_var.get() or 1)
        except ValueError:
            raise ValueError("La fila de encabezado debe ser un número.")

        headers, rows = read_excel_rows(excel_path, header_row=header_row)
        state["headers"] = headers
        state["rows"] = rows

        if excel_filter_mode_var.get() == "no_filter":
            filtered = list(rows)
        else:
            filter_col_key = resolve_column(headers, filter_col_var.get())
            # Tanto Incluir como Excluir toman las filas que MATCHEAN el valor del filtro
            # (ej. POSVENTA == "si" o == "no"). La diferencia es la verificación:
            #   Incluir → el dealer DEBE estar en el form.
            #   Excluir → el dealer NO debe estar (se chequea ausencia, ver expect_absent).
            filtered = filter_rows(rows, filter_col_key, filter_value_var.get(), "include")
        state["filtered_rows"] = filtered
        state["expect_absent"] = (
            excel_filter_mode_var.get() != "no_filter" and condition_mode_var.get() == "Excluir"
        )

        # Filas OCULTAS del Excel: NO cuentan como esperadas. En Incluir se apartan del set
        # a comparar y pasan a find_extra_dealers para que, si algo del form solo está
        # declarado en una fila oculta, se reporte como OCULTO (naranja). En Excluir se
        # procesan igual pero sus resultados se marcan en naranja.
        state["hidden_filtered"] = []
        try:
            # dict {nº fila: 'filtrada' | 'oculta'}
            hidden = detect_hidden_rows(excel_path, header_row=header_row)
            state["hidden_rows"] = hidden
            hidden_in_filter = [r for r in filtered if r.get("__row__") in hidden]
            if hidden_in_filter:
                state["hidden_filtered"] = hidden_in_filter
                if not state["expect_absent"]:
                    filtered = [r for r in filtered if r.get("__row__") not in hidden]
                    state["filtered_rows"] = filtered
                dealer_key = resolve_column(headers, col_dealer_var.get()) if has_dealer_var.get() else None
                n_filt = sum(1 for r in hidden_in_filter if hidden.get(r.get("__row__")) == "filtrada")
                n_ocul = len(hidden_in_filter) - n_filt
                nombres = ", ".join(
                    (f"fila {r.get('__row__')} [{hidden.get(r.get('__row__'), 'oculta')}]"
                     + (f": {r.get(dealer_key, '') or '?'}" if dealer_key else ""))
                    for r in hidden_in_filter[:15]
                )
                mas = "…" if len(hidden_in_filter) > 15 else ""
                detalle_tipo = []
                if n_filt:
                    detalle_tipo.append(f"{n_filt} filtrada(s) por un AutoFilter")
                if n_ocul:
                    detalle_tipo.append(f"{n_ocul} oculta(s) a mano")
                ui_log(
                    f"⚠ {len(hidden_in_filter)} fila(s) del filtro no se ven en el Excel "
                    f"({' y '.join(detalle_tipo)}) → {nombres}{mas}. "
                    + ("No cuentan como esperadas: si algo de eso aparece en el form se reporta como OCULTO (naranja)."
                       if not state["expect_absent"] else "Se verifican igual y sus resultados van en naranja."),
                    "warn",
                )
        except Exception:
            state["hidden_rows"] = {}

        # Disclaimer de columnas OCULTAS en el Excel de dealers.
        try:
            hidden_cols = detect_hidden_columns(excel_path, header_row=header_row)
            state["hidden_columns"] = hidden_cols
            if hidden_cols:
                cols_txt = ", ".join(
                    f"{hc['column']}" + (f" ({hc['header']})" if hc.get("header") else "") for hc in hidden_cols
                )
                ui_log(f"⚠ DISCLAIMER: el Excel tiene columna(s) OCULTAS: {cols_txt}. Se procesan igual.", "warn")
        except Exception:
            state["hidden_columns"] = []

        return headers, filtered

    def _should_find_extras():
        return True  # Siempre buscar extras y duplicados como parte del chequeo normal

    # ── Modal de ejecución (mismo patrón que "Envío de Leads") ────────────────
    def _open_run_modal():
        modal = Toplevel(root)
        modal.overrideredirect(True)  # Quitar bordes de Windows
        modal_width, modal_height = 460, 300
        px = root.winfo_rootx() + (root.winfo_width() - modal_width) // 2
        py = root.winfo_rooty() + (root.winfo_height() - modal_height) // 2
        modal.geometry(f"{modal_width}x{modal_height}+{px}+{py}")
        modal.configure(bg=MODAL_BG, bd=1, highlightthickness=1, highlightbackground=BORDER)

        # Al cerrar el modal por CUALQUIER motivo, el botón EJECUTAR vuelve a su estado original.
        def _reset_exec_btn(_e=None):
            try:
                btn_run.config(text=" EJECUTAR", image=get_button_icon("play_green.png"))
                refresh_execute_state()
            except Exception:
                pass
        modal.bind("<Destroy>", lambda e: _reset_exec_btn() if str(e.widget) == str(modal) else None, add="+")

        def on_cerrar():
            _reset_exec_btn()
            modal.destroy()
            ui_log("Ventana de ejecución cerrada.", "info")

        # Modal real: NO usamos grab_set() para permitir que el usuario minimice o cierre
        # la ventana principal desde la barra de título de Windows. En su lugar, deshabilitamos
        # la interacción con el área cliente del main window agregando BlockTag a sus widgets.
        modal.transient(root)
        modal.attributes("-topmost", False)
        modal.lift()
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
            if "Comparando" in title_lbl.cget("text"):
                if messagebox.askyesno("Detener ejecución", "¿Querés detener la comparación y cerrar la ventana?"):
                    state["stop_event"].set()
                    _reset_exec_btn()
                    modal.destroy()
            else:
                on_cerrar()

        def on_root_close_request():
            if modal.winfo_exists() and "Comparando" in title_lbl.cget("text"):
                if messagebox.askyesno("Salir", "Hay una comparación en curso. ¿Querés detenerla y salir de la app?"):
                    state["stop_event"].set()
                    modal.destroy()
                    root.destroy()
            else:
                if modal.winfo_exists():
                    modal.destroy()
                root.destroy()

        orig_close_protocol = root.protocol("WM_DELETE_WINDOW")
        root.protocol("WM_DELETE_WINDOW", on_root_close_request)

        # Binds para seguir minimizar/restaurar/foco de la ventana principal
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

        # Custom Title Bar con minimizar y cerrar
        title_bar = Frame(modal, bg=MODAL_BG)
        title_bar.pack(fill="x", padx=15, pady=(5, 0))
        title_bar_lbl = Label(title_bar, text="Comparador Dealers", font=("Segoe UI", 8, "bold"),
                               bg=MODAL_BG, fg="#C5A9DF")
        title_bar_lbl.pack(side="left")

        # Hacer el modal arrastrable/movible
        def _start_drag(event):
            modal._drag_x = event.x
            modal._drag_y = event.y

        def _drag(event):
            x = modal.winfo_x() - modal._drag_x + event.x
            y = modal.winfo_y() - modal._drag_y + event.y
            modal.geometry(f"+{x}+{y}")

        title_bar.bind("<Button-1>", _start_drag)
        title_bar.bind("<B1-Motion>", _drag)
        title_bar_lbl.bind("<Button-1>", _start_drag)
        title_bar_lbl.bind("<B1-Motion>", _drag)

        btn_cls = Button(title_bar, text="✕", font=("Segoe UI", 8, "bold"), bg=MODAL_BG, fg="#C5A9DF",
                          relief="flat", bd=0, cursor="hand2", padx=6, pady=2, command=on_close_modal)
        btn_cls.pack(side="right")
        btn_cls.bind("<Enter>", lambda e: btn_cls.config(bg="#E74C3C", fg="white"))
        btn_cls.bind("<Leave>", lambda e: btn_cls.config(bg=MODAL_BG, fg="#C5A9DF"))

        btn_min = Button(title_bar, text="—", font=("Segoe UI", 8, "bold"), bg=MODAL_BG, fg="#C5A9DF",
                          relief="flat", bd=0, cursor="hand2", padx=6, pady=2, command=root.iconify)
        btn_min.pack(side="right", padx=2)
        btn_min.bind("<Enter>", lambda e: btn_min.config(bg="#38234D"))
        btn_min.bind("<Leave>", lambda e: btn_min.config(bg=MODAL_BG))

        # 1. Header (Comparando / Completo)
        header = Frame(modal, bg=MODAL_BG)
        header.pack(fill="x", padx=20, pady=(5, 10))
        icon_lbl = Label(header, text="↻", font=("Segoe UI", 16, "bold"), bg=MODAL_BG, fg="#C5A9DF")
        icon_lbl.pack(side="left")
        title_info = Frame(header, bg=MODAL_BG)
        title_info.pack(side="left", padx=10)
        title_lbl = Label(title_info, text="Comparando dealers...", font=("Segoe UI", 12, "bold"),
                           bg=MODAL_BG, fg="white")
        title_lbl.pack(anchor="w")
        subtitle_lbl = Label(title_info, text=f"{state['pais']}", font=("Segoe UI", 9),
                              bg=MODAL_BG, fg=TEXT_S)
        subtitle_lbl.pack(anchor="w")

        # Animación de rotación del icono
        rotation_glyphs = ["↻", "➔", "↻", "➔"]

        def _rotate_icon(idx=0):
            if modal.winfo_exists() and "Comparando" in title_lbl.cget("text"):
                icon_lbl.config(text=rotation_glyphs[idx % len(rotation_glyphs)])
                modal.after(250, lambda: _rotate_icon(idx + 1))
        _rotate_icon()

        def _on_detener():
            # Mismo comportamiento que la X de cerrar: si hay una comparación en curso,
            # pide confirmación antes de cortar. on_close_modal ya maneja el confirm.
            on_close_modal()

        btn_detener = Button(header, text=" Detener", image=get_button_icon("stop_coral.png"), compound="left",
                              font=("Segoe UI", 9, "bold"), bg="#3D1220", fg="#F1948A", relief="flat", bd=0,
                              highlightthickness=1, highlightbackground="#F1948A", cursor="hand2",
                              command=_on_detener, padx=12, pady=4)
        btn_detener.pack(side="right")
        btn_detener.bind("<Enter>", lambda e: btn_detener.config(bg="#5E1D31") if btn_detener["state"] == "normal" else None)
        btn_detener.bind("<Leave>", lambda e: btn_detener.config(bg="#3D1220") if btn_detener["state"] == "normal" else None)

        progress_row = Frame(modal, bg=MODAL_BG)
        progress_row.pack(fill="x", padx=20, pady=(0, 6))
        pb_canvas = Canvas(progress_row, height=8, bg="#35164D", highlightthickness=0)
        pb_canvas.pack(fill="x")
        pb_fill = pb_canvas.create_rectangle(0, 0, 0, 8, fill="#F8C471", width=0)

        status_lbl = Label(modal, text="Iniciando...", font=("Segoe UI", 8), bg=MODAL_BG, fg=TEXT_S,
                            wraplength=420, justify="left")
        status_lbl.pack(anchor="w", padx=20, pady=(2, 6))

        summary_lbl = Label(modal, text="", font=("Segoe UI", 10, "bold"), bg=MODAL_BG, fg="white")
        summary_lbl.pack(anchor="w", padx=20, pady=(4, 4))

        btn_close = Button(modal, text="Cerrar resultados", font=("Segoe UI", 10, "bold"), bg="#AED6F1", fg="#110518",
                            relief="flat", bd=0, cursor="hand2", pady=6, command=on_cerrar)
        btn_close.bind("<Enter>", lambda e: btn_close.config(bg="#D4E6F1"))
        btn_close.bind("<Leave>", lambda e: btn_close.config(bg="#AED6F1"))

        def _set_progress(cur, total, label_txt):
            def _apply():
                if not modal.winfo_exists():
                    return
                w = max(1, pb_canvas.winfo_width())
                frac = (cur / total) if total else 0
                pb_canvas.coords(pb_fill, 0, 0, int(w * max(0.0, min(1.0, frac))), 8)
                status_lbl.config(text=f"{cur}/{total} — {label_txt}")
            root.after(0, _apply)

        def _set_complete(ok, counts, detenido=False, error_msg=None):
            def _apply():
                if not modal.winfo_exists():
                    return
                # Defensivo: pase lo que pase arriba, el botón Cerrar SIEMPRE debe aparecer,
                # para que el modal nunca quede tapando el resto de la pestaña sin forma de cerrarlo.
                try:
                    try:
                        btn_detener.pack_forget()
                    except Exception:
                        pass
                    if detenido:
                        icon_lbl.config(text="■", fg="#F1948A")
                        title_lbl.config(text="Comparación detenida")
                    elif error_msg:
                        icon_lbl.config(text="✕", fg="#F1948A")
                        title_lbl.config(text="Comparación con error")
                        status_lbl.config(text=error_msg, fg="#F1948A")
                    else:
                        icon_lbl.config(text="✓", fg="#82E0AA")
                        title_lbl.config(text="Comparación completada")
                    w = max(1, pb_canvas.winfo_width())
                    pb_canvas.coords(pb_fill, 0, 0, w, 8)
                    pb_canvas.itemconfig(pb_fill, fill="#F1948A" if (counts.get("FAIL") or error_msg or detenido)
                                          else "#82E0AA")
                    _total = sum(counts.get(k, 0) for k in ("PASS", "FAIL", "EXTRA", "DUPLICADO", "OCULTO", "NOTA"))
                    # Línea 1: PASS/FAIL bien grande. Línea 2: el resto de estados solo si hay.
                    resumen = f"🟢 {counts.get('PASS', 0)} PASS      🔴 {counts.get('FAIL', 0)} FAIL"
                    extras_line = []
                    if counts.get("EXTRA"):
                        extras_line.append(f"🟡 {counts['EXTRA']} EXTRA")
                    if counts.get("DUPLICADO"):
                        extras_line.append(f"🔵 {counts['DUPLICADO']} DUPLICADO")
                    if counts.get("OCULTO"):
                        extras_line.append(f"🟠 {counts['OCULTO']} OCULTO")
                    if counts.get("NOTA"):
                        extras_line.append(f"🔷 {counts['NOTA']} NOTA")
                    if extras_line:
                        resumen += "\n" + "   ".join(extras_line)
                    resumen += f"\nTotal chequeado: {_total}"
                    if not detenido and not error_msg:
                        resumen += ("   —   ✓ Todo OK" if not counts.get("FAIL")
                                    else f"   —   ⚠ {counts.get('FAIL', 0)} con problemas, revisá el Excel")
                    summary_lbl.config(text=resumen, justify="left")
                except Exception as ui_err:  # noqa: BLE001
                    LOGGER.warning("Error actualizando modal de resultados: %s", ui_err)
                finally:
                    btn_close.pack(fill="x", padx=20, pady=(10, 10))
            root.after(0, _apply)

        state["modal"] = {"window": modal, "set_progress": _set_progress, "set_complete": _set_complete}
        return state["modal"]

    def _progress_cb(cur, total, label_txt):
        modal = state.get("modal")
        if modal:
            modal["set_progress"](cur, total, label_txt)

    def _set_running(running):
        state["running"] = running
        if running:
            btn_run.config(text="  EJECUCIÓN EN CURSO", image="", bg=BTN_INACTIVE, fg="#8A7E9E",
                           disabledforeground="#8A7E9E", state="disabled", cursor="arrow")
        else:
            btn_run.config(text=" EJECUTAR", image=get_button_icon("play_green.png"), state="normal", cursor="hand2")
            refresh_execute_state()

    def _tiene_datos_minimos():
        """Chequeo rápido de lo mínimo para poder ejecutar: Excel de dealers elegido,
        columna Dealer definida (solo si está activo), y Excel de URLs cargado."""
        if not excel_path_var.get().strip():
            return False
        if has_dealer_var.get() and not col_dealer_var.get().strip():
            return False
        if not url_excel_path_var.get().strip() or not os.path.exists(url_excel_path_var.get().strip()):
            return False
        return True

    def refresh_execute_state():
        if state.get("running"):
            return
        if _tiene_datos_minimos():
            btn_run.config(state="normal", bg=EXECUTE_BG, fg=EXECUTE_FG, cursor="hand2")
        else:
            btn_run.config(state="disabled", bg=BTN_INACTIVE, fg="#9B86B5", cursor="arrow")

    def _validate_and_prepare():
        """Corre en el hilo principal, ANTES de abrir el modal de ejecución. Si algo está mal
        (Excel no seleccionado, columna faltante, URL vacía) tira ValueError con el motivo y
        nunca se llega a abrir el modal — así un error de tipeo no deja la pestaña bloqueada."""
        headers, filtered_rows = _load_and_filter_excel()
        ui_log(f"Excel cargado: {len(state['rows'])} filas, {len(filtered_rows)} tras filtro.", "info")
        if not filtered_rows:
            raise ValueError("No hay filas para comparar después del filtro.")

        column_map = _column_map()
        if has_dealer_var.get() and not column_map["dealer"]:
            raise ValueError(f"No se encontró la columna de dealer '{col_dealer_var.get()}' en el Excel.")
        if has_region_var.get() and not column_map["region"]:
            raise ValueError(f"No se encontró la columna de región '{col_region_var.get()}' en el Excel.")
        if has_city_var.get() and not column_map["city"]:
            raise ValueError(f"No se encontró la columna de ciudad '{col_city_var.get()}' en el Excel.")

        # Pares (landing, form) del Excel de URLs, ya filtrados a URLs reales por
        # _read_url_pairs_from_excel; acá se descartan duplicados exactos para no generar
        # dos carpetas de reporte idénticas.
        url_pairs = []
        _vistos = set()
        for landing, form_url in _parse_urls_text():
            if not form_url:
                continue
            clave = (landing.strip().lower(), form_url.strip().lower())
            if clave in _vistos:
                continue
            _vistos.add(clave)
            url_pairs.append((landing, form_url))
        if not url_pairs:
            raise ValueError("Ingresá al menos una URL de form en el bloque de URLs.")

        return filtered_rows, column_map, url_pairs

    def _extract_form_slug(url):
        """Extrae un slug legible de la URL del form (ej. 'colision' de '.../gm_forms/colision')."""
        import re as _re
        url = (url or "").strip().rstrip("/").split("?")[0].split("#")[0]
        parts = url.rstrip("/").split("/")
        slug = parts[-1] if parts else "form"
        slug = _re.sub(r"[^\w-]", "_", slug)[:40]
        return slug or "form"

    def _slug(txt):
        import re as _re_sub
        return _re_sub.sub(r"[^\w]+", "_", str(txt or "").strip().lower()).strip("_")

    def _pair_output_dir(form_url, stamp):
        """Carpeta propia por form: país + form + columna de filtro (o 'sinfiltro') +
        incluidos/excluidos + timestamp. Ahí quedan el Excel y las capturas juntos, sin ZIP."""
        form_slug = _extract_form_slug(form_url)
        if excel_filter_mode_var.get() == "with_filter" and filter_col_var.get().strip():
            filtro_slug = _slug(filter_col_var.get())
        else:
            filtro_slug = "sinfiltro"
        incl_excl = "excluidos" if state.get("expect_absent") else "incluidos"
        name = "_".join(p for p in (_slug(state["pais"]), form_slug, filtro_slug, incl_excl, stamp) if p)
        path = os.path.join(_get_results_dir(), name)
        os.makedirs(path, exist_ok=True)
        return path

    def _worker_run(filtered_rows, column_map, url_pairs):
        driver = None
        counts = {"PASS": 0, "FAIL": 0, "EXTRA": 0, "DUPLICADO": 0, "OCULTO": 0, "NOTA": 0}
        error_msg = None

        # Cargar config de visible_browser y pausar_autenticacion
        _ver_navegador = False
        _pausar_auth = False
        try:
            from .helpers_interface import cargar_config_global
            _cfg = cargar_config_global().get("ui_prefs", {})
            _ver_navegador = bool(_cfg.get("visible_browser", False))
            _pausar_auth = bool(_cfg.get("pausar_autenticacion", False))
        except Exception:
            pass

        try:
            driver = BrowserManager.create_browser(
                browser_type=browser_var.get(), viewport=viewport_var.get(),
                headless=False, background=not _ver_navegador,
            )
            state["driver"] = driver

            output_mode = output_mode_var.get()
            total_pairs = len(url_pairs)
            current_url_mode = url_mode_var.get()
            expect_absent = state.get("expect_absent", False)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            all_results_all_pairs = []
            report_paths = []

            for pair_idx, (landing_url, form_url) in enumerate(url_pairs, start=1):
                if state["stop_event"].is_set():
                    raise StopRequested()
                ui_log(f"\n=== Form {pair_idx}/{total_pairs}: {form_url} ===", "info")
                if landing_url:
                    ui_log(f"    Landing: {landing_url}", "info")

                pair_dir = _pair_output_dir(form_url, stamp)

                # --- FASE 1: comparación rápida (sin capturas) ---
                open_target(driver, current_url_mode, landing_url, form_url)

                if pair_idx == 1 and _pausar_auth:
                    from tkinter import messagebox
                    ui_log("PAUSA DE AUTENTICACIÓN: Esperando que el usuario inicie sesión...", "info")
                    res = messagebox.askokcancel(
                        "Ocioso — Autenticación Manual",
                        "Se ha pausado la comparación de dealers antes del primer formulario para que puedas iniciar sesión.\n\n"
                        "1. Completá la autenticación / login en la ventana del navegador.\n"
                        "2. Asegurate de estar en la página del formulario.\n"
                        "3. Hacé click en 'Aceptar' para continuar con la comparación, o 'Cancelar' para detener el proceso.",
                        parent=None
                    )
                    if not res:
                        raise StopRequested("Ejecución cancelada por el usuario durante la pausa de autenticación.")

                    ui_log("Resumiendo comparación. Buscando contexto del formulario...", "info")
                    # Ya autenticado: mandar el navegador fuera de pantalla para que la
                    # comparación siga en segundo plano y no le robe el foco al usuario.
                    try:
                        driver.set_window_position(10000, 0)
                    except Exception:
                        pass
                    driver.switch_to.default_content()
                    iframe = _locate_form_iframe(driver, form_url, wait_seconds=8)
                    if iframe is not None:
                        driver.switch_to.frame(iframe)
                        ui_log("Contexto cambiado al iframe del formulario.", "info")

                # Forms multi-paso: avanzar hasta el paso donde están los dropdowns
                # region/city/dealer (en los de 1 paso los encuentra al instante).
                advance_to_selects(
                    driver, level_ids=DEFAULT_SELECT_IDS,
                    has_region=has_region_var.get(), has_city=has_city_var.get(),
                    has_dealer=has_dealer_var.get(), log_cb=ui_log, stop_flag=state["stop_event"],
                )

                model_field_id = None
                models_to_run = None
                if has_models_var.get():
                    model_field_id = models_field_id_var.get().strip() or "models"
                    if models_mode_var.get() == "specific":
                        models_to_run = [m.strip() for m in models_list_var.get().split(",") if m.strip()]
                        if not models_to_run:
                            ui_log("No ingresaste ningún modelo específico, se omite el filtro por modelo.", "warn")
                    else:
                        models_to_run = list_model_options(driver, model_field_id)
                        ui_log(f"Modelos detectados en el form: {', '.join(models_to_run) or '(ninguno)'}", "info")

                def _compare_progress_cb(cur, total, label_txt, _pair_idx=pair_idx):
                    _progress_cb(cur, total, f"[Form {_pair_idx}/{total_pairs}] Comparando {cur}/{total}: {label_txt}")

                def _extras_progress_cb(cur, total, label_txt, _pair_idx=pair_idx):
                    _progress_cb(cur, total, f"[Form {_pair_idx}/{total_pairs}] Buscando extras {cur}/{total}: {label_txt}")

                pair_results = compare_dealers(
                    driver, filtered_rows, column_map,
                    level_ids=DEFAULT_SELECT_IDS,
                    has_region=has_region_var.get(), has_city=has_city_var.get(),
                    has_dealer=has_dealer_var.get(),
                    chk_bac=chk_bac_var.get(), field_checks=_field_checks_resolved(),
                    model_field_id=model_field_id, models=models_to_run,
                    log_cb=ui_log, progress_cb=_compare_progress_cb, stop_flag=state["stop_event"],
                    screenshot_cb=None,
                    expect_absent=expect_absent,
                )
                _hidden_filas = state.get("hidden_rows") or {}
                for r in pair_results:
                    r["url_form"] = form_url
                    r["url_landing"] = landing_url if current_url_mode == "landing_form" else ""
                    # Resultados de filas no visibles del Excel (modo Excluir): naranja en el
                    # reporte, distinguiendo filtrada (AutoFilter) vs oculta (a mano).
                    if r.get("fila") in _hidden_filas:
                        r["hidden_in_excel"] = True
                        _tipo = _hidden_filas.get(r.get("fila"), "oculta")
                        r["hidden_kind"] = _tipo
                        r.setdefault("fails", []).append(
                            f"⚠ La fila del Excel está {'FILTRADA (AutoFilter)' if _tipo == 'filtrada' else 'OCULTA (a mano)'}"
                        )

                if _should_find_extras() and not expect_absent:
                    ui_log("Buscando EXTRAS y DUPLICADOS en el form...", "info")
                    extras = find_extra_dealers(
                        driver, filtered_rows, column_map, level_ids=DEFAULT_SELECT_IDS,
                        has_region=has_region_var.get(), has_city=has_city_var.get(),
                        log_cb=ui_log, progress_cb=_extras_progress_cb, stop_flag=state["stop_event"],
                        has_dealer=has_dealer_var.get(),
                        hidden_rows_data=state.get("hidden_filtered") or [],
                    )
                    for r in extras:
                        r["url_form"] = form_url
                        r["url_landing"] = landing_url if current_url_mode == "landing_form" else ""
                    pair_results = pair_results + extras

                for r in pair_results:
                    counts[r["status"]] = counts.get(r["status"], 0) + 1
                all_results_all_pairs.extend(pair_results)
                state["results"] = all_results_all_pairs
                ui_log(
                    f"Form {pair_idx}/{total_pairs} comparado: "
                    f"PASS={sum(1 for r in pair_results if r['status']=='PASS')} "
                    f"FAIL={sum(1 for r in pair_results if r['status']=='FAIL')} "
                    f"EXTRA={sum(1 for r in pair_results if r['status']=='EXTRA')} "
                    f"DUPLICADO={sum(1 for r in pair_results if r['status']=='DUPLICADO')}", "ok",
                )

                # --- Exportar el reporte de ESTE form ANTES de las capturas (la comparación
                # rápida ya terminó; así el Excel queda disponible aunque las capturas tarden
                # o se detengan a mitad de camino) ---
                excel_filename = os.path.basename(pair_dir) + ".xlsx"
                export_path = export_results_excel(
                    pair_results, output_path=os.path.join(pair_dir, excel_filename), pais=state["pais"],
                    hidden_rows=state.get("hidden_rows"), hidden_columns=state.get("hidden_columns"),
                )
                report_paths.append(export_path)
                ui_log(f"Reporte Form {pair_idx}/{total_pairs}: {export_path}", "ok")

                # --- FASE 2 (opcional): capturas — mismo form, misma carpeta que el Excel ---
                if output_mode == "caps" and not state["stop_event"].is_set():
                    ui_log(f"Generando capturas para Form {pair_idx}/{total_pairs}...", "info")
                    open_target(driver, current_url_mode, landing_url, form_url)
                    if pair_idx == 1 and _pausar_auth:
                        driver.switch_to.default_content()
                        iframe = _locate_form_iframe(driver, form_url, wait_seconds=8)
                        if iframe is not None:
                            driver.switch_to.frame(iframe)

                    # Mismo avance de pasos que en Fase 1, para llegar a los dropdowns
                    advance_to_selects(
                        driver, level_ids=DEFAULT_SELECT_IDS,
                        has_region=has_region_var.get(), has_city=has_city_var.get(),
                        has_dealer=has_dealer_var.get(), log_cb=lambda *_a, **_k: None,
                        stop_flag=state["stop_event"],
                    )

                    def _shot(result, _form_url=form_url, _landing_url=landing_url,
                              _url_mode=current_url_mode, _pair_dir=pair_dir, _pair_idx=pair_idx):
                        combo_parts = []
                        if has_region_var.get() and result.get("region"):
                            combo_parts.append(result["region"])
                        if has_city_var.get() and result.get("city"):
                            combo_parts.append(result["city"])
                        if has_dealer_var.get() and result.get("dealer"):
                            combo_parts.append(result["dealer"])
                        if not combo_parts:
                            combo_parts.append("dealer")

                        import re
                        safe_parts = []
                        for part in combo_parts:
                            clean = re.sub(r"[^\w\s\.-]+", "", str(part)).strip()
                            clean = re.sub(r"\s+", "_", clean)
                            if clean:
                                safe_parts.append(clean)
                        combo_name = "_".join(safe_parts)[:60]
                        filename = f"{result['status']}_fila{result.get('fila') or 'x'}_{combo_name}.png"
                        shot_landing = _landing_url if _url_mode == "landing_form" else ""
                        banner_extra = banner_lines_for_result(
                            result, has_region=has_region_var.get(),
                            has_city=has_city_var.get(), has_dealer=has_dealer_var.get(),
                        )

                        if expect_absent:
                            # Evidencia de exclusión: se despliega el PRIMER nivel de la cadena
                            # región→ciudad→dealer que no se encontró (si la región no está,
                            # tampoco pueden estar su ciudad ni su dealer). El <select> nativo
                            # no es fotografiable, así que se clona la lista completa de
                            # opciones arriba del form, resaltando el buscado si aparece.
                            nivel, sel_key, buscado, found = evidence_level_for_result(
                                result, has_region=has_region_var.get(),
                                has_city=has_city_var.get(), has_dealer=has_dealer_var.get(),
                            )
                            capture_dropdown_evidence(
                                driver, DEFAULT_SELECT_IDS[sel_key], _pair_dir, filename,
                                title=f"Dropdown {nivel} desplegado", searched_text=buscado,
                                found=found,
                                form_url=_form_url, landing_url=shot_landing,
                                extra_lines=banner_extra,
                            )
                        else:
                            capture_result_screenshot(
                                driver, _pair_dir, filename,
                                form_url=_form_url, landing_url=shot_landing,
                                extra_lines=banner_extra,
                            )

                    def _capture_progress_cb(cur, total, label_txt, _pair_idx=pair_idx):
                        _progress_cb(cur, total, f"[Form {_pair_idx}/{total_pairs}] Capturando {cur}/{total}: {label_txt}")

                    compare_dealers(
                        driver, filtered_rows, column_map,
                        level_ids=DEFAULT_SELECT_IDS,
                        has_region=has_region_var.get(), has_city=has_city_var.get(),
                        has_dealer=has_dealer_var.get(),
                        chk_bac=chk_bac_var.get(), field_checks=_field_checks_resolved(),
                        model_field_id=model_field_id, models=models_to_run,
                        log_cb=lambda *_a: None, progress_cb=_capture_progress_cb, stop_flag=state["stop_event"],
                        screenshot_cb=_shot,
                        expect_absent=expect_absent,
                    )

                # --- Email por form: usa el flag global de la barra superior ---
                from .helpers_interface import cargar_config_global as _cfg_g
                if bool(_cfg_g().get("enviar_mail", False)):
                    try:
                        from .helpers_interface import enviar_email_comparador_dealers
                        _pair_counts = {}
                        for r in pair_results:
                            _pair_counts[r["status"]] = _pair_counts.get(r["status"], 0) + 1
                        enviado = enviar_email_comparador_dealers(
                            state["pais"], pair_dir, export_path, _pair_counts,
                            form_url=form_url, landing_url=landing_url,
                            incluir_capturas=(output_mode == "caps"),
                        )
                        ui_log(f"Email {'encolado' if enviado else 'NO enviado'} para Form {pair_idx}/{total_pairs}.",
                               "ok" if enviado else "warn")
                    except Exception as _mail_err:  # noqa: BLE001
                        ui_log(f"No se pudo enviar el email del Form {pair_idx}: {_mail_err}", "err")

            ui_log(
                f"\nTodos los forms procesados. Reportes generados: {len(report_paths)}\n" +
                "\n".join(f"  - {p}" for p in report_paths), "ok",
            )

        except StopRequested:
            ui_log("Ejecución detenida por el usuario.", "warn")
        except Exception as e:  # noqa: BLE001
            LOGGER.exception("Error en Comparador Dealers")
            error_msg = str(e)
            ui_log(f"Error: {e}", "err")
        finally:
            if driver is not None:
                try:
                    driver.quit()
                except Exception:
                    pass
            state["driver"] = None
            modal = state.get("modal")
            if modal:
                modal["set_complete"](not error_msg, counts, detenido=state["stop_event"].is_set(),
                                       error_msg=error_msg)
            root.after(0, lambda: _set_running(False))

    def _run_or_stop():
        if state["running"]:
            state["stop_event"].set()
            ui_log("Deteniendo...", "warn")
            return
        try:
            filtered_rows, column_map, url_pairs = _validate_and_prepare()
        except ValueError as e:
            messagebox.showwarning("Comparador Dealers", str(e))
            return
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Comparador Dealers", f"No se pudo leer el Excel:\n{e}")
            return
        _persist_current_settings()
        state["stop_event"].clear()
        _set_running(True)
        _open_run_modal()
        threading.Thread(
            target=_worker_run, args=(filtered_rows, column_map, url_pairs), daemon=True
        ).start()



    btn_run = Button(footer_btns, text=" EJECUTAR", image=get_button_icon("play_green.png"), compound="left",
                      font=("Segoe UI", 9, "bold"), bg=EXECUTE_BG, fg=EXECUTE_FG,
                      relief="flat", bd=0, activebackground=EXECUTE_HOVER, activeforeground=EXECUTE_FG,
                      padx=18, pady=6, cursor="hand2", command=_run_or_stop)
    btn_run.pack(side="left", padx=6)
    btn_run.bind("<Enter>", lambda e: btn_run.config(bg=EXECUTE_HOVER) if btn_run["state"] == "normal" else None)
    btn_run.bind("<Leave>", lambda e: btn_run.config(bg=EXECUTE_BG) if btn_run["state"] == "normal" else None)

    # Habilitar/deshabilitar EJECUTAR en vivo, apenas cambian los datos mínimos necesarios
    excel_path_var.trace_add("write", lambda *_a: refresh_execute_state())
    col_dealer_var.trace_add("write", lambda *_a: refresh_execute_state())
    url_excel_path_var.trace_add("write", lambda *_a: refresh_execute_state())
    has_dealer_var.trace_add("write", lambda *_a: refresh_execute_state())

    # No pre-seleccionar ningún país para no confundir al usuario
    _refresh_pais_cards()
    refresh_execute_state()
    return root_frame
