# -*- coding: utf-8 -*-
"""
dealer_comparator_ui.py — Pestaña "Comparador Dealers".

Permite cargar un Excel de dealers esperados (fila de encabezado y columnas configurables,
porque el Excel que mandan varía de país a país y no siempre arranca en la fila 1), abrir
un formulario real (landing+form o solo form, igual criterio que "Envío de Leads"), navegar
región→ciudad→dealer y comparar contra el Excel filtrado, generando un reporte
PASS/FAIL/EXTRA/MISSING (Excel, y opcionalmente capturas de pantalla en ZIP).

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
    capture_result_screenshot,
    compare_dealers,
    export_results_excel,
    filter_rows,
    find_extra_dealers,
    get_country_level_defaults,
    list_model_options,
    open_target,
    read_excel_rows,
    resolve_column,
    zip_screenshots,
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
        "find_extras": False,
        "field_checks": [],
        "has_models": False,
        "models_field_id": "models",
        "models_mode": "all",
        "models_list": "",
        "url_mode": "landing_form",
        "urls_text": "",
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
    ENTRY_BG = ctx["ENTRY_BG"]
    TEXT_DELETE = ctx["TEXT_DELETE"]
    get_button_icon = ctx["get_button_icon"]
    make_scrollable_tab_container = ctx["make_scrollable_tab_container"]

    # Estilo propio para los combobox (dropdown "Configuraciones guardadas" y "Condición")
    style = ttk.Style()
    style.configure("Dealer.TCombobox", fieldbackground=ENTRY_BG, background=BTN_ACTIVE,
                     foreground="white", arrowcolor="white", bordercolor=BORDER,
                     lightcolor=BORDER, darkcolor=BORDER, padding=4)
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
        "pais": PAISES_LIST[0],
        "headers": [], "rows": [], "filtered_rows": [], "results": [],
        "driver": None, "running": False,
        "stop_event": threading.Event(),
        "field_check_widgets": [],
        "modal": None,
    }

    # ── Barra de acciones fija (solo Borrar URLs + Ejecutar), siempre visible ──
    actions_bar = Frame(tab_frame, bg=CARD_BG, bd=0, highlightthickness=1, highlightbackground=BORDER)
    actions_bar.pack(side="bottom", fill="x", pady=(6, 0))
    footer_btns = Frame(actions_bar, bg=CARD_BG)
    footer_btns.pack(fill="x", padx=15, pady=8)

    root_frame = make_scrollable_tab_container(tab_frame)

    def card(title, subtitle=None):
        block = Frame(root_frame, bg=CARD_BG, bd=0, highlightthickness=1, highlightbackground=BORDER)
        block.pack(fill="x", pady=(0, 8), ipady=5)
        head = Frame(block, bg=CARD_BG)
        head.pack(fill="x", padx=15, pady=(6, 4))
        Label(head, text=title, font=("Segoe UI", 9, "bold"), bg=CARD_BG, fg=TEXT_S).pack(side="left")
        if subtitle:
            Label(head, text=subtitle, font=("Segoe UI", 8, "italic"), bg=CARD_BG, fg="#C5A9DF",
                  wraplength=750, justify="left").pack(side="left", padx=12)
        return block

    def labeled_entry(parent_frame, label_text, var, width=16):
        col = Frame(parent_frame, bg=CARD_BG)
        col.pack(side="left", padx=6, pady=3)
        Label(col, text=label_text, font=("Segoe UI", 8), bg=CARD_BG, fg=TEXT_S).pack(anchor="w")
        Entry(col, textvariable=var, width=width, bg=ENTRY_BG, fg="white",
              insertbackground="white", relief="flat").pack()

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
            b = Button(parent, text=label_text, font=("Segoe UI", 8, "bold"), bg=BTN_INACTIVE, fg=TEXT_S,
                       relief="flat", bd=0, activebackground=BTN_HOVER, activeforeground="white",
                       highlightthickness=1, highlightbackground=BTN_INACTIVE,
                       padx=10, pady=3, cursor="hand2", command=lambda v=val: _select(v))
            b.pack(side="left", padx=2)
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

        def _toggle():
            if locked:
                return
            var.set(not var.get())
            _refresh()

        btn = Button(parent, text=label_text, font=("Segoe UI", 8, "bold"), bg=BTN_INACTIVE, fg=TEXT_S,
                     relief="flat", bd=0, activebackground=(BTN_HOVER if not locked else BTN_ACTIVE),
                     activeforeground="white", highlightthickness=1, highlightbackground=BTN_INACTIVE,
                     padx=10, pady=3, cursor=("arrow" if locked else "hand2"), command=_toggle)
        btn.pack(side="left", padx=2)
        _refresh()
        return btn

    def ui_log(msg, level="info"):
        LOGGER.info("[%s] %s", level, msg)

    # ── 1. Mercado ────────────────────────────────────────────────────────────
    mercado_card = card("🌐 MERCADO A CHEQUEAR")
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

    # ── 2. URLs a procesar (mismo bloque visual que "Generar Excels con Datos") ─
    urls_card = Frame(root_frame, bg=CARD_BG, bd=0, highlightthickness=1, highlightbackground=BORDER)
    urls_card.pack(fill="x", pady=(0, 8), ipady=6)

    urls_header = Frame(urls_card, bg=CARD_BG)
    urls_header.pack(fill="x", padx=15, pady=(6, 4))
    Label(urls_header, text="🔗 URL DEL FORMULARIO A CHEQUEAR", font=("Segoe UI", 9, "bold"),
          bg=CARD_BG, fg=TEXT_S).pack(side="left")

    mode_btn_frame = Frame(urls_header, bg=CARD_BG)
    mode_btn_frame.pack(side="right")

    url_mode_var = StringVar(value="landing_form")
    mode_btns = {}

    def switch_url_mode(mode):
        url_mode_var.set(mode)
        for m, btn in mode_btns.items():
            if m == mode:
                btn.config(bg=BTN_ACTIVE, fg="white", highlightthickness=1, highlightbackground=ACCENT)
            else:
                btn.config(bg=BTN_INACTIVE, fg=TEXT_S, highlightthickness=1, highlightbackground=BTN_INACTIVE)
        if mode == "landing_form":
            fmt_val_lbl.config(text="FORMATO: url landing  •  url form  •  url landing  •  url form  •  ... "
                                     "(una o varias, se corren todas en la misma pasada)")
        else:
            fmt_val_lbl.config(text="FORMATO: url form  •  url form  •  ... (una por línea, una o varias)")

    for m_val, m_txt in [("landing_form", "URL Landing + URL Form"), ("solo_forms", "Solo URL Form")]:
        b = Button(mode_btn_frame, text=m_txt, font=("Segoe UI", 8, "bold"), bg=BTN_INACTIVE, fg=TEXT_S,
                   relief="flat", bd=0, activebackground=BTN_HOVER, activeforeground="white",
                   highlightthickness=1, highlightbackground=BTN_INACTIVE,
                   padx=10, pady=3, cursor="hand2", command=lambda m=m_val: switch_url_mode(m))
        b.pack(side="left", padx=1)
        mode_btns[m_val] = b

    fmt_row = Frame(urls_card, bg=CARD_BG)
    fmt_row.pack(fill="x", padx=15, pady=(4, 4))
    fmt_val_lbl = Label(fmt_row, text="FORMATO: url landing  •  url form  •  url landing  •  url form  •  ...",
                         font=("Segoe UI", 8, "bold"), bg=CARD_BG, fg="#C5A9DF", wraplength=750, justify="left")
    fmt_val_lbl.pack(side="left")

    url_text_border = Frame(urls_card, bg=BORDER, padx=1, pady=1)
    url_text_border.pack(fill="x", padx=15, pady=4)
    v_scroll_url = ttk.Scrollbar(url_text_border, orient="vertical")
    url_text_area = Text(url_text_border, bg=ENTRY_BG, fg="white", insertbackground="white",
                          bd=0, relief="flat", height=5, font=("Consolas", 9),
                          yscrollcommand=v_scroll_url.set)
    v_scroll_url.config(command=url_text_area.yview)
    v_scroll_url.pack(side="right", fill="y")
    url_text_area.pack(fill="both", expand=True, padx=(3, 0), pady=3)

    url_warn_var = StringVar(value="")
    url_warn_lbl = Label(urls_card, textvariable=url_warn_var, font=("Segoe UI", 8, "italic"),
                          bg=CARD_BG, fg="#F8C471", wraplength=750, justify="left")
    url_warn_lbl.pack(anchor="w", padx=15, pady=(0, 2))

    def _check_url_country_mismatch(*_a):
        detected = _detect_url_country(url_text_area.get("1.0", "end-1c"))
        if detected and detected != state["pais"]:
            url_warn_var.set(
                f"⚠ Las URLs parecen de {detected} pero tenés seleccionado {state['pais']}. "
                "Verificá el mercado antes de ejecutar."
            )
        else:
            url_warn_var.set("")

    url_text_area.bind("<KeyRelease>", _check_url_country_mismatch)

    device_row = Frame(urls_card, bg=CARD_BG)
    device_row.pack(fill="x", padx=15, pady=(4, 4))
    Label(device_row, text="Navegador:", font=("Segoe UI", 8), bg=CARD_BG, fg=TEXT_S).pack(side="left")
    browser_var = StringVar(value="chrome")
    make_single_select(device_row, BROWSER_OPTIONS, browser_var, default="chrome")

    ver_navegador_var = BooleanVar(value=False)
    Checkbutton(urls_card, text="Ver navegador mientras corre (si está apagado, corre atrás sin molestar)",
                variable=ver_navegador_var, bg=CARD_BG, fg=TEXT_S, selectcolor=ENTRY_BG,
                activebackground=CARD_BG).pack(anchor="w", padx=15, pady=(0, 6))

    # Vista siempre escritorio (sin selector: el comparador de dealers no necesita mobile)
    viewport_var = StringVar(value="fullscreen")

    switch_url_mode("landing_form")

    def _parse_urls_text():
        """Devuelve una lista de pares (landing_url, form_url) — soporta uno o varios forms
        en la misma pasada. En modo 'solo_forms' cada línea es un form suelto (landing="")."""
        lines = [ln.strip() for ln in url_text_area.get("1.0", "end-1c").split("\n") if ln.strip()]
        if url_mode_var.get() == "solo_forms":
            return [("", ln) for ln in lines]
        pairs = []
        for i in range(0, len(lines) - 1, 2):
            pairs.append((lines[i], lines[i + 1]))
        return pairs

    # ── 3. Excel de dealers ──────────────────────────────────────────────────
    excel_card = card("📄 EXCEL DE DEALERS A CHEQUEAR",
                       "Configurá fila de encabezado y columna de filtro: el Excel varía de país a país.")
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
          fg=TEXT_S, wraplength=650, justify="left").pack(side="left", padx=10)

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

    CONDITION_LABELS = ["Incluir", "Excluir", "Buscar extras"]
    condition_mode_var = StringVar(value="Incluir")
    cond_col = Frame(filter_row, bg=CARD_BG)
    cond_col.pack(side="left", padx=6, pady=3)
    Label(cond_col, text="Condición", font=("Segoe UI", 8), bg=CARD_BG, fg=TEXT_S).pack(anchor="w")
    ttk.Combobox(cond_col, textvariable=condition_mode_var, state="readonly", width=22,
                 style="Dealer.TCombobox", values=CONDITION_LABELS).pack()

    # ── 4. Columnas del Excel ────────────────────────────────────────────────
    columns_card = card("🧭 COLUMNAS DEL EXCEL")
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
    make_toggle_pill(levels_row, "dealer", has_dealer_var, locked=True)

    Label(columns_card,
          text="ℹ Estos son los id tal cual aparecen en el HTML del <select> del form (region / city / "
               "dealer) — no cambian de país a país aunque el texto que ve el usuario sí (ej. en Argentina "
               "se ve como \"Provincia\" pero el id sigue siendo \"region\"). Desactivá acá el que tu form "
               "no tenga; dealer siempre es obligatorio.",
          font=("Segoe UI", 8, "italic"), bg=CARD_BG, fg="#C5A9DF", wraplength=750,
          justify="left").pack(anchor="w", padx=15, pady=(2, 6))

    cols_row = Frame(columns_card, bg=CARD_BG)
    cols_row.pack(fill="x", padx=9, pady=4)
    col_region_var = StringVar(value="REGION")
    col_city_var = StringVar(value="CIUDAD")
    col_dealer_var = StringVar(value="NOMBRE")
    col_bac_var = StringVar(value="BAC")
    labeled_entry(cols_row, "Columna Región", col_region_var)
    labeled_entry(cols_row, "Columna Ciudad", col_city_var)
    labeled_entry(cols_row, "Columna Dealer", col_dealer_var)
    labeled_entry(cols_row, "Columna BAC", col_bac_var)

    opts_row = Frame(columns_card, bg=CARD_BG)
    opts_row.pack(fill="x", padx=15, pady=(0, 4))
    chk_bac_var = BooleanVar(value=False)
    Checkbutton(opts_row, text="Verificar BAC (opcional — muchos forms no exponen data-bac en el HTML)",
                variable=chk_bac_var, bg=CARD_BG, fg=TEXT_S, selectcolor=ENTRY_BG,
                activebackground=CARD_BG).pack(anchor="w")
    find_extras_var = BooleanVar(value=False)
    Checkbutton(opts_row,
                text="También buscar dealers EXTRA (en el form pero no en el Excel) y DUPLICADOS "
                     "(repetidos en el <select> del form) — mismo efecto que elegir \"Buscar extras\" "
                     "en Condición",
                variable=find_extras_var, bg=CARD_BG, fg=TEXT_S, selectcolor=ENTRY_BG,
                activebackground=CARD_BG).pack(anchor="w")

    field_check_block = Frame(columns_card, bg=CARD_BG)
    field_check_block.pack(fill="x", padx=15, pady=(4, 6))
    Label(field_check_block, text="Columnas adicionales a comprobar en el form",
          font=("Segoe UI", 8, "bold"), bg=CARD_BG, fg=TEXT_S).pack(anchor="w")
    Label(field_check_block,
          text="Agregá cualquier otro campo del form que quieras validar: el nombre de la columna del "
               "Excel y el id del campo tal cual aparece en el HTML del form (ej. columna \"CEP\" → id "
               "\"customer-cep\").",
          font=("Segoe UI", 8), bg=CARD_BG, fg=TEXT_S, wraplength=750, justify="left").pack(anchor="w", pady=(0, 4))

    field_check_header = Frame(field_check_block, bg=CARD_BG)
    field_check_header.pack(fill="x")
    Label(field_check_header, text="Columna del Excel", font=("Segoe UI", 8, "bold"), bg=CARD_BG,
          fg="#C5A9DF", width=22, anchor="w").pack(side="left", padx=2)
    Label(field_check_header, text="ID del campo en el form", font=("Segoe UI", 8, "bold"), bg=CARD_BG,
          fg="#C5A9DF", width=22, anchor="w").pack(side="left", padx=2)

    field_check_rows_container = Frame(field_check_block, bg=CARD_BG)
    field_check_rows_container.pack(fill="x", pady=(2, 4))

    def _add_field_check_row(column="", field_id=""):
        row = Frame(field_check_rows_container, bg=CARD_BG)
        row.pack(fill="x", pady=2)
        column_var = StringVar(value=column)
        field_id_var = StringVar(value=field_id)
        Entry(row, textvariable=column_var, width=22, bg=ENTRY_BG, fg="white",
              insertbackground="white", relief="flat").pack(side="left", padx=2)
        Entry(row, textvariable=field_id_var, width=22, bg=ENTRY_BG, fg="white",
              insertbackground="white", relief="flat").pack(side="left", padx=2)

        widgets = {"frame": row, "column_var": column_var, "field_id_var": field_id_var}

        def _remove():
            row.destroy()
            state["field_check_widgets"].remove(widgets)

        Button(row, text="×", font=("Segoe UI", 9, "bold"), bg="#3d1414", fg="#f85149",
               relief="flat", bd=0, width=2, cursor="hand2", command=_remove).pack(side="left", padx=4)
        state["field_check_widgets"].append(widgets)

    Button(field_check_block, text="+ Agregar columna a comprobar",
           font=("Segoe UI", 8, "bold"), bg=BTN_ACTIVE, fg="white", relief="flat", bd=0,
           cursor="hand2", command=lambda: _add_field_check_row()).pack(anchor="w", pady=2)

    # ── 4b. Modelos ───────────────────────────────────────────────────────────
    models_card = card("🚗 MODELOS",
                        "Si el form tiene selector de modelo (id \"models\"), podés correr la comparación "
                        "para uno o varios modelos puntuales, o para todos los que tenga el form.")
    has_models_var = BooleanVar(value=False)
    models_toggle_row = Frame(models_card, bg=CARD_BG)
    models_toggle_row.pack(fill="x", padx=15, pady=(0, 4))
    make_toggle_pill(models_toggle_row, "El form tiene selector de Modelo", has_models_var)
    Label(models_card, text="⚠ Este apartado solo funciona con forms T1 (el id \"models\" estándar). "
                             "Forms T2/T3 con selector de modelo distinto no están soportados todavía.",
          font=("Segoe UI", 8, "italic"), bg=CARD_BG, fg="#F8C471", wraplength=750,
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
    Entry(models_list_row, textvariable=models_list_var, bg=ENTRY_BG, fg="white",
          insertbackground="white", relief="flat", width=60).pack(anchor="w", pady=(2, 0))
    Label(models_list_row, text="ej: Onix, Tracker, S10", font=("Segoe UI", 8, "italic"), bg=CARD_BG,
          fg="#C5A9DF").pack(anchor="w")
    _refresh_models_mode_visibility()

    # ── 5. Modo de salida ─────────────────────────────────────────────────────
    output_card = card("📦 MODO DE SALIDA")
    output_mode_var = StringVar(value="excel")
    output_row = Frame(output_card, bg=CARD_BG)
    output_row.pack(fill="x", padx=15, pady=(0, 6))
    for val, txt in (("excel", "Solo Excel"), ("caps", "Excel + Capturas (ZIP)")):
        ttk.Radiobutton(output_row, text=txt, value=val, variable=output_mode_var).pack(side="left", padx=8)

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
                                 style="Dealer.TCombobox")
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
        ui_log(f"Configuración '{name}' cargada.", "ok")

    def _delete_preset():
        name = preset_selected_var.get().strip()
        if name and name in all_settings["presets"]:
            del all_settings["presets"][name]
            _save_all_settings(all_settings)
            _refresh_preset_combo()
            preset_selected_var.set("")
            ui_log(f"Configuración '{name}' eliminada.", "warn")

    Button(presets_row2, text="📂 Cargar", font=("Segoe UI", 8, "bold"),
           bg=BTN_ACTIVE, fg="white", relief="flat", bd=0, cursor="hand2",
           command=_load_preset).pack(side="left", padx=4)
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
        condition_mode_var.set(_LEGACY_CONDITION_MAP.get(raw_condition, raw_condition))
        has_region_var.set(bool(cfg.get("has_region", True)))
        has_city_var.set(bool(cfg.get("has_city", True)))
        col_region_var.set(cfg.get("col_region", "REGION"))
        col_city_var.set(cfg.get("col_city", "CIUDAD"))
        col_dealer_var.set(cfg.get("col_dealer", "NOMBRE"))
        col_bac_var.set(cfg.get("col_bac", "BAC"))
        chk_bac_var.set(bool(cfg.get("chk_bac", False)))
        find_extras_var.set(bool(cfg.get("find_extras", False)))
        has_models_var.set(bool(cfg.get("has_models", False)))
        models_field_id_var.set(cfg.get("models_field_id", "models"))
        select_models_mode(cfg.get("models_mode", "all"))
        models_list_var.set(cfg.get("models_list", ""))
        url_mode_var.set(cfg.get("url_mode", "landing_form"))
        switch_url_mode(url_mode_var.get())
        url_text_area.delete("1.0", "end")
        url_text_area.insert("1.0", cfg.get("urls_text", ""))
        _check_url_country_mismatch()
        browser_var.set(cfg.get("browser", "chrome"))
        viewport_var.set(cfg.get("viewport", "fullscreen"))
        ver_navegador_var.set(bool(cfg.get("ver_navegador", False)))
        output_mode_var.set(cfg.get("output_mode", "excel"))

        for row_widgets in list(state["field_check_widgets"]):
            row_widgets["frame"].destroy()
        state["field_check_widgets"] = []
        for check in cfg.get("field_checks", []):
            _add_field_check_row(check.get("column", ""), check.get("field_id", ""))

    def _apply_country_settings(pais):
        cfg = all_settings["paises"].get(pais) or _default_country_settings(pais)
        _apply_settings_dict(cfg)

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
            "col_region": col_region_var.get(),
            "col_city": col_city_var.get(),
            "col_dealer": col_dealer_var.get(),
            "col_bac": col_bac_var.get(),
            "chk_bac": chk_bac_var.get(),
            "find_extras": find_extras_var.get(),
            "has_models": has_models_var.get(),
            "models_field_id": models_field_id_var.get(),
            "models_mode": models_mode_var.get(),
            "models_list": models_list_var.get(),
            "url_mode": url_mode_var.get(),
            "urls_text": url_text_area.get("1.0", "end-1c"),
            "browser": browser_var.get(),
            "viewport": viewport_var.get(),
            "ver_navegador": ver_navegador_var.get(),
            "output_mode": output_mode_var.get(),
            "field_checks": [
                {"column": w["column_var"].get().strip(), "field_id": w["field_id_var"].get().strip()}
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
            "region": resolve_column(headers, col_region_var.get()),
            "city": resolve_column(headers, col_city_var.get()),
            "dealer": resolve_column(headers, col_dealer_var.get()),
            "bac": resolve_column(headers, col_bac_var.get()),
        }

    def _field_checks_resolved():
        headers = state["headers"]
        out = []
        for w in state["field_check_widgets"]:
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
            mode = "exclude" if condition_mode_var.get() == "Excluir" else "include"
            filtered = filter_rows(rows, filter_col_key, filter_value_var.get(), mode)
        state["filtered_rows"] = filtered
        return headers, filtered

    def _should_find_extras():
        return find_extras_var.get() or condition_mode_var.get() == "Buscar extras"

    # ── Modal de ejecución (mismo lenguaje visual que "Envío de Leads") ──────
    def _open_run_modal():
        modal = Toplevel(root)
        modal.overrideredirect(True)
        modal_width, modal_height = 460, 280
        px = root.winfo_rootx() + (root.winfo_width() - modal_width) // 2
        py = root.winfo_rooty() + (root.winfo_height() - modal_height) // 2
        modal.geometry(f"{modal_width}x{modal_height}+{px}+{py}")
        modal.configure(bg=MODAL_BG, bd=1, highlightthickness=1, highlightbackground=BORDER)
        modal.transient(root)
        modal.lift()
        modal.focus_set()
        modal.grab_set()  # bloquea la interacción con el resto de la app mientras corre

        title_bar = Frame(modal, bg=MODAL_BG)
        title_bar.pack(fill="x", padx=15, pady=(5, 0))
        title_bar_lbl = Label(title_bar, text="Comparador Dealers", font=("Segoe UI", 8, "bold"),
                               bg=MODAL_BG, fg="#C5A9DF")
        title_bar_lbl.pack(side="left")

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

        def _on_close_click():
            if "Comparando" in title_lbl.cget("text"):
                if messagebox.askyesno("Detener ejecución", "¿Querés detener la comparación y cerrar la ventana?"):
                    state["stop_event"].set()
                    modal.destroy()
            else:
                modal.destroy()

        btn_cls = Button(title_bar, text="✕", font=("Segoe UI", 8, "bold"), bg=MODAL_BG, fg="#C5A9DF",
                          relief="flat", bd=0, cursor="hand2", padx=6, pady=2, command=_on_close_click)
        btn_cls.pack(side="right")

        header = Frame(modal, bg=MODAL_BG)
        header.pack(fill="x", padx=20, pady=(8, 10))
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

        rotation_glyphs = ["↻", "➔", "↻", "➔"]

        def _rotate_icon(idx=0):
            if modal.winfo_exists() and "Comparando" in title_lbl.cget("text"):
                icon_lbl.config(text=rotation_glyphs[idx % len(rotation_glyphs)])
                modal.after(250, lambda: _rotate_icon(idx + 1))
        _rotate_icon()

        def _on_detener():
            state["stop_event"].set()
            btn_detener.config(state="disabled", text=" Deteniendo...")

        btn_detener = Button(header, text=" Detener", image=get_button_icon("stop_coral.png"), compound="left",
                              font=("Segoe UI", 9, "bold"), bg="#3D1220", fg="#F1948A", relief="flat", bd=0,
                              highlightthickness=1, highlightbackground="#F1948A", cursor="hand2",
                              command=_on_detener, padx=12, pady=4)
        btn_detener.pack(side="right")

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

        btn_close = Button(modal, text="Cerrar", font=("Segoe UI", 10, "bold"), bg="#AED6F1", fg="#110518",
                            relief="flat", bd=0, cursor="hand2", pady=6, command=modal.destroy)

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
                    summary_lbl.config(
                        text=f"🟢 {counts.get('PASS', 0)} PASS   🔴 {counts.get('FAIL', 0)} FAIL   "
                             f"🟡 {counts.get('EXTRA', 0)} EXTRA"
                    )
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
            btn_run.config(text=" DETENER", image=get_button_icon("stop_coral.png"), bg="#b91c1c", fg="white")
        else:
            btn_run.config(text=" EJECUTAR", image=get_button_icon("download_blue.png"),
                            bg=VALIDATE_BG, fg=VALIDATE_FG)

    def _validate_and_prepare():
        """Corre en el hilo principal, ANTES de abrir el modal de ejecución. Si algo está mal
        (Excel no seleccionado, columna faltante, URL vacía) tira ValueError con el motivo y
        nunca se llega a abrir el modal — así un error de tipeo no deja la pestaña bloqueada."""
        headers, filtered_rows = _load_and_filter_excel()
        ui_log(f"Excel cargado: {len(state['rows'])} filas, {len(filtered_rows)} tras filtro.", "info")
        if not filtered_rows:
            raise ValueError("No hay filas para comparar después del filtro.")

        column_map = _column_map()
        if not column_map["dealer"]:
            raise ValueError(f"No se encontró la columna de dealer '{col_dealer_var.get()}' en el Excel.")

        url_pairs = [(landing, form_url) for landing, form_url in _parse_urls_text() if form_url]
        if not url_pairs:
            raise ValueError("Ingresá al menos una URL de form en el bloque de URLs.")

        return filtered_rows, column_map, url_pairs

    def _worker_run(filtered_rows, column_map, url_pairs):
        driver = None
        counts = {"PASS": 0, "FAIL": 0, "EXTRA": 0, "DUPLICADO": 0}
        error_msg = None
        try:
            driver = BrowserManager.create_browser(
                browser_type=browser_var.get(), viewport=viewport_var.get(),
                headless=False, background=not ver_navegador_var.get(),
            )
            state["driver"] = driver

            all_results = []
            screenshots = []
            output_mode = output_mode_var.get()
            total_pairs = len(url_pairs)

            for pair_idx, (landing_url, form_url) in enumerate(url_pairs, start=1):
                if state["stop_event"].is_set():
                    raise StopRequested()
                ui_log(f"\n=== Form {pair_idx}/{total_pairs}: {form_url} ===", "info")
                open_target(driver, url_mode_var.get(), landing_url, form_url)

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

                def _shot(result, _form_url=form_url, _landing_url=landing_url):
                    if output_mode != "caps":
                        return
                    safe_name = "".join(
                        c for c in (result.get("dealer") or "dealer") if c.isalnum() or c in " _-"
                    )[:40]
                    filename = f"{result['status']}_{pair_idx}_{result.get('fila')}_{safe_name}.png".replace(" ", "_")
                    path = capture_result_screenshot(driver, _get_results_dir(), filename, _form_url or _landing_url)
                    screenshots.append(path)

                def _pair_progress_cb(cur, total, label_txt, _pair_idx=pair_idx):
                    _progress_cb(cur, total, f"[Form {_pair_idx}/{total_pairs}] {label_txt}")

                results = compare_dealers(
                    driver, filtered_rows, column_map,
                    level_ids=DEFAULT_SELECT_IDS,
                    has_region=has_region_var.get(), has_city=has_city_var.get(),
                    chk_bac=chk_bac_var.get(), field_checks=_field_checks_resolved(),
                    model_field_id=model_field_id, models=models_to_run,
                    log_cb=ui_log, progress_cb=_pair_progress_cb, stop_flag=state["stop_event"],
                    screenshot_cb=_shot,
                )
                for r in results:
                    r["url_form"] = form_url

                if _should_find_extras():
                    ui_log("Buscando dealers EXTRA en el form...", "info")
                    extras = find_extra_dealers(
                        driver, filtered_rows, column_map, level_ids=DEFAULT_SELECT_IDS,
                        has_region=has_region_var.get(), has_city=has_city_var.get(),
                        log_cb=ui_log, progress_cb=_pair_progress_cb, stop_flag=state["stop_event"],
                    )
                    for r in extras:
                        r["url_form"] = form_url
                    results = results + extras

                all_results.extend(results)
                state["results"] = all_results

            for r in all_results:
                counts[r["status"]] = counts.get(r["status"], 0) + 1
            ui_log(
                f"Comparación terminada: PASS={counts.get('PASS', 0)} FAIL={counts.get('FAIL', 0)} "
                f"EXTRA={counts.get('EXTRA', 0)} DUPLICADO={counts.get('DUPLICADO', 0)}", "ok"
            )

            results_dir = _get_results_dir()
            os.makedirs(results_dir, exist_ok=True)
            export_path = os.path.join(
                results_dir, f"dealer_comparator_{state['pais']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            )
            export_path = export_results_excel(all_results, output_path=export_path, pais=state["pais"])
            ui_log(f"Reporte Excel: {export_path}", "ok")

            if output_mode == "caps" and screenshots:
                zip_path = os.path.join(
                    _get_results_dir(),
                    f"capturas_{state['pais']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                )
                zip_screenshots(screenshots, zip_path)
                # Los PNG sueltos ya quedaron adentro del ZIP: no hace falta dejarlos también sueltos.
                for shot_path in screenshots:
                    try:
                        os.remove(shot_path)
                    except Exception:
                        pass
                ui_log(f"Capturas empaquetadas: {zip_path}", "ok")

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

    def _borrar_urls():
        url_text_area.delete("1.0", "end")
        ui_log("URLs borradas.", "info")

    btn_borrar = Button(footer_btns, text=" Borrar URLs", image=get_button_icon("trash_coral.png"), compound="left",
                         font=("Segoe UI", 9, "bold"), bg=BTN_INACTIVE, fg=TEXT_DELETE,
                         relief="flat", bd=0, activebackground=BTN_HOVER, activeforeground=TEXT_DELETE,
                         padx=18, pady=6, cursor="hand2", command=_borrar_urls)
    btn_borrar.pack(side="left", padx=(0, 6))
    btn_borrar.bind("<Enter>", lambda e: btn_borrar.config(bg=BTN_HOVER))
    btn_borrar.bind("<Leave>", lambda e: btn_borrar.config(bg=BTN_INACTIVE))

    btn_run = Button(footer_btns, text=" EJECUTAR", image=get_button_icon("download_blue.png"), compound="left",
                      font=("Segoe UI", 9, "bold"), bg=VALIDATE_BG, fg=VALIDATE_FG,
                      relief="flat", bd=0, activebackground=VALIDATE_HOVER, activeforeground=VALIDATE_FG,
                      padx=18, pady=6, cursor="hand2", command=_run_or_stop)
    btn_run.pack(side="left", padx=6)

    select_country(state["pais"])
    return root_frame
