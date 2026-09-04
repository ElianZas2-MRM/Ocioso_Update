"""
field_validation_ui.py — Pestaña de validación de campos en la interfaz gráfica.
Permite testear las reglas de validación de cada campo del formulario (regex, largo, etc.)
lanzando un browser Selenium contra la URL real del formulario.
"""
import json
import logging
import os
import re
import sys
import threading
from copy import deepcopy
import pandas as pd
from tkinter import BOTH, END, LEFT, RIGHT, W, BooleanVar, Button, Checkbutton, Entry, Frame, Label, StringVar, Text, Toplevel
from tkinter import messagebox
from tkinter import ttk

from core.browser_manager import BrowserManager
from interface.helpers_interface import DATA_DIR, cargar_config_global, guardar_config_global, obtener_email_destinatario
from interface.ids_dinamicos_ui import registrar_campos_importados_desde_excel
from utils.crm_excel_importer import CrmValidacionesImporter, construir_reporte, guardar_reporte
from validation.selenium_validation_runner import run_field_validations
from validation.validation_email import send_validation_report_email
from validation.validation_exporter import export_validation_results


LOGGER = logging.getLogger(__name__)

AVAILABLE_COUNTRIES = [
    "Argentina",
    "Bolivia",
    "Brasil",
    "Chile",
    "Colombia",
    "Ecuador",
    "Paraguay",
    "Peru",
    "Uruguay",
]

BROWSER_OPTIONS = ["chrome", "firefox", "edge"]
VIEWPORT_OPTIONS = ["fullscreen", "600x738"]

# Chile puede ser CL o CH
COUNTRY_ABBREVIATIONS = {
    "Argentina": "AR",
    "Bolivia": "BO",
    "Brasil": "BR",
    "Chile": "CH",
    "Colombia": "CO",
    "Ecuador": "EC",
    "Paraguay": "PY",
    "Peru": "PE",
    "Uruguay": "UY",
}
# Abreviaciones alternativas
ALTERNATIVE_ABBREVIATIONS = {
    "CL": "Chile",
    "CH": "Chile",
}
ABBREVIATION_TO_COUNTRY = {abbrev: country for country, abbrev in COUNTRY_ABBREVIATIONS.items()}
ABBREVIATION_TO_COUNTRY.update(ALTERNATIVE_ABBREVIATIONS)
# Create a mapping for lowercase country abbreviations for validation
ABREV_PAISES_LOWER = {abbrev.lower(): abbrev for abbrev in COUNTRY_ABBREVIATIONS.values()}
for alt in ALTERNATIVE_ABBREVIATIONS:
    ABREV_PAISES_LOWER[alt.lower()] = alt
ABREV_TO_COUNTRY_LOWER = {abbrev.lower(): country for country, abbrev in COUNTRY_ABBREVIATIONS.items()}
for alt, country in ALTERNATIVE_ABBREVIATIONS.items():
    ABREV_TO_COUNTRY_LOWER[alt.lower()] = country
COUNTRY_NAMES_LOWER = {country.lower(): country for country in AVAILABLE_COUNTRIES}


def _get_base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _get_rules_path(pais=None):
    """Retorna la ruta del JSON de validación. Si pais está definido, SIEMPRE apunta al
    archivo per-país (exista o no todavía) — nunca cae al global, porque _save_rules_file
    lo crearía ahí y mezclaría campos de un país en el archivo compartido por los demás."""
    json_dir = os.path.join(_get_base_dir(), "json")
    if pais:
        pais_key = pais.lower().replace("é", "e").replace("ú", "u").replace("ó", "o")
        return os.path.join(json_dir, f"field_validation_rules_{pais_key}.json")
    return os.path.join(json_dir, "field_validation_rules.json")


def _get_asset_dir():
    if getattr(sys, "frozen", False):
        bundle_dir = getattr(sys, "_MEIPASS", os.path.join(_get_base_dir(), "_internal"))
        return os.path.join(bundle_dir, "Asset")
    return os.path.join(_get_base_dir(), "Asset")


def _get_validation_urls_excel_path():
    return os.path.join(DATA_DIR, "Field_Validation_URLs.xlsx")


def _get_results_dir():
    return os.path.join(_get_base_dir(), "resultados")


def _load_rules_file(pais=None):
    rules_path = _get_rules_path(pais)
    if not os.path.exists(rules_path):
        return {
            "url": "",
            "urls": [],
            "landing_urls": [],
            "form_urls": [],
            "browser": "chrome",
            "viewport": "fullscreen",
            "selected_browsers": ["chrome"],
            "selected_viewports": ["fullscreen"],
            "selected_paises": list(AVAILABLE_COUNTRIES),
            "fields": {},
        }

    with open(rules_path, "r", encoding="utf-8") as file_handle:
        return json.load(file_handle)


def _save_rules_file(payload, pais=None):
    rules_path = _get_rules_path(pais)
    os.makedirs(os.path.dirname(rules_path), exist_ok=True)
    with open(rules_path, "w", encoding="utf-8") as file_handle:
        json.dump(payload, file_handle, indent=2, ensure_ascii=False)


def _apply_row_colors(tree_widget):
    # Intentar obtener el color de texto (foreground) configurado en el estilo
    style = ttk.Style()
    fg = style.lookup("Section.Treeview", "foreground")
    if not fg:
        fg = "white"  # Valor por defecto si no se encuentra
    
    # Comprobar si el texto es claro u oscuro para asignar los fondos apropiados
    def _is_light_color(color_str):
        if not color_str:
            return False
        color_str = color_str.strip().lower()
        if color_str in ("white", "#ffffff", "#fff", "yellow", "cyan", "lavender"):
            return True
        if color_str in ("black", "none", ""):
            return False
        if color_str.startswith("#"):
            try:
                hex_val = color_str.lstrip("#")
                if len(hex_val) == 3:
                    hex_val = "".join([c*2 for c in hex_val])
                r = int(hex_val[0:2], 16)
                g = int(hex_val[2:4], 16)
                b = int(hex_val[4:6], 16)
                brightness = (r * 299 + g * 587 + b * 114) / 1000
                return brightness > 130
            except Exception:
                return False
        return False

    if _is_light_color(fg):
        # Texto claro -> Fondos oscuros contrastantes (Premium Figma-style)
        tree_widget.tag_configure("odd", background="#2E1146", foreground="white")
        tree_widget.tag_configure("even", background="#4A2666", foreground="#E6D6F2")
    else:
        # Texto oscuro -> Fondos claros contrastantes
        tree_widget.tag_configure("odd", background="white", foreground="black")
        tree_widget.tag_configure("even", background="#f1edf6", foreground="#333333")

    for index, item in enumerate(tree_widget.get_children()):
        tree_widget.item(item, tags=(("odd" if index % 2 == 0 else "even"),))


def build_field_validation_tab(parent, palette, shared_config=None):
    section_bg = palette["section_bg"]
    container_bg = palette["container_bg"]
    app_bg = palette["app_bg"]
    text_color = palette["text_color"]
    button_bg = palette.get("button_bg", "#9c6fb4")
    button_fg = palette.get("button_fg", "black")
    popup_bg = app_bg
    popup_text = "white"
    popup_button_bg = button_bg
    popup_button_fg = button_fg

    # Colores de campos/tablas: por defecto blanco/negro (run original). El demo puede
    # pasar valores oscuros para que combinen con su tema.
    entry_bg = palette.get("entry_bg", "white")
    entry_fg = palette.get("entry_fg", "black")
    tree_bg = palette.get("tree_bg", container_bg)
    tree_fg = palette.get("tree_fg", text_color)
    heading_bg = palette.get("heading_bg", button_bg)

    _vstyle = ttk.Style()
    _vstyle.configure("TEntry", fieldbackground=entry_bg, background=entry_bg,
                      foreground=entry_fg, insertcolor=entry_fg,
                      bordercolor=button_bg, lightcolor=button_bg, darkcolor=button_bg)
    _vstyle.map("TEntry",
                fieldbackground=[("readonly", entry_bg), ("focus", entry_bg),
                                 ("active", entry_bg), ("disabled", "#656565"), ("!disabled", entry_bg)],
                foreground=[("disabled", "#aaaaaa"), ("!disabled", entry_fg)])
    _vstyle.configure("Section.Treeview", background=tree_bg, foreground=tree_fg,
                      fieldbackground=tree_bg, borderwidth=0)
    _vstyle.configure("Section.Treeview.Heading", background=heading_bg, foreground=text_color,
                      font=("Segoe UI", 9, "bold"), borderwidth=0)

    state = {
        "rows": [],
        "fields": [],
        "error_rows": [],
        "unmapped_ids": [],
        "summary": {},
        "export_path": None,
    }

    def _format_user_error(raw_error, fallback="Ocurrió un error inesperado.", max_len=220):
        """Devuelve un mensaje corto y amigable para mostrar en UI."""
        if raw_error is None:
            return fallback

        text = str(raw_error).strip()
        if not text:
            return fallback

        split_tokens = [
            "(Session info:",
            "Stacktrace:",
            "Call Stack:",
            "Build info:",
            "System info:",
            "Driver info:",
            "\nStack trace:",
            "\n\tat ",
        ]
        for token in split_tokens:
            if token in text:
                text = text.split(token)[0].strip()

        text = re.sub(r"\s+", " ", text).strip()
        if len(text) > max_len:
            text = text[: max_len - 1].rstrip() + "..."

        return text or fallback

    def _build_execution_warning(execution_errors):
        """Compacta errores por ejecución para evitar popups gigantes."""
        if not execution_errors:
            return ""

        max_items = 5
        lines = ["La validación finalizó con fallas en algunas ejecuciones.", ""]
        for item in execution_errors[:max_items]:
            lines.append(f"- {_format_user_error(item, fallback='Error no especificado', max_len=180)}")

        remaining = len(execution_errors) - max_items
        if remaining > 0:
            lines.append("")
            lines.append(f"... y {remaining} error(es) más. Revisá el Excel de resultados para el detalle.")

        return "\n".join(lines)

    browser_vars = (
        shared_config.get("browser_vars")
        if shared_config and shared_config.get("browser_vars")
        else {browser: BooleanVar(value=(browser == "chrome")) for browser in BROWSER_OPTIONS}
    )
    viewport_vars = (
        shared_config.get("viewport_vars")
        if shared_config and shared_config.get("viewport_vars")
        else {viewport: BooleanVar(value=(viewport == "fullscreen")) for viewport in VIEWPORT_OPTIONS}
    )
    teclado_mobile_var = BooleanVar(value=False)
    ver_navegador_var = BooleanVar(value=True)
    rule_filter_text_var = StringVar(value="")
    rule_filter_country_var = StringVar(value="Todos")
    upsert_button_text_var = StringVar(value="Agregar regla")
    field_name_var = StringVar()
    element_id_var = StringVar()
    dropdown_var = BooleanVar(value=False)
    regex_full_var = StringVar()
    regex_char_var = StringVar()
    status_var = StringVar(value="Sin ejecuciones.")
    rule_country_vars = {country: BooleanVar(value=False) for country in AVAILABLE_COUNTRIES}
    email_var = (
        shared_config.get("email_var")
        if shared_config and shared_config.get("email_var")
        # obtener_email_destinatario() devuelve una LISTA; pasarla directo a StringVar
        # la serializa como lista Tcl (pierde las comas, aparecen llaves "{...}").
        else StringVar(value=", ".join(obtener_email_destinatario()))
    )

    cfg_global = cargar_config_global()
    if "enviar_mail" not in cfg_global:
        cfg_global["enviar_mail"] = False
    if "adjuntar_resultados" not in cfg_global:
        cfg_global["adjuntar_resultados"] = True
    guardar_config_global(cfg_global)

    enviar_mail_var = (
        shared_config.get("enviar_mail_var")
        if shared_config and shared_config.get("enviar_mail_var")
        else BooleanVar(value=bool(cfg_global.get("enviar_mail", False)))
    )
    adjuntar_resultados_var = (
        shared_config.get("adjuntar_resultados_var")
        if shared_config and shared_config.get("adjuntar_resultados_var")
        else BooleanVar(value=bool(cfg_global.get("adjuntar_resultados", True)))
    )

    root_frame = Frame(parent, bg=app_bg, highlightthickness=0, bd=0)
    root_frame.pack(fill=BOTH, expand=True, padx=20, pady=(20, 15))  # Más espacio arriba

    top_frame = Frame(root_frame, bg=app_bg, highlightthickness=0, bd=0)
    top_frame.pack(fill="x", pady=(0, 10))

    config_frame = Frame(top_frame, bg=app_bg)
    config_frame.pack(fill="x", padx=12, pady=(8, 6))

    # --- Preview Excel (debajo de los CTAs) ---
    # Ahora los CTAs van debajo del preview, así que solo declaramos el frame aquí
    urls_actions_frame = Frame(config_frame, bg=app_bg)
    # No se hace pack aquí, se hará después del preview
    actions_frame = urls_actions_frame

    # --- Preview Excel ---
    urls_excel_frame = Frame(config_frame, bg=app_bg)
    urls_excel_frame.pack(fill="both", expand=True, pady=(12, 0))  # Espacio extra arriba del preview

    urls_tree_frame = Frame(urls_excel_frame, bg=app_bg)
    urls_tree_frame.pack(fill="both", expand=True)
    urls_excel_tree = ttk.Treeview(
        urls_tree_frame,
        columns=("pais", "landing_url", "form_url"),
        show="headings",
        height=5,
        style="Section.Treeview",
    )
    urls_excel_tree.heading("pais", text="País")
    urls_excel_tree.heading("landing_url", text="URL")
    urls_excel_tree.heading("form_url", text="Formulario")
    urls_excel_tree.column("pais", width=80, stretch=False)
    urls_excel_tree.column("landing_url", width=380, stretch=True)
    urls_excel_tree.column("form_url", width=380, stretch=True)
    urls_v_scroll = ttk.Scrollbar(urls_tree_frame, orient="vertical", command=urls_excel_tree.yview)
    urls_h_scroll = ttk.Scrollbar(urls_tree_frame, orient="horizontal", command=urls_excel_tree.xview)

    urls_tree_frame.grid_columnconfigure(0, weight=1)
    urls_tree_frame.grid_rowconfigure(0, weight=1)
    urls_excel_tree.grid(row=0, column=0, sticky="nsew")
    urls_v_scroll.grid(row=0, column=1, sticky="ns")
    urls_h_scroll.grid(row=1, column=0, sticky="ew")
    urls_excel_tree.configure(yscrollcommand=urls_v_scroll.set, xscrollcommand=urls_h_scroll.set)

    # --- CTAs row (debajo del preview Excel) ---
    urls_actions_frame.pack(fill="x", pady=(10, 6))  # margen superior pequeño

    def _limpiar_url(url):
        """Limpia una URL: elimina espacios, saltos de línea, tabulaciones."""
        if not url:
            return ""
        url_limpia = str(url).strip()
        url_limpia = url_limpia.replace('\n', '').replace('\r', '').replace('\t', '')
        url_limpia = url_limpia.strip()
        return url_limpia

    def _validar_y_normalizar_pais(pais_abrev):
        """Valida y normaliza un código de país a minúsculas. Retorna (pais_normalizado, error)."""
        if not pais_abrev:
            return None, "País no especificado"
        pais_normalizado = str(pais_abrev).strip().lower()
        if pais_normalizado not in ABREV_PAISES_LOWER:
            return None, f"País '{pais_abrev}' no válido. Use: {', '.join(sorted(ABREV_PAISES_LOWER.keys()))}"
        return pais_normalizado, None

    def _create_validation_urls_excel_if_missing():
        excel_path = _get_validation_urls_excel_path()
        if os.path.exists(excel_path):
            return excel_path
        df = pd.DataFrame(columns=["Pais", "URL", "Formulario"])
        df.to_excel(excel_path, index=False)
        return excel_path

    def _write_url_pairs_to_excel(url_triples):
        excel_path = _create_validation_urls_excel_if_missing()
        df = pd.DataFrame(url_triples, columns=["Pais", "URL", "Formulario"])
        df.to_excel(excel_path, index=False)

    def _read_url_pairs_from_excel(require_form=True):
        excel_path = _create_validation_urls_excel_if_missing()
        try:
            df = pd.read_excel(excel_path)
        except Exception as exc:
            raise ValueError(f"No se pudo leer el Excel de URLs: {exc}") from exc

        if df.empty:
            return []

        if len(df.columns) < 3:
            raise ValueError("El Excel debe tener al menos 3 columnas: A=País, B=URL, C=Formulario.")

        pais_col = df.columns[0]
        landing_col = df.columns[1]
        form_col = df.columns[2]

        url_triples = []
        for row_idx, row in df.iterrows():
            pais_raw = str(row.get(pais_col, "") or "").strip()
            landing_url_raw = str(row.get(landing_col, "") or "").strip()
            form_url_raw = str(row.get(form_col, "") or "").strip()

            if not pais_raw and not landing_url_raw and not form_url_raw:
                continue

            # Validar y normalizar país
            pais_normalizado, pais_error = _validar_y_normalizar_pais(pais_raw)
            if pais_error:
                raise ValueError(f"Fila {row_idx + 2}: {pais_error}")

            # Limpiar URLs
            landing_url = _limpiar_url(landing_url_raw)
            form_url = _limpiar_url(form_url_raw)

            if not landing_url:
                raise ValueError(f"Fila {row_idx + 2}: URL no especificada o vacía.")
            if require_form and not form_url:
                raise ValueError(f"Fila {row_idx + 2}: Formulario no especificado o vacío.")

            url_triples.append((pais_normalizado, landing_url, form_url))

        return url_triples

    def _reload_urls_preview_from_excel():
        import importlib
        importlib.invalidate_caches()  # Forzar invalidación de caché de sistema de archivos
        for item_id in urls_excel_tree.get_children():
            urls_excel_tree.delete(item_id)

        # Forzar recarga de Excel sin caché usando pandas y openpyxl
        excel_path = _get_validation_urls_excel_path()
        try:
            # Leer el archivo con openpyxl para evitar caché de xlrd
            df = pd.read_excel(excel_path, engine="openpyxl")
        except Exception:
            # Fallback a engine por defecto si openpyxl falla
            df = pd.read_excel(excel_path)

        if df.empty or len(df.columns) < 3:
            url_triples = []
        else:
            pais_col = df.columns[0]
            landing_col = df.columns[1]
            form_col = df.columns[2]
            url_triples = []
            for row_idx, row in df.iterrows():
                pais_raw = str(row.get(pais_col, "") or "").strip()
                landing_url_raw = str(row.get(landing_col, "") or "").strip()
                form_url_raw = str(row.get(form_col, "") or "").strip()
                if not pais_raw and not landing_url_raw and not form_url_raw:
                    continue
                url_triples.append((pais_raw, landing_url_raw, form_url_raw))

        for pais, landing_url, form_url in url_triples:
            urls_excel_tree.insert("", END, values=(pais, landing_url, form_url))
        _apply_row_colors(urls_excel_tree)
        return url_triples

    def _open_urls_excel():
        excel_path = _create_validation_urls_excel_if_missing()
        try:
            os.startfile(excel_path)
        except Exception as exc:
            messagebox.showerror("Validación de campos", f"No se pudo abrir el Excel: {exc}")

    def _refresh_urls_preview():
        _reload_urls_preview_from_excel()
        status_var.set("Previsualización de Excel actualizada.")

    Button(
        urls_actions_frame,
        text="Abrir Excel",
        command=_open_urls_excel,
        bg=button_bg,
        fg=button_fg,
        relief="flat",
        font=("Segoe UI", 10, "bold"),
        cursor="hand2",
        padx=10,
        pady=3,
    ).pack(side=LEFT, padx=(0, 6))

    Button(
        urls_actions_frame,
        text="Actualizar",
        command=_refresh_urls_preview,
        bg=button_bg,
        fg=button_fg,
        relief="flat",
        font=("Segoe UI", 10, "bold"),
        cursor="hand2",
        padx=10,
        pady=3,
    ).pack(side=LEFT)

    Checkbutton(
        urls_actions_frame,
        text="Ver navegador",
        variable=ver_navegador_var,
        bg=app_bg,
        fg=text_color,
        selectcolor=entry_bg,
        activebackground=app_bg,
        activeforeground=text_color,
        font=("Segoe UI", 10, "bold"),
        cursor="hand2",
        bd=0,
    ).pack(side=LEFT, padx=(12, 0))

    def _collect_url_pairs_from_widget():
        return _read_url_pairs_from_excel(require_form=True)

    def _set_url_pairs_in_widget(url_triples):
        _write_url_pairs_to_excel(url_triples)
        _reload_urls_preview_from_excel()

    # Expander para Configuración de ID
    id_section_container = Frame(root_frame, bg=container_bg, highlightthickness=0, bd=0)
    id_section_container.pack(fill="x", padx=12, pady=(0, 10))

    expander_state = {'open': False}

    # Color más oscuro para la barra del expander
    from tkinter import colorchooser
    def darken_color(hex_color, factor=0.85):
        hex_color = hex_color.lstrip('#')
        rgb = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        dark_rgb = tuple(max(0, int(c * factor)) for c in rgb)
        return '#{:02x}{:02x}{:02x}'.format(*dark_rgb)

    expander_bar_bg = darken_color(palette["section_bg"], 0.85)

    def toggle_id_section():
        if expander_state['open']:
            mapping_container.pack_forget()
            expander_state['open'] = False
            expander_btn.config(text="▶ Configuración de ID")
        else:
            mapping_container.pack(fill="x", padx=0, pady=(0, 0))
            expander_state['open'] = True
            expander_btn.config(text="▼ Configuración de ID")

    expander_btn = Button(
        id_section_container,
        text="▶ Configuración de ID",
        font=("Segoe UI", 11, "bold"),
        bg=expander_bar_bg,
        fg="white",
        relief="flat",
        anchor="w",
        command=toggle_id_section,
        cursor="hand2",
        padx=8,
        pady=2,
        bd=0,
        highlightthickness=0,
        activebackground=expander_bar_bg,
        activeforeground="white",
    )
    expander_btn.pack(fill="x", padx=0, pady=(0, 0))

    # Margen interno para el contenido del expander
    mapping_container = Frame(id_section_container, bg=section_bg, highlightthickness=0, bd=0)
    # Empieza colapsado, no se hace pack aquí

    # Margen interno para el contenido del expander
    mapping_inner = Frame(mapping_container, bg=section_bg)
    mapping_inner.pack(fill="both", expand=True, padx=10, pady=10)

    mapping_paned = ttk.Panedwindow(mapping_inner, orient="vertical")
    mapping_paned.pack(fill="x", expand=True)

    mapping_form_container = Frame(mapping_paned, bg=section_bg)

    mapping_form = Frame(mapping_form_container, bg=section_bg)
    mapping_form.pack(fill="x", pady=(16, 8))  # Espacio extra arriba de los campos ID/Descripción

    Label(mapping_form, text="ID", bg=section_bg, fg=text_color).grid(row=0, column=0, sticky="w", pady=(0, 2))
    ttk.Entry(mapping_form, textvariable=element_id_var).grid(row=0, column=1, sticky="ew", padx=(6, 14), pady=(0, 8))
    Label(mapping_form, text="Descripción", bg=section_bg, fg=text_color).grid(row=0, column=2, sticky="w", pady=(0, 2))
    ttk.Entry(mapping_form, textvariable=field_name_var).grid(row=0, column=3, sticky="ew", padx=(6, 0), pady=(0, 8))

    controles_tipo_row = Frame(mapping_form, bg=section_bg)
    controles_tipo_row.grid(row=1, column=0, columnspan=4, sticky="w", pady=(0, 6))
    Checkbutton(
        controles_tipo_row,
        text="Dropdown",
        variable=dropdown_var,
        bg=section_bg,
        fg="white",
        activebackground=section_bg,
        activeforeground="white",
        selectcolor=section_bg,
    ).pack(side=LEFT)
    Checkbutton(
        controles_tipo_row,
        text='Inputmode = "numeric"',
        variable=teclado_mobile_var,
        bg=section_bg,
        fg="white",
        activebackground=section_bg,
        activeforeground="white",
        selectcolor=section_bg,
    ).pack(side=LEFT, padx=(12, 0))

    Label(mapping_form, text="Regex full", bg=section_bg, fg=text_color).grid(row=2, column=0, sticky="w", pady=(0, 2))
    regex_full_entry = ttk.Entry(mapping_form, textvariable=regex_full_var)
    regex_full_entry.grid(row=2, column=1, columnspan=3, sticky="ew", pady=(0, 8))

    Label(mapping_form, text="Regex char", bg=section_bg, fg=text_color).grid(row=3, column=0, sticky="w", pady=(0, 2))
    regex_char_entry = ttk.Entry(mapping_form, textvariable=regex_char_var)
    regex_char_entry.grid(row=3, column=1, columnspan=3, sticky="ew", pady=(0, 8))

    texto_prueba_label_row = Frame(mapping_form, bg=section_bg)
    texto_prueba_label_row.grid(row=4, column=0, columnspan=4, sticky="ew", pady=(0, 2))
    Label(texto_prueba_label_row, text="Texto de prueba", bg=section_bg, fg=text_color).pack(side=LEFT)

    separador_vertical = Frame(texto_prueba_label_row, bg="#cfcfcf", width=1, height=18)
    separador_vertical.pack(side=LEFT, padx=10, pady=1)
    separador_vertical.pack_propagate(False)

    Label(texto_prueba_label_row, text="País", bg=section_bg, fg=text_color).pack(side=LEFT, padx=(0, 8))

    _COUNTRY_DISPLAY_ABBREV = {
        "Argentina": "AR",
        "Bolivia": "BO",
        "Brasil": "BR",
        "Chile": "CH",
        "Colombia": "CO",
        "Ecuador": "EC",
        "Paraguay": "PA",
        "Peru": "PE",
        "Uruguay": "UY",
    }

    rule_countries_frame = Frame(texto_prueba_label_row, bg=section_bg)
    rule_countries_frame.pack(side=LEFT)
    for country_name in AVAILABLE_COUNTRIES:
        Checkbutton(
            rule_countries_frame,
            text=_COUNTRY_DISPLAY_ABBREV.get(country_name, country_name),
            variable=rule_country_vars[country_name],
            bg=section_bg,
            fg="white",
            activebackground=section_bg,
            activeforeground="white",
            selectcolor=section_bg,
        ).pack(side=LEFT, padx=(0, 6))
    test_text_widget = Text(mapping_form, width=52, height=3, bg=entry_bg, fg=entry_fg,
                            insertbackground=entry_fg, relief="flat", highlightthickness=1,
                            highlightbackground=button_bg)
    test_text_widget.grid(row=5, column=0, columnspan=4, sticky="ew", pady=(0, 8))

    mapping_form.columnconfigure(1, weight=1)
    mapping_form.columnconfigure(3, weight=1)

    # Botonera de reglas: fondo igual al contenido del expander, distribución equitativa
    mapping_buttons = Frame(mapping_form_container, bg=section_bg)
    mapping_buttons.pack(fill="x", pady=(0, 6))

    filters_frame = Frame(mapping_form_container, bg=section_bg)
    filters_frame.pack(fill="x", pady=(0, 6))

    Label(
        filters_frame,
        text="Filtro:",
        bg=section_bg,
        fg=text_color,
    ).pack(side=LEFT, padx=(0, 6))
    rule_filter_entry = Entry(filters_frame, font=("Segoe UI", 10), width=20, textvariable=rule_filter_text_var,
                              bg=entry_bg, fg=entry_fg, insertbackground=entry_fg, relief="flat",
                              highlightthickness=1, highlightbackground=button_bg)
    rule_filter_entry.pack(side=LEFT, padx=(0, 10))

    filtro_pais_values = ["Todos"] + [COUNTRY_ABBREVIATIONS[c] for c in AVAILABLE_COUNTRIES if c in COUNTRY_ABBREVIATIONS]
    filtro_combo_style = ttk.Style()
    filtro_combo_style.configure(
        "ValidationRuleFilter.TCombobox",
        fieldbackground="white",
        background="white",
        foreground="black",
    )
    filtro_combo_style.map(
        "ValidationRuleFilter.TCombobox",
        fieldbackground=[("readonly", "white")],
        selectbackground=[("readonly", "white")],
        selectforeground=[("readonly", "black")],
        foreground=[("readonly", "black")],
        background=[("readonly", "white")],
    )
    ttk.Combobox(
        filters_frame,
        textvariable=rule_filter_country_var,
        values=filtro_pais_values,
        state="readonly",
        width=8,
        style="ValidationRuleFilter.TCombobox",
    ).pack(side=LEFT)

    def _clear_rule_filters():
        rule_filter_text_var.set("")
        rule_filter_country_var.set("Todos")

    Button(
        filters_frame,
        text="Limpiar filtros",
        command=_clear_rule_filters,
        bg=app_bg,
        fg="white",
        relief="flat",
        font=("Segoe UI", 9, "bold"),
        cursor="hand2",
        padx=8,
        pady=2,
    ).pack(side=LEFT, padx=(10, 0))

    Button(
        filters_frame,
        text="Generador regex",
        command=lambda: _open_regex_generator_popup(),
        bg=app_bg,
        fg="white",
        relief="flat",
        font=("Segoe UI", 10, "bold"),
        cursor="hand2",
        padx=10,
        pady=3,
    ).pack(side=RIGHT, padx=6)
    Button(
        filters_frame,
        text="Limpiar campos",
        command=lambda: _clear_form_inputs(),
        bg=app_bg,
        fg="white",
        relief="flat",
        font=("Segoe UI", 10, "bold"),
        cursor="hand2",
        padx=10,
        pady=3,
    ).pack(side=RIGHT, padx=6)

    # Import de validaciones desde el Excel externo del CRM (nunca versionado). Cruza
    # contra field_validation_rules_<pais>.json — no contra el JSON global que usa el resto
    # de esta pestaña — porque es el archivo que realmente lee el motor de detección/
    # validación en runtime (core/base_form_filler.py, validation/selenium_validation_runner.py).
    crm_import_frame = Frame(mapping_form_container, bg=section_bg)
    crm_import_frame.pack(fill="x", pady=(0, 6))

    crm_import_pais_var = StringVar(value=AVAILABLE_COUNTRIES[0])

    Label(
        crm_import_frame,
        text="País (import Excel CRM):",
        bg=section_bg,
        fg=text_color,
    ).pack(side=LEFT, padx=(0, 6))
    ttk.Combobox(
        crm_import_frame,
        textvariable=crm_import_pais_var,
        values=list(AVAILABLE_COUNTRIES),
        state="readonly",
        width=12,
        style="ValidationRuleFilter.TCombobox",
    ).pack(side=LEFT, padx=(0, 10))

    Button(
        crm_import_frame,
        text="📥 Importar validaciones desde Excel CRM",
        command=lambda: _importar_crm_excel_pais_actual(),
        bg=button_bg,
        fg=button_fg,
        relief="flat",
        font=("Segoe UI", 10, "bold"),
        cursor="hand2",
        padx=10,
        pady=3,
    ).pack(side=LEFT, padx=(0, 6))
    Button(
        crm_import_frame,
        text="📥 Importar en TODOS los países",
        command=lambda: _importar_crm_excel_todos_los_paises(),
        bg=button_bg,
        fg=button_fg,
        relief="flat",
        font=("Segoe UI", 10, "bold"),
        cursor="hand2",
        padx=10,
        pady=3,
    ).pack(side=LEFT)

    columns = ("element_id", "dropdown", "descripcion", "dependencies", "regex_full", "regex_char", "test_text", "paises", "teclado_mobile")
    mapping_tree_frame = Frame(mapping_paned, bg=section_bg)

    mapping_tree = ttk.Treeview(mapping_tree_frame, columns=columns, show="headings", height=6, style="Section.Treeview")
    for column_id, label, width in (
        ("element_id", "ID", 140),
        ("dropdown", "Dropdown", 90),
        ("descripcion", "Descripción", 180),
        ("dependencies", "Dependencias", 220),
        ("regex_full", "Regex full", 290),
        ("regex_char", "Regex char", 180),
        ("test_text", "Texto de prueba", 180),
        ("paises", "Paises", 200),
        ("teclado_mobile", "Teclado mobile", 110),
    ):
        mapping_tree.heading(column_id, text=label)
        mapping_tree.column(column_id, width=width, stretch=(column_id in ("dependencies", "regex_full", "test_text", "paises")))
    mapping_v_scroll = ttk.Scrollbar(mapping_tree_frame, orient="vertical", command=mapping_tree.yview)
    mapping_h_scroll = ttk.Scrollbar(mapping_tree_frame, orient="horizontal", command=mapping_tree.xview)

    mapping_tree_frame.grid_columnconfigure(0, weight=1)
    mapping_tree_frame.grid_rowconfigure(0, weight=1)
    mapping_tree.grid(row=0, column=0, sticky="nsew")
    mapping_v_scroll.grid(row=0, column=1, sticky="ns")
    mapping_h_scroll.grid(row=1, column=0, sticky="ew")
    mapping_tree.configure(yscrollcommand=mapping_v_scroll.set, xscrollcommand=mapping_h_scroll.set)

    mapping_paned.add(mapping_form_container, weight=2)
    mapping_paned.add(mapping_tree_frame, weight=3)

    def _mostrar_resultado_import_crm(titulo, texto):
        popup = Toplevel()
        popup.title(titulo)
        popup.configure(bg=app_bg)
        resultado_widget = Text(popup, width=92, height=28, bg=entry_bg, fg=entry_fg, wrap="word")
        resultado_widget.insert("1.0", texto)
        resultado_widget.configure(state="disabled")
        resultado_widget.pack(fill=BOTH, expand=True, padx=10, pady=10)
        Button(
            popup, text="Cerrar", command=popup.destroy,
            bg=button_bg, fg=button_fg, relief="flat", cursor="hand2", padx=10, pady=3,
        ).pack(pady=(0, 10))

    def _formatear_resumen_import_crm(pais, comparacion, cambios):
        pendientes = len(comparacion["incompletos"]) - cambios["completados"]
        return (
            f"{pais}: {len(comparacion['faltantes'])} faltantes "
            f"({cambios['agregados']} agregados al JSON), "
            f"{cambios['completados']} reglas derivadas de la prosa del excel, "
            f"{cambios['recalculados']} recalculadas por conflicto con el mensaje de error "
            f"(regex anterior guardado en regex_full_previo), "
            f"{pendientes} obligatorios siguen sin regex — revisar a mano, "
            f"{len(comparacion['ok'])} ya cubiertos."
        )

    def _importar_crm_para_pais(importer, pais):
        """Compara y guarda field_validation_rules_<pais>.json. Completa las reglas que la
        prosa del excel permite derivar y recalcula las que su mensaje de error contradice,
        preservando el regex anterior (ver CrmValidacionesImporter.aplicar_merge)."""
        reglas = _load_rules_file(pais=pais)
        comparacion = importer.comparar(pais, reglas)
        actualizado, cambios = importer.aplicar_merge(pais, comparacion, reglas)
        if any(cambios.values()):
            _save_rules_file(actualizado, pais=pais)
        if cambios["agregados"]:
            registrar_campos_importados_desde_excel(pais, comparacion["faltantes"])
        guardar_reporte(pais, construir_reporte(pais, pais, comparacion))
        return comparacion, cambios

    def _importar_crm_excel_pais_actual():
        importer = CrmValidacionesImporter()
        disponible, info = importer.disponible()
        if not disponible:
            messagebox.showwarning("Importar validaciones desde Excel CRM", info)
            return

        pais = crm_import_pais_var.get()
        try:
            comparacion, cambios = _importar_crm_para_pais(importer, pais)
        except Exception as exc:
            LOGGER.exception("Error importando validaciones del CRM para %s", pais)
            messagebox.showerror("Importar validaciones desde Excel CRM", f"Error importando {pais}: {exc}")
            return

        _mostrar_resultado_import_crm(
            "Importar validaciones desde Excel CRM",
            _formatear_resumen_import_crm(pais, comparacion, cambios),
        )
        status_var.set(
            f"Import Excel CRM ({pais}): {cambios['agregados']} agregado(s), "
            f"{cambios['completados']} derivada(s), {cambios['recalculados']} recalculada(s)."
        )

    def _importar_crm_excel_todos_los_paises():
        importer = CrmValidacionesImporter()
        disponible, info = importer.disponible()
        if not disponible:
            messagebox.showwarning("Importar en TODOS los países", info)
            return

        lineas = []
        totales = {"agregados": 0, "completados": 0, "recalculados": 0}
        for pais in AVAILABLE_COUNTRIES:
            try:
                comparacion, cambios = _importar_crm_para_pais(importer, pais)
            except Exception as exc:
                LOGGER.exception("Error importando validaciones del CRM para %s", pais)
                lineas.append(f"{pais}: ERROR — {exc}")
                continue
            for clave in totales:
                totales[clave] += cambios[clave]
            lineas.append(_formatear_resumen_import_crm(pais, comparacion, cambios))

        _mostrar_resultado_import_crm("Importar en TODOS los países", "\n".join(lineas))
        status_var.set(
            f"Import Excel CRM (todos los países): {totales['agregados']} agregado(s), "
            f"{totales['completados']} derivada(s), {totales['recalculados']} recalculada(s)."
        )

    def _clear_form_inputs():
        field_name_var.set("")
        element_id_var.set("")
        dropdown_var.set(False)
        regex_full_var.set("")
        regex_char_var.set("")
        teclado_mobile_var.set(False)
        for country_name in AVAILABLE_COUNTRIES:
            rule_country_vars[country_name].set(False)
        test_text_widget.delete("1.0", END)
        current_error_message_patterns.clear()
        current_dependencies.clear()
        current_dropdown_error_message_var.set("")
        mapping_tree.selection_remove(mapping_tree.selection())
        upsert_button_text_var.set("Agregar regla")

    def _parse_rule_countries(raw_value):
        # Si no viene valor, aplica a todos los paises.
        if raw_value is None:
            return list(AVAILABLE_COUNTRIES)

        if isinstance(raw_value, list):
            candidates = raw_value
        else:
            raw_text = str(raw_value or "")
            if not raw_text.strip():
                return list(AVAILABLE_COUNTRIES)
            candidates = [item.strip() for item in raw_text.split("|") if item.strip()]

        normalized = []
        for raw_country in candidates:
            token = str(raw_country or "").strip()
            if not token:
                continue

            lower_token = token.lower()
            country_name = None

            if token in AVAILABLE_COUNTRIES:
                country_name = token
            elif lower_token in COUNTRY_NAMES_LOWER:
                country_name = COUNTRY_NAMES_LOWER[lower_token]
            elif token.upper() in ABBREVIATION_TO_COUNTRY:
                country_name = ABBREVIATION_TO_COUNTRY[token.upper()]
            elif lower_token in ABREV_TO_COUNTRY_LOWER:
                country_name = ABREV_TO_COUNTRY_LOWER[lower_token]

            if country_name and country_name not in normalized:
                normalized.append(country_name)

        # Si había contenido pero no se reconocio ningun pais, no forzar "Todos".
        return normalized

    def _normalize_selected_country(country_value):
        token = str(country_value or "").strip()
        if not token:
            return ""
        if token in AVAILABLE_COUNTRIES:
            return token

        lower_token = token.lower()
        if lower_token in COUNTRY_NAMES_LOWER:
            return COUNTRY_NAMES_LOWER[lower_token]
        if token.upper() in ABBREVIATION_TO_COUNTRY:
            return ABBREVIATION_TO_COUNTRY[token.upper()]
        if lower_token in ABREV_TO_COUNTRY_LOWER:
            return ABREV_TO_COUNTRY_LOWER[lower_token]
        return token

    def _collect_selected_browsers():
        return [browser_name for browser_name, var in browser_vars.items() if bool(var.get())]

    def _collect_selected_viewports():
        return [viewport_name for viewport_name, var in viewport_vars.items() if bool(var.get())]

    def _collect_rule_countries():
        selected = [country_name for country_name, var in rule_country_vars.items() if bool(var.get())]
        return selected or list(AVAILABLE_COUNTRIES)

    def _set_rule_countries(selected_countries):
        normalized = _parse_rule_countries(selected_countries)
        for country_name in AVAILABLE_COUNTRIES:
            rule_country_vars[country_name].set(country_name in normalized)

    def _format_rule_countries_abbrev(countries):
        normalized = _parse_rule_countries(countries)
        if set(normalized) == set(AVAILABLE_COUNTRIES):
            return "Todos"
        return "|".join(COUNTRY_ABBREVIATIONS.get(country_name, country_name) for country_name in normalized)

    def _normalize_dependencies(raw_dependencies):
        if not raw_dependencies:
            return []

        if isinstance(raw_dependencies, dict):
            raw_dependencies = [raw_dependencies]
        if not isinstance(raw_dependencies, list):
            return []

        normalized = []
        for raw_dependency in raw_dependencies:
            if not isinstance(raw_dependency, dict):
                continue

            dependency_id = str(
                raw_dependency.get("element_id")
                or raw_dependency.get("id")
                or raw_dependency.get("field_id")
                or ""
            ).strip()
            dependency_value = str(
                raw_dependency.get("value")
                or raw_dependency.get("option")
                or raw_dependency.get("valor")
                or ""
            ).strip()

            if not dependency_id:
                continue

            dep = {"element_id": dependency_id}
            if dependency_value:
                dep["value"] = dependency_value
            normalized.append(dep)

        return normalized

    def _format_dependencies(dependencies):
        normalized = _normalize_dependencies(dependencies)
        if not normalized:
            return ""
        def format_dep(dep):
            val = dep.get('value', None)
            if val:
                return f"{dep.get('element_id', '')} - {val}"
            else:
                return f"{dep.get('element_id', '')} - valor: No se especificó, se completará automáticamente"
        return "; ".join(format_dep(dep) for dep in normalized)

    def _get_active_dependencies():
        row_index = _get_selected_mapping_index()
        if row_index is not None:
            return deepcopy(_normalize_dependencies(mapping_rows[row_index].get("dependencies") or []))
        return deepcopy(_normalize_dependencies(current_dependencies))

    def _set_active_dependencies(dependencies):
        normalized = _normalize_dependencies(dependencies)
        row_index = _get_selected_mapping_index()
        if row_index is not None:
            mapping_rows[row_index]["dependencies"] = deepcopy(normalized)
            _refresh_mapping_tree_view()
            _save_rules_from_ui(require_form_urls=False)
        else:
            current_dependencies.clear()
            current_dependencies.extend(deepcopy(normalized))

    mapping_rows = []
    current_error_message_patterns = []
    current_dropdown_error_message_var = StringVar(value="")
    current_dependencies = []

    def _normalize_error_message_patterns(raw_patterns):
        if not isinstance(raw_patterns, list):
            return []

        normalized = []
        for raw_rule in raw_patterns:
            if not isinstance(raw_rule, dict):
                continue
            regex_text = str(raw_rule.get("regex") or "").strip()
            message_text = str(raw_rule.get("mensaje") or raw_rule.get("message") or "").strip()
            if not regex_text or not message_text:
                continue
            normalized.append({"regex": regex_text, "mensaje": message_text})

        return normalized

    def _validate_regex_pattern(pattern_text, field_label):
        normalized_pattern = str(pattern_text or "").strip()
        if not normalized_pattern:
            messagebox.showerror("Validación de campos", f"{field_label} es obligatorio.")
            return False
        try:
            re.compile(normalized_pattern)
            return True
        except re.error as exc:
            messagebox.showerror(
                "Validación de campos",
                f"{field_label} inválido:\n{normalized_pattern}\n\nDetalle: {exc}",
            )
            return False

    def _get_selected_mapping_index():
        selected = mapping_tree.selection()
        if not selected:
            return None
        row_index = int(selected[0])
        if row_index < 0 or row_index >= len(mapping_rows):
            return None
        return row_index

    def _get_active_error_message_patterns():
        row_index = _get_selected_mapping_index()
        if row_index is not None:
            return deepcopy(_normalize_error_message_patterns(mapping_rows[row_index].get("error_message_patterns") or []))
        return deepcopy(_normalize_error_message_patterns(current_error_message_patterns))

    def _set_active_error_message_patterns(patterns):
        normalized = _normalize_error_message_patterns(patterns)
        row_index = _get_selected_mapping_index()
        if row_index is not None:
            mapping_rows[row_index]["error_message_patterns"] = deepcopy(normalized)
            _refresh_mapping_tree_view()
            _save_rules_from_ui(require_form_urls=False)
        else:
            current_error_message_patterns.clear()
            current_error_message_patterns.extend(deepcopy(normalized))

    def _get_active_dropdown_error_message():
        row_index = _get_selected_mapping_index()
        if row_index is not None:
            return str(mapping_rows[row_index].get("dropdown_error_message") or "").strip()
        return str(current_dropdown_error_message_var.get() or "").strip()

    def _set_active_dropdown_error_message(message_text):
        normalized_message = str(message_text or "").strip()
        row_index = _get_selected_mapping_index()
        if row_index is not None:
            mapping_rows[row_index]["dropdown_error_message"] = normalized_message
            _refresh_mapping_tree_view()
            _save_rules_from_ui(require_form_urls=False)
        else:
            current_dropdown_error_message_var.set(normalized_message)

    _dropdown_disabled_style = "DropdownDisabled.TEntry"
    _dropdown_normal_style = "TEntry"
    _dropdown_entry_style = ttk.Style()
    _dropdown_entry_style.configure(_dropdown_disabled_style, fieldbackground="#656565", foreground="#aaaaaa")

    def _apply_dropdown_mode(clear_values=False):
        if dropdown_var.get():
            regex_full_entry.config(state="disabled", style=_dropdown_disabled_style)
            regex_char_entry.config(state="disabled", style=_dropdown_disabled_style)
            test_text_widget.config(state="normal")
            if clear_values:
                regex_full_var.set("")
                regex_char_var.set("")
                test_text_widget.delete("1.0", END)
            test_text_widget.config(state="disabled", bg="#656565", fg="#aaaaaa")
        else:
            regex_full_entry.config(state="normal", style=_dropdown_normal_style)
            regex_char_entry.config(state="normal", style=_dropdown_normal_style)
            test_text_widget.config(state="normal", bg=entry_bg, fg=entry_fg)

    def _on_dropdown_toggle(*_):
        _apply_dropdown_mode(clear_values=True)

    dropdown_var.trace_add("write", _on_dropdown_toggle)
    _apply_dropdown_mode(clear_values=False)

    def _set_popup_initial_size(window, min_width=640, min_height=360):
        # Compute a content-driven initial size bounded by the current screen.
        window.update_idletasks()
        req_width = window.winfo_reqwidth()
        req_height = window.winfo_reqheight()
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()

        width = max(min_width, req_width + 20)
        height = max(min_height, req_height + 20)

        max_width = int(screen_width * 0.92)
        max_height = int(screen_height * 0.88)
        width = min(width, max_width)
        height = min(height, max_height)

        window.minsize(min_width, min_height)
        window.geometry(f"{width}x{height}")

    def _escape_for_char_class(chars):
        escaped = ""
        for char in chars:
            if char in "\\-]^":
                escaped += "\\" + char
            else:
                escaped += char
        return escaped

    def _build_basic_regex(required_empty, email, lowercase, uppercase, accents, spaces, symbols, numbers, min_len, max_len, max_double_chars=False, require_consonant=False, require_vocal=False, require_bos=False, bos_text="", not_all_same=False, all_same_chars=False):
        if required_empty:
            return "^$", "^$"

        if email:
            email_regex = "^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)+$"
            email_char_regex = "^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~@-]*$"
            return email_regex, email_char_regex

        if all_same_chars and not_all_same:
            return "", ""

        char_parts = []
        if lowercase:
            char_parts.append("a-z")
        if uppercase:
            char_parts.append("A-Z")
        if accents:
            char_parts.append("\u00E1\u00E9\u00ED\u00F3\u00FA\u00C1\u00C9\u00CD\u00D3\u00DA\u00F1\u00D1")
        if spaces:
            char_parts.append(" ")
        if symbols:
            char_parts.append(_escape_for_char_class(".,;:!?\"'()_-+/@#%&*"))
        if numbers:
            char_parts.append("0-9")

        charset = "".join(char_parts)
        if not charset:
            return "", ""

        min_text = str(min_len or "").strip()
        max_text = str(max_len or "").strip()
        min_value = None
        max_value = None

        if min_text:
            if not min_text.isdigit():
                return "", ""
            min_value = int(min_text)
        if max_text:
            if not max_text.isdigit():
                return "", ""
            max_value = int(max_text)

        if min_value is not None and max_value is not None and min_value > max_value:
            return "", ""

        if all_same_chars:
            escaped_charset = charset
            if min_value is not None and max_value is not None:
                if min_value == max_value:
                    repetition = str(min_value - 1)
                    inner_quantifier = "{" + repetition + "}"
                else:
                    lower_rep = min_value - 1
                    upper_rep = max_value - 1
                    inner_quantifier = "{" + str(lower_rep) + "," + str(upper_rep) + "}"
            elif min_value is not None:
                lower_rep = min_value - 1
                inner_quantifier = "{" + str(lower_rep) + ",}"
            elif max_value is not None:
                upper_rep = max_value - 1
                inner_quantifier = "{0," + str(upper_rep) + "}"
            else:
                inner_quantifier = "*"
            return "^([" + escaped_charset + "])\\1" + inner_quantifier + "$", "^[" + escaped_charset + "]*$"

        vowels = []
        consonants = []
        if lowercase:
            vowels.append("aeiou")
            consonants.append("bcdfghjklmnpqrstvwxyz")
        if uppercase:
            vowels.append("AEIOU")
            consonants.append("BCDFGHJKLMNPQRSTVWXYZ")
        if accents:
            vowels.append("\u00E1\u00E9\u00ED\u00F3\u00FA\u00C1\u00C9\u00CD\u00D3\u00DA")
            consonants.append("\u00F1\u00D1")

        vowel_charset = "".join(dict.fromkeys("".join(vowels)))
        consonant_charset = "".join(dict.fromkeys("".join(consonants)))

        if require_vocal and not vowel_charset:
            return "", ""
        if require_consonant and not consonant_charset:
            return "", ""

        if min_value is None and max_value is None:
            quantifier = "+"
        elif min_value is not None and max_value is None:
            quantifier = "{" + str(min_value) + ",}"
        elif min_value is None and max_value is not None:
            quantifier = "{0," + str(max_value) + "}"
        elif min_value == max_value:
            quantifier = "{" + str(min_value) + "}"
        else:
            quantifier = "{" + str(min_value) + "," + str(max_value) + "}"

        regex_char = "[" + charset + "]"

        bos_options = []
        if require_bos:
            bos_options = [option.strip() for option in str(bos_text or "").split(";") if option.strip()]
            if not bos_options:
                return "", ""
            allowed_prefix_pattern = "^" + regex_char + "+$"
            for option in bos_options:
                if not re.fullmatch(allowed_prefix_pattern, option):
                    return "", ""

        lookaheads = ""
        if max_double_chars:
            lookaheads += "(?!.*(.)(\\1){2})"
        if not_all_same:
            lookaheads += "(?!(.)(\\1+$))"
        if require_vocal:
            lookaheads += "(?=.*[" + vowel_charset + "])"
        if require_consonant:
            lookaheads += "(?=.*[" + consonant_charset + "])"

        if bos_options:
            escaped_options = [re.escape(option) for option in bos_options]
            lookaheads += "(?=(?:" + "|".join(escaped_options) + "))"
        regex_full = "^" + lookaheads + regex_char + quantifier + "$"
        return regex_full, regex_char

    def _open_regex_generator_popup(target_regex_var=None):
        generator_popup = Toplevel(root_frame)
        generator_popup.title("Generador de regex básico")
        generator_popup.configure(bg=popup_bg)
        generator_popup.resizable(True, True)

        icon_path = os.path.join(_get_asset_dir(), "icon.ico")
        if os.path.exists(icon_path):
            try:
                generator_popup.iconbitmap(icon_path)
            except Exception:
                try:
                    generator_popup.iconbitmap(default=icon_path)
                except Exception:
                    pass

        lowercase_var = BooleanVar(value=True)
        uppercase_var = BooleanVar(value=False)
        accents_var = BooleanVar(value=False)
        spaces_var = BooleanVar(value=True)
        symbols_var = BooleanVar(value=False)
        numbers_var = BooleanVar(value=False)
        email_var = BooleanVar(value=False)
        required_var = BooleanVar(value=False)
        max_double_chars_var = BooleanVar(value=False)
        require_consonant_var = BooleanVar(value=False)
        require_vocal_var = BooleanVar(value=False)
        require_bos_var = BooleanVar(value=False)
        bos_text_var = StringVar(value="")
        not_all_same_var = BooleanVar(value=False)
        all_same_chars_var = BooleanVar(value=False)
        min_len_var = StringVar(value="")
        max_len_var = StringVar(value="")
        preview_full_var = StringVar(value="")
        preview_char_var = StringVar(value="")
        helper_var = StringVar(value="Seleccioná opciones para generar el regex.")

        options_frame = Frame(generator_popup, bg=popup_bg)
        options_frame.pack(fill="x", padx=12, pady=(12, 8))

        Checkbutton(
            options_frame,
            text="Letras minúsculas",
            variable=lowercase_var,
            bg=popup_bg,
            fg=popup_text,
            activebackground=popup_bg,
            activeforeground=popup_text,
            selectcolor=popup_bg,
        ).grid(row=0, column=0, sticky="w", padx=(0, 10), pady=(0, 6))
        Checkbutton(
            options_frame,
            text="Letras mayúsculas",
            variable=uppercase_var,
            bg=popup_bg,
            fg=popup_text,
            activebackground=popup_bg,
            activeforeground=popup_text,
            selectcolor=popup_bg,
        ).grid(row=0, column=1, sticky="w", padx=(0, 10), pady=(0, 6))
        Checkbutton(
            options_frame,
            text="Acentos",
            variable=accents_var,
            bg=popup_bg,
            fg=popup_text,
            activebackground=popup_bg,
            activeforeground=popup_text,
            selectcolor=popup_bg,
        ).grid(row=0, column=2, sticky="w", padx=(0, 10), pady=(0, 6))
        Checkbutton(
            options_frame,
            text="Espacios",
            variable=spaces_var,
            bg=popup_bg,
            fg=popup_text,
            activebackground=popup_bg,
            activeforeground=popup_text,
            selectcolor=popup_bg,
        ).grid(row=1, column=0, sticky="w", padx=(0, 10), pady=(0, 6))
        Checkbutton(
            options_frame,
            text="Símbolos",
            variable=symbols_var,
            bg=popup_bg,
            fg=popup_text,
            activebackground=popup_bg,
            activeforeground=popup_text,
            selectcolor=popup_bg,
        ).grid(row=1, column=1, sticky="w", padx=(0, 10), pady=(0, 6))
        Checkbutton(
            options_frame,
            text="Números",
            variable=numbers_var,
            bg=popup_bg,
            fg=popup_text,
            activebackground=popup_bg,
            activeforeground=popup_text,
            selectcolor=popup_bg,
        ).grid(row=1, column=2, sticky="w", padx=(0, 10), pady=(0, 6))
        Checkbutton(
            options_frame,
            text="Máx. 2 iguales seguidos",
            variable=max_double_chars_var,
            bg=popup_bg,
            fg=popup_text,
            activebackground=popup_bg,
            activeforeground=popup_text,
            selectcolor=popup_bg,
        ).grid(row=2, column=0, sticky="w", padx=(0, 10), pady=(0, 6))
        Checkbutton(
            options_frame,
            text="No todos iguales",
            variable=not_all_same_var,
            bg=popup_bg,
            fg=popup_text,
            activebackground=popup_bg,
            activeforeground=popup_text,
            selectcolor=popup_bg,
        ).grid(row=2, column=1, sticky="w", padx=(0, 10), pady=(0, 6))
        Checkbutton(
            options_frame,
            text="Todos iguales",
            variable=all_same_chars_var,
            bg=popup_bg,
            fg=popup_text,
            activebackground=popup_bg,
            activeforeground=popup_text,
            selectcolor=popup_bg,
        ).grid(row=2, column=2, sticky="w", padx=(0, 10), pady=(0, 6))
        Checkbutton(
            options_frame,
            text="Al menos una vocal",
            variable=require_vocal_var,
            bg=popup_bg,
            fg=popup_text,
            activebackground=popup_bg,
            activeforeground=popup_text,
            selectcolor=popup_bg,
        ).grid(row=3, column=0, sticky="w", padx=(0, 10), pady=(0, 6))
        Checkbutton(
            options_frame,
            text="Al menos una consonante",
            variable=require_consonant_var,
            bg=popup_bg,
            fg=popup_text,
            activebackground=popup_bg,
            activeforeground=popup_text,
            selectcolor=popup_bg,
        ).grid(row=3, column=1, sticky="w", padx=(0, 10), pady=(0, 6))
        Checkbutton(
            options_frame,
            text="Email",
            variable=email_var,
            bg=popup_bg,
            fg=popup_text,
            activebackground=popup_bg,
            activeforeground=popup_text,
            selectcolor=popup_bg,
        ).grid(row=3, column=2, sticky="w", padx=(0, 10), pady=(0, 6))
        Checkbutton(
            options_frame,
            text="Campo obligatorio",
            variable=required_var,
            bg=popup_bg,
            fg=popup_text,
            activebackground=popup_bg,
            activeforeground=popup_text,
            selectcolor=popup_bg,
        ).grid(row=4, column=0, sticky="w", padx=(0, 10), pady=(0, 4))
        Label(
            options_frame,
            text=(
                "*Campo obligatorio: solo para mostrar mensaje de error por obligatoriedad. "
                "No combinar con otros ítems."
            ),
            bg=popup_bg,
            fg=popup_text,
            font=("Segoe UI", 8, "italic"),
            justify="left",
            wraplength=520,
            anchor="w",
        ).grid(row=5, column=0, columnspan=3, sticky="w", pady=(0, 8))

        bos_frame = Frame(generator_popup, bg=popup_bg)
        bos_frame.pack(fill="x", padx=12, pady=(0, 8))
        Checkbutton(
            bos_frame,
            text="Debe empezar por",
            variable=require_bos_var,
            bg=popup_bg,
            fg=popup_text,
            activebackground=popup_bg,
            activeforeground=popup_text,
            selectcolor=popup_bg,
        ).pack(side=LEFT)
        ttk.Entry(bos_frame, textvariable=bos_text_var, width=40).pack(side=LEFT, padx=(8, 0))
        Label(
            bos_frame,
            text="(separá múltiples opciones con ;)",
            bg=popup_bg,
            fg=popup_text,
            font=("Segoe UI", 8, "italic"),
        ).pack(side=LEFT, padx=(8, 0))

        length_frame = Frame(generator_popup, bg=popup_bg)
        length_frame.pack(fill="x", padx=12, pady=(0, 8))
        Label(length_frame, text="Longitud minima", bg=popup_bg, fg=popup_text).pack(side=LEFT)
        ttk.Entry(length_frame, textvariable=min_len_var, width=8).pack(side=LEFT, padx=(8, 16))
        Label(length_frame, text="Longitud maxima", bg=popup_bg, fg=popup_text).pack(side=LEFT)
        ttk.Entry(length_frame, textvariable=max_len_var, width=8).pack(side=LEFT, padx=(8, 0))

        preview_frame = Frame(generator_popup, bg=popup_bg)
        preview_frame.pack(fill="both", expand=True, padx=12, pady=(0, 8))
        Label(preview_frame, text="Preview regex full - Mensaje de error", bg=popup_bg, fg=popup_text, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        ttk.Entry(preview_frame, textvariable=preview_full_var).pack(fill="x", pady=(4, 6))
        Label(preview_frame, text="Preview regex char", bg=popup_bg, fg=popup_text, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        ttk.Entry(preview_frame, textvariable=preview_char_var).pack(fill="x", pady=(4, 6))
        Label(preview_frame, textvariable=helper_var, bg=popup_bg, fg=popup_text).pack(anchor="w")

        def _refresh_preview(*_):
            regex_full, regex_char = _build_basic_regex(
                required_empty=bool(required_var.get()),
                email=bool(email_var.get()),
                lowercase=bool(lowercase_var.get()),
                uppercase=bool(uppercase_var.get()),
                accents=bool(accents_var.get()),
                spaces=bool(spaces_var.get()),
                symbols=bool(symbols_var.get()),
                numbers=bool(numbers_var.get()),
                min_len=min_len_var.get(),
                max_len=max_len_var.get(),
                max_double_chars=bool(max_double_chars_var.get()),
                require_consonant=bool(require_consonant_var.get()),
                require_vocal=bool(require_vocal_var.get()),
                require_bos=bool(require_bos_var.get()),
                bos_text=bos_text_var.get(),
                not_all_same=bool(not_all_same_var.get()),
                all_same_chars=bool(all_same_chars_var.get()),
            )
            preview_full_var.set(regex_full)
            preview_char_var.set(regex_char)
            if regex_full and regex_char:
                helper_var.set("Regex full y regex char listos para copiar.")
            else:
                helper_var.set("Revisá selección de opciones, dependencias lógicas y longitudes (solo números y min <= max).")

        for variable in (
            lowercase_var,
            uppercase_var,
            accents_var,
            spaces_var,
            symbols_var,
            numbers_var,
            email_var,
            required_var,
            max_double_chars_var,
            require_consonant_var,
            require_vocal_var,
            require_bos_var,
            bos_text_var,
            not_all_same_var,
            all_same_chars_var,
            min_len_var,
            max_len_var,
        ):
            variable.trace_add("write", _refresh_preview)

        def _copy_preview_full():
            regex_text = preview_full_var.get().strip()
            if not regex_text:
                messagebox.showerror("Generador de regex", "No hay regex full generado para copiar.")
                return
            generator_popup.clipboard_clear()
            generator_popup.clipboard_append(regex_text)
            helper_var.set("Regex full copiado al portapapeles.")

        def _copy_preview_char():
            regex_text = preview_char_var.get().strip()
            if not regex_text:
                messagebox.showerror("Generador de regex", "No hay regex char generado para copiar.")
                return
            generator_popup.clipboard_clear()
            generator_popup.clipboard_append(regex_text)
            helper_var.set("Regex char copiado al portapapeles.")

        buttons_frame = Frame(generator_popup, bg=popup_bg)
        buttons_frame.pack(fill="x", padx=12, pady=(0, 12))
        Button(
            buttons_frame,
            text="Copiar regex full",
            command=_copy_preview_full,
            bg=popup_button_bg,
            fg=popup_button_fg,
            relief="flat",
            font=("Segoe UI", 9, "bold"),
            cursor="hand2",
            padx=10,
            pady=3,
        ).pack(side=LEFT)
        Button(
            buttons_frame,
            text="Copia regex char",
            command=_copy_preview_char,
            bg=popup_button_bg,
            fg=popup_button_fg,
            relief="flat",
            font=("Segoe UI", 9, "bold"),
            cursor="hand2",
            padx=10,
            pady=3,
        ).pack(side=LEFT, padx=(8, 0))
        Button(
            buttons_frame,
            text="Cerrar",
            command=generator_popup.destroy,
            bg=popup_button_bg,
            fg=popup_button_fg,
            relief="flat",
            font=("Segoe UI", 9, "bold"),
            cursor="hand2",
            padx=10,
            pady=3,
        ).pack(side=RIGHT)

        _refresh_preview()
        _set_popup_initial_size(generator_popup, min_width=680, min_height=380)

    def _open_dependencies_popup():
        target_id = element_id_var.get().strip()
        if not target_id:
            row_index = _get_selected_mapping_index()
            if row_index is not None:
                target_id = str(mapping_rows[row_index].get("element_id") or "").strip()
        if not target_id:
            messagebox.showerror("Dependencias", "Completá o seleccioná un ID antes de configurar dependencias.")
            return

        popup = Toplevel(root_frame)
        popup.title(f"Dependencias - {target_id}")
        popup.configure(bg=popup_bg)
        popup.transient(root_frame.winfo_toplevel())
        popup.resizable(True, True)
        # Icono
        icon_path = os.path.join(_get_asset_dir(), "icon.ico")
        if os.path.exists(icon_path):
            try:
                popup.iconbitmap(icon_path)
            except Exception:
                try:
                    popup.iconbitmap(default=icon_path)
                except Exception:
                    pass

        dependencies = _get_active_dependencies()
        dependency_id_var = StringVar(value="")
        dependency_value_var = StringVar(value="")

        Label(
            popup,
            text=f"Regla ID: {target_id}",
            bg=popup_bg,
            fg=popup_text,
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", padx=12, pady=(10, 4))

        Label(
            popup,
            text="Definí de qué campo depende esta regla y qué valor debe tener para validarla.",
            bg=popup_bg,
            fg=popup_text,
            font=("Segoe UI", 9),
            justify="left",
        ).pack(anchor="w", padx=12, pady=(0, 8))

        form_row = Frame(popup, bg=popup_bg)
        form_row.pack(fill="x", padx=12, pady=(0, 8))
        Label(form_row, text="ID dependiente", bg=popup_bg, fg=popup_text).grid(row=0, column=0, sticky="w")
        Label(form_row, text="Valor", bg=popup_bg, fg=popup_text).grid(row=0, column=1, sticky="w", padx=(10, 0))
        ttk.Entry(form_row, textvariable=dependency_id_var).grid(row=1, column=0, sticky="ew", pady=(2, 0))
        ttk.Entry(form_row, textvariable=dependency_value_var).grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=(2, 0))
        form_row.grid_columnconfigure(0, weight=1)
        form_row.grid_columnconfigure(1, weight=1)

        tree_frame = Frame(popup, bg=popup_bg)
        tree_frame.pack(fill="both", expand=True, padx=12, pady=(0, 8))
        dep_tree = ttk.Treeview(
            tree_frame,
            columns=("element_id", "value"),
            show="headings",
            height=7,
            style="Section.Treeview",
        )
        dep_tree.heading("element_id", text="ID dependiente")
        dep_tree.heading("value", text="Valor")
        dep_tree.column("element_id", width=260, stretch=True)
        dep_tree.column("value", width=260, stretch=True)
        dep_tree.pack(fill="both", expand=True)

        def _refresh_dependencies_tree():
            for item_id in dep_tree.get_children():
                dep_tree.delete(item_id)
            for index, dep in enumerate(dependencies):
                # Si no hay 'value', mostrar vacío
                dep_tree.insert(
                    "",
                    END,
                    iid=str(index),
                    values=(dep.get("element_id", ""), dep.get("value", "")),
                )
            _apply_row_colors(dep_tree)

        def _load_selected_dependency(_event=None):
            selected = dep_tree.selection()
            if not selected:
                return
            row_index = int(selected[0])
            if row_index < 0 or row_index >= len(dependencies):
                return
            row = dependencies[row_index]
            dependency_id_var.set(str(row.get("element_id") or ""))
            # Si no hay 'value', dejar el campo vacío
            dependency_value_var.set(str(row.get("value")) if "value" in row else "")

        def _upsert_dependency():
            dep_id = dependency_id_var.get().strip()
            dep_value = dependency_value_var.get().strip()
            if not dep_id:
                messagebox.showerror("Dependencias", "El ID dependiente es obligatorio.")
                return

            new_dep = {"element_id": dep_id}
            if dep_value:
                new_dep["value"] = dep_value
            selected = dep_tree.selection()
            if selected:
                row_index = int(selected[0])
                if 0 <= row_index < len(dependencies):
                    dependencies[row_index] = new_dep
            else:
                dependencies.append(new_dep)

            _set_active_dependencies(dependencies)
            _refresh_dependencies_tree()
            dependency_id_var.set("")
            dependency_value_var.set("")
            _autosave_mapping_to_json()

        def _remove_dependency():
            selected = dep_tree.selection()
            if not selected:
                return
            row_index = int(selected[0])
            if 0 <= row_index < len(dependencies):
                dependencies.pop(row_index)
            _set_active_dependencies(dependencies)
            _refresh_dependencies_tree()
            dependency_id_var.set("")
            dependency_value_var.set("")
            _autosave_mapping_to_json()

        actions_row = Frame(popup, bg=popup_bg)
        actions_row.pack(fill="x", padx=12, pady=(0, 12))
        Button(
            actions_row,
            text="Agregar / Editar dependencia",
            command=_upsert_dependency,
            bg=popup_button_bg,
            fg=popup_button_fg,
            relief="flat",
            font=("Segoe UI", 9, "bold"),
            cursor="hand2",
            padx=10,
            pady=3,
        ).pack(side=LEFT, padx=(0, 6))
        Button(
            actions_row,
            text="Eliminar dependencia",
            command=_remove_dependency,
            bg=popup_button_bg,
            fg=popup_button_fg,
            relief="flat",
            font=("Segoe UI", 9, "bold"),
            cursor="hand2",
            padx=10,
            pady=3,
        ).pack(side=LEFT)
        Button(
            actions_row,
            text="Cerrar",
            command=popup.destroy,
            bg=popup_button_bg,
            fg=popup_button_fg,
            relief="flat",
            font=("Segoe UI", 9, "bold"),
            cursor="hand2",
            padx=10,
            pady=3,
        ).pack(side=RIGHT)

        def on_close():
            _autosave_mapping_to_json()
            popup.destroy()

        popup.protocol("WM_DELETE_WINDOW", on_close)

        dep_tree.bind("<<TreeviewSelect>>", _load_selected_dependency)
        _refresh_dependencies_tree()
        _set_popup_initial_size(popup, min_width=760, min_height=420)

    def _open_error_messages_popup():
        target_id = element_id_var.get().strip()
        if not target_id:
            row_index = _get_selected_mapping_index()
            if row_index is not None:
                target_id = str(mapping_rows[row_index].get("element_id") or "").strip()
        if not target_id:
            messagebox.showerror("Mensajes de error", "Completá o seleccioná un ID antes de configurar mensajes.")
            return

        is_dropdown = bool(dropdown_var.get())
        row_index = _get_selected_mapping_index()
        if row_index is not None:
            is_dropdown = bool(mapping_rows[row_index].get("dropdown", False))

        if is_dropdown:
            popup = Toplevel(root_frame)
            popup.title(f"Mensaje de error dropdown - {target_id}")
            popup.configure(bg=popup_bg)
            popup.transient(root_frame.winfo_toplevel())
            popup.resizable(True, True)
            # Icono
            icon_path = os.path.join(_get_asset_dir(), "icon.ico")
            if os.path.exists(icon_path):
                try:
                    popup.iconbitmap(icon_path)
                except Exception:
                    try:
                        popup.iconbitmap(default=icon_path)
                    except Exception:
                        pass

            dropdown_message_var = StringVar(value=_get_active_dropdown_error_message())

            Label(
                popup,
                text=f"Configuración de mensaje para dropdown ID: {target_id}",
                bg=popup_bg,
                fg=popup_text,
                font=("Segoe UI", 10, "bold"),
            ).pack(anchor="w", padx=12, pady=(10, 8))

            form_row = Frame(popup, bg=popup_bg)
            form_row.pack(fill="x", padx=12, pady=(0, 8))
            Label(form_row, text="Mensaje de error", bg=popup_bg, fg=popup_text).pack(anchor="w")
            ttk.Entry(form_row, textvariable=dropdown_message_var).pack(fill="x", pady=(4, 0))

            def _save_dropdown_message():
                message_text = dropdown_message_var.get().strip()
                if not message_text:
                    messagebox.showerror("Mensajes de error", "El mensaje de error es obligatorio para dropdown.")
                    return
                _set_active_dropdown_error_message(message_text)
                _autosave_mapping_to_json()
                popup.destroy()

            actions_row = Frame(popup, bg=popup_bg)
            actions_row.pack(fill="x", padx=12, pady=(0, 12))
            Button(
                actions_row,
                text="Guardar mensaje",
                command=_save_dropdown_message,
                bg=popup_button_bg,
                fg=popup_button_fg,
                relief="flat",
                font=("Segoe UI", 9, "bold"),
                cursor="hand2",
                padx=10,
                pady=3,
            ).pack(side=LEFT)
            Button(
                actions_row,
                text="Cerrar",
                command=popup.destroy,
                bg=popup_button_bg,
                fg=popup_button_fg,
                relief="flat",
                font=("Segoe UI", 9, "bold"),
                cursor="hand2",
                padx=10,
                pady=3,
            ).pack(side=RIGHT)

            _set_popup_initial_size(popup, min_width=560, min_height=220)
            return

        popup = Toplevel(root_frame)
        popup.title(f"Mensajes de error - {target_id}")
        popup.configure(bg=popup_bg)
        popup.transient(root_frame.winfo_toplevel())
        popup.resizable(True, True)
        # Icono
        icon_path = os.path.join(_get_asset_dir(), "icon.ico")
        if os.path.exists(icon_path):
            try:
                popup.iconbitmap(icon_path)
            except Exception:
                try:
                    popup.iconbitmap(default=icon_path)
                except Exception:
                    pass

        patterns = _get_active_error_message_patterns()
        regex_var = StringVar()
        message_var = StringVar()

        header = Label(
            popup,
            text=f"Configuración de mensajes para ID: {target_id}",
            bg=popup_bg,
            fg=popup_text,
            font=("Segoe UI", 10, "bold"),
        )
        header.pack(anchor="w", padx=12, pady=(10, 8))

        form_row = Frame(popup, bg=popup_bg)
        form_row.pack(fill="x", padx=12, pady=(0, 6))
        Label(form_row, text="Regex", bg=popup_bg, fg=popup_text).grid(row=0, column=0, sticky="w")
        Label(form_row, text="Mensaje", bg=popup_bg, fg=popup_text).grid(row=0, column=1, sticky="w", padx=(10, 0))
        ttk.Entry(form_row, textvariable=regex_var).grid(row=1, column=0, sticky="ew", pady=(2, 0))
        ttk.Entry(form_row, textvariable=message_var).grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=(2, 0))
        form_row.grid_columnconfigure(0, weight=1)
        form_row.grid_columnconfigure(1, weight=2)

        tree_frame = Frame(popup, bg=popup_bg)
        tree_frame.pack(fill="both", expand=True, padx=12, pady=(0, 8))
        msg_tree = ttk.Treeview(
            tree_frame,
            columns=("regex", "mensaje"),
            show="headings",
            height=8,
            style="Section.Treeview",
        )
        msg_tree.heading("regex", text="Regex")
        msg_tree.heading("mensaje", text="Mensaje")
        msg_tree.column("regex", width=300, stretch=True)
        msg_tree.column("mensaje", width=500, stretch=True)
        msg_tree.pack(fill="both", expand=True)

        def _refresh_messages_tree():
            for item_id in msg_tree.get_children():
                msg_tree.delete(item_id)
            for index, rule in enumerate(patterns):
                msg_tree.insert("", END, iid=str(index), values=(rule.get("regex", ""), rule.get("mensaje", "")))
            _apply_row_colors(msg_tree)

        def _load_selected_message(_event=None):
            selected = msg_tree.selection()
            if not selected:
                return
            row_index = int(selected[0])
            if row_index < 0 or row_index >= len(patterns):
                return
            row = patterns[row_index]
            regex_var.set(row.get("regex", ""))
            message_var.set(row.get("mensaje", ""))

        def _upsert_message_rule():
            regex_text = regex_var.get().strip()
            message_text = message_var.get().strip()
            if not regex_text or not message_text:
                messagebox.showerror("Mensajes de error", "Regex y Mensaje son obligatorios.")
                return
            if not _validate_regex_pattern(regex_text, "Regex de mensaje"):
                return

            selected = msg_tree.selection()
            new_rule = {"regex": regex_text, "mensaje": message_text}
            if selected:
                row_index = int(selected[0])
                if 0 <= row_index < len(patterns):
                    patterns[row_index] = new_rule
            else:
                patterns.append(new_rule)

            _set_active_error_message_patterns(patterns)
            _refresh_messages_tree()
            regex_var.set("")
            message_var.set("")

        def _remove_message_rule():
            selected = msg_tree.selection()
            if not selected:
                return
            row_index = int(selected[0])
            if 0 <= row_index < len(patterns):
                patterns.pop(row_index)
                _set_active_error_message_patterns(patterns)
                _refresh_messages_tree()
                regex_var.set("")
                message_var.set("")

        actions_row = Frame(popup, bg=popup_bg)
        actions_row.pack(fill="x", padx=12, pady=(0, 12))
        Button(
            actions_row,
            text="Agregar / Editar mensaje",
            command=_upsert_message_rule,
            bg=popup_button_bg,
            fg=popup_button_fg,
            relief="flat",
            font=("Segoe UI", 9, "bold"),
            cursor="hand2",
            padx=10,
            pady=3,
        ).pack(side=LEFT, padx=(0, 6))
        Button(
            actions_row,
            text="Eliminar mensaje",
            command=_remove_message_rule,
            bg=popup_button_bg,
            fg=popup_button_fg,
            relief="flat",
            font=("Segoe UI", 9, "bold"),
            cursor="hand2",
            padx=10,
            pady=3,
        ).pack(side=LEFT)
        Button(
            actions_row,
            text="Cerrar",
            command=lambda: (_autosave_mapping_to_json(), popup.destroy()),
            bg=popup_button_bg,
            fg=popup_button_fg,
            relief="flat",
            font=("Segoe UI", 9, "bold"),
            cursor="hand2",
            padx=10,
            pady=3,
        ).pack(side=RIGHT)

        msg_tree.bind("<<TreeviewSelect>>", _load_selected_message)
        _refresh_messages_tree()
        _set_popup_initial_size(popup, min_width=780, min_height=440)

    def _mapping_row_to_values(mapping_row):
        return (
            mapping_row["element_id"],
            "SI" if mapping_row.get("dropdown", False) else "",
            mapping_row["descripcion"],
            _format_dependencies(mapping_row.get("dependencies", [])),
            mapping_row["regex_full"],
            mapping_row["regex_char"],
            mapping_row["test_text"],
            _format_rule_countries_abbrev(mapping_row["paises"]),
            "numeric" if mapping_row.get("teclado_mobile") else "",
        )

    def _row_matches_filter(mapping_row):
        filtro_texto = (rule_filter_text_var.get() or "").strip().lower()
        filtro_pais_abrev = (rule_filter_country_var.get() or "Todos").strip()
        filtro_pais = ABBREVIATION_TO_COUNTRY.get(filtro_pais_abrev)

        if filtro_texto:
            searchable = " ".join(
                [
                    mapping_row.get("element_id", ""),
                    "dropdown" if mapping_row.get("dropdown", False) else "",
                    mapping_row.get("descripcion", ""),
                    _format_dependencies(mapping_row.get("dependencies", [])),
                    mapping_row.get("regex_full", ""),
                    mapping_row.get("regex_char", ""),
                    mapping_row.get("test_text", ""),
                    _format_rule_countries_abbrev(mapping_row.get("paises", [])),
                ]
            ).lower()
            if filtro_texto not in searchable:
                return False

        if filtro_pais:
            row_countries = mapping_row.get("paises") or list(AVAILABLE_COUNTRIES)
            row_countries = _parse_rule_countries(row_countries)
            applies_all = set(row_countries) == set(AVAILABLE_COUNTRIES)
            if not applies_all and filtro_pais not in row_countries:
                return False

        return True

    def _refresh_mapping_tree_view():
        for item_id in mapping_tree.get_children():
            mapping_tree.delete(item_id)

        for index, mapping_row in enumerate(mapping_rows):
            if not _row_matches_filter(mapping_row):
                continue
            mapping_tree.insert("", END, iid=str(index), values=_mapping_row_to_values(mapping_row))

        _apply_row_colors(mapping_tree)

    rule_filter_text_var.trace_add("write", lambda *_: _refresh_mapping_tree_view())
    rule_filter_country_var.trace_add("write", lambda *_: _refresh_mapping_tree_view())

    def _collect_mapping_from_tree():
        mapping = {}
        for mapping_row in mapping_rows:
            normalized_paises = _parse_rule_countries(mapping_row.get("paises"))
            normalized_element_id = str(mapping_row.get("element_id") or "").strip()
            normalized_description = str(mapping_row.get("descripcion") or "").strip() or normalized_element_id
            base_rule_key = str(mapping_row.get("rule_key") or normalized_description or normalized_element_id).strip()
            if not base_rule_key:
                base_rule_key = normalized_element_id or "rule"

            rule_key = base_rule_key
            if rule_key in mapping:
                suffix = 2
                while f"{base_rule_key}__{suffix}" in mapping:
                    suffix += 1
                rule_key = f"{base_rule_key}__{suffix}"

            mapping[rule_key] = {
                "descripcion": normalized_description,
                "campo": normalized_element_id,
                "element_id": normalized_element_id,
                "regex_full": str(mapping_row.get("regex_full") or ""),
                "regex_char": str(mapping_row.get("regex_char") or ""),
                "test_text": str(mapping_row.get("test_text") or ""),
                "dropdown": bool(mapping_row.get("dropdown", False)),
                "dropdown_error_message": str(mapping_row.get("dropdown_error_message") or ""),
                "dependencies": deepcopy(_normalize_dependencies(mapping_row.get("dependencies") or [])),
                "paises": normalized_paises,
                "teclado_mobile": bool(mapping_row.get("teclado_mobile", False)),
                "rules": dict(mapping_row.get("rules") or {}),
                "error_messages": dict(mapping_row.get("error_messages") or {}),
                "error_message_patterns": deepcopy(_normalize_error_message_patterns(mapping_row.get("error_message_patterns") or [])),
                "error_config": dict(mapping_row.get("error_config") or {}),
                "error_priority": list(mapping_row.get("error_priority") or ["required", "invalid_chars", "min_length", "max_length"]),
            }
        return mapping

    def _filter_mapping_by_countries(mapping, selected_countries):
        if not selected_countries:
            return {}

        normalized_selected = []
        for country_name in selected_countries:
            normalized = _normalize_selected_country(country_name)
            if normalized and normalized not in normalized_selected:
                normalized_selected.append(normalized)

        if not normalized_selected:
            return {}

        filtered = {}
        for field_name, config in mapping.items():
            field_countries = _parse_rule_countries(config.get("paises"))
            if any(country_name in normalized_selected for country_name in field_countries):
                filtered[field_name] = dict(config)
        return filtered

    def _populate_mapping_tree(field_mapping):
        mapping_rows.clear()

        for field_name, config in field_mapping.items():
            paises = _parse_rule_countries(config.get("paises"))
            descripcion = config.get("descripcion") or field_name
            mapping_rows.append(
                {
                    "rule_key": str(field_name),
                    "element_id": config.get("element_id") or config.get("id") or field_name,
                    "descripcion": descripcion,
                    "regex_full": config.get("regex_full", ""),
                    "regex_char": config.get("regex_char", ""),
                    "test_text": config.get("test_text") or config.get("texto_prueba") or "",
                    "dropdown": bool(config.get("dropdown", False)),
                    "dropdown_error_message": str(config.get("dropdown_error_message") or ""),
                    "dependencies": deepcopy(_normalize_dependencies(config.get("dependencies") or [])),
                    "paises": paises,
                    "teclado_mobile": bool(config.get("teclado_mobile", False)),
                    "rules": dict(config.get("rules") or {}),
                    "error_messages": dict(config.get("error_messages") or {}),
                    "error_message_patterns": deepcopy(_normalize_error_message_patterns(config.get("error_message_patterns") or [])),
                    "error_config": dict(config.get("error_config") or {}),
                    "error_priority": list(config.get("error_priority") or ["required", "invalid_chars", "min_length", "max_length"]),
                }
            )
        _refresh_mapping_tree_view()

    def _load_selected_mapping(_event=None):
        selected = mapping_tree.selection()
        if not selected:
            upsert_button_text_var.set("Agregar regla")
            return

        row_index = int(selected[0])
        if row_index < 0 or row_index >= len(mapping_rows):
            upsert_button_text_var.set("Agregar regla")
            return
        mapping_row = mapping_rows[row_index]
        upsert_button_text_var.set("Editar regla")
        element_id_var.set(mapping_row.get("element_id", ""))
        field_name_var.set(mapping_row.get("descripcion", ""))
        dropdown_var.set(bool(mapping_row.get("dropdown", False)))
        regex_full_var.set(mapping_row.get("regex_full", ""))
        regex_char_var.set(mapping_row.get("regex_char", ""))
        current_dropdown_error_message_var.set(str(mapping_row.get("dropdown_error_message") or ""))
        test_text_widget.delete("1.0", END)
        test_text_widget.insert("1.0", mapping_row.get("test_text", ""))
        _apply_dropdown_mode(clear_values=False)
        _set_rule_countries(mapping_row.get("paises"))
        teclado_mobile_var.set(bool(mapping_row.get("teclado_mobile", False)))
        current_error_message_patterns.clear()
        current_error_message_patterns.extend(deepcopy(_normalize_error_message_patterns(mapping_row.get("error_message_patterns") or [])))
        current_dependencies.clear()
        current_dependencies.extend(deepcopy(_normalize_dependencies(mapping_row.get("dependencies") or [])))

    def _upsert_mapping():
        field_name = field_name_var.get().strip()
        element_id = element_id_var.get().strip()
        is_dropdown = bool(dropdown_var.get())
        regex_full = "" if is_dropdown else regex_full_var.get().strip()
        regex_char = "" if is_dropdown else regex_char_var.get().strip()
        test_text = "" if is_dropdown else test_text_widget.get("1.0", END).strip("\n")
        rule_countries = _collect_rule_countries()
        dropdown_error_message = _get_active_dropdown_error_message()

        if not field_name or not element_id:
            messagebox.showerror("Validación de campos", "Campo e ID son obligatorios.")
            return
        if not is_dropdown:
            if not regex_full or not regex_char:
                messagebox.showerror("Validación de campos", "regex_full y regex_char son obligatorios para campos no dropdown.")
                return
            if not _validate_regex_pattern(regex_full, "Regex full"):
                return
            if not _validate_regex_pattern(regex_char, "Regex char"):
                return

        new_row = {
            "rule_key": "",
            "element_id": element_id,
            "descripcion": field_name,
            "regex_full": regex_full,
            "regex_char": regex_char,
            "test_text": test_text,
            "dropdown": is_dropdown,
            "dropdown_error_message": dropdown_error_message,
            "dependencies": deepcopy(_normalize_dependencies(current_dependencies)),
            "paises": rule_countries,
            "teclado_mobile": bool(teclado_mobile_var.get()),
            "rules": {},
            "error_messages": {},
            "error_message_patterns": deepcopy(_normalize_error_message_patterns(current_error_message_patterns)),
            "error_config": {},
            "error_priority": ["required", "invalid_chars", "min_length", "max_length"],
        }

        selected = mapping_tree.selection()
        if selected:
            row_index = int(selected[0])
            if 0 <= row_index < len(mapping_rows):
                new_row["rule_key"] = str(mapping_rows[row_index].get("rule_key") or "")
                for key in ("rules", "error_messages", "error_message_patterns", "error_config", "error_priority", "dropdown_error_message", "dependencies"):
                    if key in mapping_rows[row_index] and not new_row.get(key):
                        new_row[key] = mapping_rows[row_index].get(key)
                mapping_rows[row_index] = new_row
        else:
            # Nuevo registro: crear una key estable evitando pisar reglas con misma descripción.
            base_key = field_name or element_id or "rule"
            existing_keys = {str(row.get("rule_key") or "") for row in mapping_rows}
            new_key = base_key
            if new_key in existing_keys:
                suffix = 2
                while f"{base_key}__{suffix}" in existing_keys:
                    suffix += 1
                new_key = f"{base_key}__{suffix}"
            new_row["rule_key"] = new_key
            mapping_rows.append(new_row)

        _refresh_mapping_tree_view()
        _clear_form_inputs()
        _autosave_mapping_to_json()

    def _remove_mapping():
        selected = mapping_tree.selection()
        if not selected:
            return
        row_index = int(selected[0])
        if 0 <= row_index < len(mapping_rows):
            mapping_rows.pop(row_index)
        _refresh_mapping_tree_view()
        _clear_form_inputs()
        _autosave_mapping_to_json()

    def _load_rules_into_ui():
        try:
            payload = _load_rules_file()
            landing_urls = payload.get("landing_urls") or []
            form_urls = payload.get("form_urls") or []

            if isinstance(landing_urls, str):
                landing_urls = [landing_urls]
            if isinstance(form_urls, str):
                form_urls = [form_urls]

            if not landing_urls:
                fallback_urls = payload.get("urls") or []
                if isinstance(fallback_urls, str):
                    fallback_urls = [fallback_urls]
                if not fallback_urls:
                    fallback_url = str(payload.get("url", "")).strip()
                    fallback_urls = [fallback_url] if fallback_url else []
                landing_urls = fallback_urls

            if not form_urls and landing_urls:
                form_urls = [""] * len(landing_urls)

            existing_excel_pairs = []
            try:
                existing_excel_pairs = _read_url_pairs_from_excel(require_form=False)
            except Exception:
                existing_excel_pairs = []

            # _set_url_pairs_in_widget espera tripletas (país, landing, form) — el país no
            # se persiste en el JSON de reglas (_save_rules_from_ui solo guarda landing_urls/
            # form_urls), así que se reconstruye con país vacío en vez de perderlo silenciosamente.
            payload_pairs = [("", landing, form) for landing, form in zip(landing_urls, form_urls)]
            if existing_excel_pairs:
                _reload_urls_preview_from_excel()
            elif payload_pairs:
                _set_url_pairs_in_widget(payload_pairs)
            else:
                _create_validation_urls_excel_if_missing()
                _reload_urls_preview_from_excel()
            email_var.set(", ".join(obtener_email_destinatario()))

            cfg = cargar_config_global()
            enviar_mail_var.set(bool(cfg.get("enviar_mail", False)))
            adjuntar_resultados_var.set(bool(cfg.get("adjuntar_resultados", True)))

            selected_browsers = payload.get("selected_browsers") or []
            if isinstance(selected_browsers, str):
                selected_browsers = [selected_browsers]
            if not selected_browsers:
                fallback_browser = str(payload.get("browser", "chrome")).strip().lower()
                selected_browsers = [fallback_browser] if fallback_browser in BROWSER_OPTIONS else ["chrome"]
            for browser_name in BROWSER_OPTIONS:
                browser_vars[browser_name].set(browser_name in selected_browsers)

            selected_viewports = payload.get("selected_viewports") or []
            if isinstance(selected_viewports, str):
                selected_viewports = [selected_viewports]
            if not selected_viewports:
                fallback_viewport = str(payload.get("viewport", "fullscreen")).strip().lower()
                selected_viewports = [fallback_viewport] if fallback_viewport in VIEWPORT_OPTIONS else ["fullscreen"]
            for viewport_name in VIEWPORT_OPTIONS:
                viewport_vars[viewport_name].set(viewport_name in selected_viewports)

            _populate_mapping_tree(payload.get("fields", {}))
        except Exception as exc:
            LOGGER.exception("No se pudieron cargar las reglas")
            messagebox.showerror(
                "Validación de campos",
                f"No se pudieron cargar las reglas. {_format_user_error(exc)}",
            )

    def _autosave_mapping_to_json():
        """Guarda automáticamente el mapping actual al JSON sin validar URLs."""
        try:
            try:
                payload = _load_rules_file()
            except Exception:
                payload = {
                    "urls": [],
                    "landing_urls": [],
                    "form_urls": [],
                    "browser": "chrome",
                    "viewport": "fullscreen",
                    "selected_browsers": ["chrome"],
                    "selected_viewports": ["fullscreen"],
                    "fields": {}
                }
            
            payload["fields"] = _collect_mapping_from_tree()
            _save_rules_file(payload)
            status_var.set(f"Regla guardada automáticamente en {os.path.basename(_get_rules_path())}.")
        except Exception as exc:
            LOGGER.warning("Error al guardar reglas automáticamente: %s", exc)
            status_var.set("No se pudo guardar automáticamente la regla.")

    def _save_rules_from_ui(require_form_urls=True):
        try:
            if require_form_urls:
                url_pairs = _collect_url_pairs_from_widget()
            else:
                url_pairs = _read_url_pairs_from_excel(require_form=False)
        except ValueError as exc:
            if require_form_urls:
                messagebox.showerror("Validación de campos", _format_user_error(exc))
            return

        landing_urls = [pair[0] for pair in url_pairs]
        form_urls = [pair[1] for pair in url_pairs]
        selected_browsers = _collect_selected_browsers()
        selected_viewports = _collect_selected_viewports()

        payload = {
            "url": landing_urls[0] if landing_urls else "",
            "urls": landing_urls,
            "landing_urls": landing_urls,
            "form_urls": form_urls,
            "browser": selected_browsers[0] if selected_browsers else "chrome",
            "viewport": selected_viewports[0] if selected_viewports else "fullscreen",
            "selected_browsers": selected_browsers,
            "selected_viewports": selected_viewports,
            "fields": _collect_mapping_from_tree(),
        }
        try:
            _save_rules_file(payload)
            status_var.set(f"Reglas guardadas en {os.path.basename(_get_rules_path())}.")
        except Exception as exc:
            LOGGER.exception("No se pudieron guardar las reglas")
            messagebox.showerror(
                "Validación de campos",
                f"No se pudieron guardar las reglas. {_format_user_error(exc)}",
            )

    def _run_validation_async():
        mapping = _collect_mapping_from_tree()
        try:
            url_triples = _collect_url_pairs_from_widget()
        except ValueError as exc:
            messagebox.showerror("Validación de campos", _format_user_error(exc))
            return

        if not url_triples:
            messagebox.showerror("Validación de campos", "Ingresá al menos una fila con País, URL y Formulario.")
            return
        if any(not triple[2] for triple in url_triples):
            messagebox.showerror("Validación de campos", "Todas las filas deben tener Formulario.")
            return
        if not mapping:
            messagebox.showerror("Validación de campos", "Agregá al menos un campo al mapping.")
            return

        selected_browsers = _collect_selected_browsers()
        selected_viewports = _collect_selected_viewports()
        if not selected_browsers:
            messagebox.showerror("Validación de campos", "Seleccioná al menos un browser.")
            return
        if not selected_viewports:
            messagebox.showerror("Validación de campos", "Seleccioná al menos un viewport.")
            return

        total_runs = len(url_triples) * len(selected_browsers) * len(selected_viewports)
        status_var.set(
            f"Ejecutando {total_runs} validación(es): {len(url_triples)} URL(s), {len(selected_browsers)} browser(s), {len(selected_viewports)} viewport(s)."
        )

        def _worker():
            all_rows = []
            all_fields = []
            all_error_rows = []
            all_unmapped_ids = []
            aggregated_summary = {
                "url_pairs": len(url_triples),
                "urls_ok": 0,
                "urls_error": 0,
                "fields": 0,
                "characters": 0,
                "errors": 0,
                "ok": 0,
                "ui_error_tests": 0,
                "ui_errors": 0,
                "ui_ok": 0,
                "ids_no_mapeados": 0,
                "regex_ok": True,
                # Detalle por URL, para que el email pueda listar landing/form y el motivo
                # del fallo en vez de solo un contador.
                "url_details": [],
            }
            execution_errors = []
            selected_countries_from_urls = list(set(triple[0] for triple in url_triples))

            try:
                run_index = 0
                # Se crea un navegador por cada combinación browser/viewport y se
                # reutiliza para todas las URLs del listado (no hace falta cerrar y
                # abrir entre formularios — solo se navega a la siguiente URL).
                for browser_name in selected_browsers:
                    for viewport_name in selected_viewports:
                        shared_driver = BrowserManager.create_browser(
                            browser_type=browser_name,
                            viewport=viewport_name,
                            headless=not ver_navegador_var.get(),
                            background=not ver_navegador_var.get(),
                        )
                        try:
                            for pais_abrev, landing_url, form_url in url_triples:
                                run_index += 1
                                root_frame.after(
                                    0,
                                    lambda i=run_index, total=total_runs, pa=pais_abrev, lu=landing_url, br=browser_name, vp=viewport_name: status_var.set(
                                        f"Validando {i}/{total}: {pa} | {br} | {vp} | {lu}"
                                    ),
                                )
                                # Filter mapping by the country of this URL triple
                                pais_completo = _normalize_selected_country(pais_abrev)
                                filtered_mapping = _filter_mapping_by_countries(mapping, [pais_completo])
                                try:
                                    if not filtered_mapping:
                                        raise ValueError(f"No hay reglas aplicables para el país {pais_abrev.upper()}.")
                                    result = run_field_validations(
                                        landing_url=landing_url,
                                        expected_form_url=form_url,
                                        field_mapping=filtered_mapping,
                                        browser=browser_name,
                                        viewport=viewport_name,
                                        driver=shared_driver,
                                    )
                                    rows = list(result.get("rows") or [])
                                    fields = list(result.get("fields") or [])
                                    error_rows = list(result.get("error_rows") or [])
                                    unmapped_ids = list(result.get("unmapped_ids") or [])
                                    summary = dict(result.get("summary") or {})

                                    all_rows.extend(rows)
                                    all_fields.extend(fields)
                                    all_error_rows.extend(error_rows)
                                    all_unmapped_ids.extend(unmapped_ids)
                                    aggregated_summary["urls_ok"] += 1
                                    _errs_url = int(summary.get("errors", 0)) + int(summary.get("ui_errors", 0))
                                    aggregated_summary["url_details"].append({
                                        "pais": pais_abrev.upper(),
                                        "browser": browser_name,
                                        "viewport": viewport_name,
                                        "landing": landing_url or "",
                                        "form": form_url or "",
                                        "ok": _errs_url == 0,
                                        "error": ("" if _errs_url == 0
                                                  else f"{_errs_url} validacion(es) fallida(s) — ver el Excel adjunto"),
                                    })
                                    aggregated_summary["fields"] += int(summary.get("fields", 0))
                                    aggregated_summary["characters"] += int(summary.get("characters", 0))
                                    aggregated_summary["errors"] += int(summary.get("errors", 0))
                                    aggregated_summary["ok"] += int(summary.get("ok", 0))
                                    aggregated_summary["ui_error_tests"] += int(summary.get("ui_error_tests", 0))
                                    aggregated_summary["ui_errors"] += int(summary.get("ui_errors", 0))
                                    aggregated_summary["ui_ok"] += int(summary.get("ui_ok", 0))
                                    aggregated_summary["ids_no_mapeados"] += int(summary.get("ids_no_mapeados", 0))
                                    aggregated_summary["regex_ok"] = aggregated_summary["regex_ok"] and bool(summary.get("regex_ok", False))
                                except Exception as url_exc:
                                    aggregated_summary["urls_error"] += 1
                                    aggregated_summary["regex_ok"] = False
                                    aggregated_summary["url_details"].append({
                                        "pais": pais_abrev.upper(),
                                        "browser": browser_name,
                                        "viewport": viewport_name,
                                        "landing": landing_url or "",
                                        "form": form_url or "",
                                        "ok": False,
                                        "error": str(url_exc)[:300],
                                    })
                                    LOGGER.warning(
                                        "Falla de validación país=%s browser=%s viewport=%s landing=%s form=%s",
                                        pais_abrev,
                                        browser_name,
                                        viewport_name,
                                        landing_url,
                                        form_url,
                                    )
                                    execution_errors.append(
                                        f"{pais_abrev} | {browser_name}/{viewport_name} | {landing_url} -> {_format_user_error(url_exc)}"
                                    )
                        finally:
                            try:
                                shared_driver.quit()
                            except Exception:
                                pass

                root_frame.after(
                    0,
                    lambda: _apply_run_result(
                        {
                            "rows": all_rows,
                            "fields": all_fields,
                            "error_rows": all_error_rows,
                            "unmapped_ids": all_unmapped_ids,
                            "summary": aggregated_summary,
                            "selected_paises": [ABBREVIATION_TO_COUNTRY.get(p.upper(), p) for p in selected_countries_from_urls],
                        },
                        execution_errors,
                    ),
                )
            except Exception as exc:
                LOGGER.exception("Error ejecutando la validacion de campos")
                root_frame.after(
                    0,
                    lambda: messagebox.showerror(
                        "Validación de campos",
                        "La validación terminó con error. "
                        + _format_user_error(exc, fallback="Revisá la configuración e intentá nuevamente."),
                    ),
                )
                root_frame.after(0, lambda: status_var.set("La validación terminó con error."))

        threading.Thread(target=_worker, daemon=True).start()

    def _open_results_dir():
        results_dir = _get_results_dir()
        os.makedirs(results_dir, exist_ok=True)
        try:
            os.startfile(results_dir)
        except AttributeError:
            messagebox.showinfo("Resultados", f"La carpeta de resultados está en:\n{results_dir}")
        except OSError as exc:
            messagebox.showerror("Resultados", f"No se pudo abrir la carpeta de resultados: {exc}")

    def _apply_run_result(result, execution_errors=None):
        state["rows"] = list(result.get("rows") or [])
        state["fields"] = list(result.get("fields") or [])
        state["error_rows"] = list(result.get("error_rows") or [])
        state["unmapped_ids"] = list(result.get("unmapped_ids") or [])
        state["summary"] = dict(result.get("summary") or {})
        state["export_path"] = None
        summary = state["summary"]
        status_message = (
            "Paises: {paises} | URLs OK/Error: {urls_ok}/{urls_error} | Campos: {fields} | Caracteres: {characters} | OK: {ok} | Errores: {errors}".format(
                paises=", ".join(result.get("selected_paises") or []),
                urls_ok=summary.get("urls_ok", 0),
                urls_error=summary.get("urls_error", 0),
                fields=summary.get("fields", 0),
                characters=summary.get("characters", 0),
                ok=summary.get("ok", 0),
                errors=summary.get("errors", 0),
            )
        )
        try:
            export_path = export_validation_results(
                {
                    "rows": state["rows"],
                    "fields": state["fields"],
                    "error_rows": state["error_rows"],
                    "unmapped_ids": state["unmapped_ids"],
                    "summary": state["summary"],
                }
            )
            state["export_path"] = export_path
            status_var.set(f"{status_message} | Excel: {os.path.basename(export_path)}")

            # "Enviar mail" manda el mail; "Adjuntar resultados" decide solo si va el Excel
            # adjunto. Antes, con adjuntar apagado, NO se enviaba nada y el usuario se
            # quedaba sin el reporte, que es la parte que importa.
            if enviar_mail_var.get():
                ok, msg = send_validation_report_email(
                    state["export_path"],
                    state["summary"],
                    recipient=(email_var.get() or "").strip(),
                    adjuntar_excel=bool(adjuntar_resultados_var.get()),
                )
                if not ok:
                    messagebox.showerror("Validación de campos", _format_user_error(msg))
        except Exception as exc:
            status_var.set(status_message)
            LOGGER.exception("No se pudo generar el Excel automático")
            messagebox.showerror(
                "Validación de campos",
                "No se pudo generar el Excel automático. " + _format_user_error(exc),
            )

        # Eliminado: cartel de aviso al finalizar el testeo de validación de campos

    id_button_bg = app_bg
    id_button_fg = "white"

    # Distribuir los botones equitativamente
    mapping_buttons.grid_columnconfigure(0, weight=1)
    mapping_buttons.grid_columnconfigure(1, weight=1)
    mapping_buttons.grid_columnconfigure(2, weight=1)
    mapping_buttons.grid_columnconfigure(3, weight=1)
    btn_error = Button(
        mapping_buttons,
        text="Mensaje de error",
        command=_open_error_messages_popup,
        bg=id_button_bg,
        fg=id_button_fg,
        relief="flat",
        font=("Segoe UI", 10, "bold"),
        cursor="hand2",
        padx=10,
        pady=3,
    )
    btn_error.grid(row=0, column=0, sticky="ew", padx=6)
    btn_dep = Button(
        mapping_buttons,
        text="Dependencia",
        command=_open_dependencies_popup,
        bg=id_button_bg,
        fg=id_button_fg,
        relief="flat",
        font=("Segoe UI", 10, "bold"),
        cursor="hand2",
        padx=10,
        pady=3,
    )
    btn_dep.grid(row=0, column=1, sticky="ew", padx=6)
    btn_upsert = Button(
        mapping_buttons,
        textvariable=upsert_button_text_var,
        command=_upsert_mapping,
        bg=id_button_bg,
        fg=id_button_fg,
        relief="flat",
        font=("Segoe UI", 10, "bold"),
        cursor="hand2",
        padx=10,
        pady=3,
    )
    btn_upsert.grid(row=0, column=2, sticky="ew", padx=6)
    btn_del = Button(
        mapping_buttons,
        text="Eliminar regla",
        command=_remove_mapping,
        bg=id_button_bg,
        fg=id_button_fg,
        relief="flat",
        font=("Segoe UI", 10, "bold"),
        cursor="hand2",
        padx=10,
        pady=3,
    )
    btn_del.grid(row=0, column=3, sticky="ew", padx=6)

    Button(
        actions_frame,
        text="Resultados",
        command=_open_results_dir,
        bg=button_bg,
        fg=button_fg,
        relief="flat",
        font=("Segoe UI", 10, "bold"),
        cursor="hand2",
        padx=10,
        pady=3,
    ).pack(side=RIGHT, padx=(6, 0))
    Button(
        actions_frame,
        text="Ejecutar validación",
        command=_run_validation_async,
        bg=button_bg,
        fg=button_fg,
        relief="flat",
        font=("Segoe UI", 10, "bold"),
        cursor="hand2",
        padx=10,
        pady=3,
    ).pack(side=RIGHT, padx=(6, 0))

    mapping_tree.bind("<<TreeviewSelect>>", _load_selected_mapping)
    _load_rules_into_ui()
    return root_frame