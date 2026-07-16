# -*- coding: utf-8 -*-
"""
ids_dinamicos_ui.py — Popup "IDs Dinámicos" reutilizable (IDs únicos, IDs Excel y Dependencias).

Extraído desde main_interface_old.py para poder engancharlo también en el refactor nuevo
(main_interface.py) sin duplicar lógica ni tocar el backend. Permite dar de alta campos
no mapeados desde la interfaz en lugar de editar el JSON a mano.
"""
import os
import json
import textwrap
from tkinter import *
from tkinter import ttk, messagebox
from tkinter import StringVar, BooleanVar

from .helpers_interface import (
    cargar_ids_dinamicos,
    guardar_ids_dinamicos,
    cargar_dependencias,
    guardar_dependencias,
    obtener_ids_mapeados_normales,
    sincronizar_excels_de_pais,
)
from utils.fixed_field_mapping_store import (
    list_available_fixed_mapping_countries,
    load_effective_country_form_config,
    save_country_fixed_field_mapping,
)

try:
    from utils.paths import ASSET_DIR, JSON_DIR
except Exception:
    ASSET_DIR = ""
    JSON_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "json")

APP_BG_COLOR = "#5D3C7A"
HEADER_BG_COLOR = "#9c6fb4"

MAPEO_PAISES = {
    "Argentina": "Formulario_Argentina_Main",
    "Bolivia": "Formulario_Bolivia_Main",
    "Brasil": "Formulario_Brasil_Main",
    "Chile": "Formulario_Chile_Main",
    "Colombia": "Formulario_Colombia_Main",
    "Ecuador": "Formulario_Ecuador_Main",
    "Paraguay": "Formulario_Paraguay_Main",
    "Peru": "Formulario_Peru_Main",
    "Uruguay": "Formulario_Uruguay_Main",
}


# === Helpers de normalización (copiados de main_interface_old.py) ===
def normalizar_valores_id_dinamico(valor):
    """Normaliza un valor de ID dinámico a una lista sin vacíos."""
    if valor is None:
        return []
    if isinstance(valor, list):
        return [str(item).strip() for item in valor if str(item).strip()]
    texto = str(valor).strip()
    if not texto:
        return []
    if "|" in texto:
        return [item.strip() for item in texto.split("|") if item.strip()]
    return [texto]


def extraer_datos_id_dinamico(raw_value):
    """Obtiene nombre de campo y valor desde formato legacy o nuevo."""
    if isinstance(raw_value, dict):
        nombre = str(
            raw_value.get("nombre_campo")
            or raw_value.get("nombre")
            or raw_value.get("campo")
            or ""
        ).strip()
        valor = (
            raw_value.get("valor")
            if "valor" in raw_value
            else raw_value.get("valores", raw_value.get("value", raw_value.get("values")))
        )
        return nombre, valor
    return "", raw_value


def compactar_valores_id_dinamico(valores):
    """Devuelve un valor escalar si hay uno solo; si no, conserva lista."""
    return valores if len(valores) > 1 else valores[0]


def normalizar_paises_id_dinamico(raw_paises):
    """Normaliza lista de países de un ID dinámico."""
    if raw_paises is None:
        return []
    if isinstance(raw_paises, str):
        candidatos = [raw_paises]
    elif isinstance(raw_paises, (list, tuple, set)):
        candidatos = list(raw_paises)
    else:
        return []

    paises = []
    for pais in candidatos:
        texto = str(pais).strip()
        if texto and texto not in paises:
            paises.append(texto)
    return paises


def _autohide_yscroll(sb, first, last):
    """Oculta el scrollbar cuando todo el contenido cabe; lo muestra cuando hay overflow."""
    first, last = float(first), float(last)
    if not hasattr(sb, "_ah_manager"):
        mgr = sb.winfo_manager()
        if mgr:
            sb._ah_manager = mgr
    mgr = getattr(sb, "_ah_manager", "pack")
    if first <= 0.0 and last >= 1.0:
        if mgr == "grid":
            try: sb.grid_remove()
            except Exception: pass
        else:
            try: sb.pack_forget()
            except Exception: pass
    else:
        if not sb.winfo_ismapped():
            if mgr == "grid":
                try: sb.grid()
                except Exception: pass
            else:
                try: sb.pack(side="right", fill="y")
                except Exception: pass
        sb.set(first, last)


def _columna_excel_desde_data_index(data_index):
    # A=URL, B=Formulario → datos desde C. data_index=0 → columna 3 → offset +3
    try:
        excel_column = int(data_index) + 3
    except Exception:
        return ""

    resultado = ""
    while excel_column > 0:
        excel_column, remainder = divmod(excel_column - 1, 26)
        resultado = chr(65 + remainder) + resultado
    return resultado


def _parsear_id_fijo_input(raw_value):
    texto = str(raw_value or "").strip()
    if not texto:
        return ""

    partes = [parte.strip() for parte in texto.split("|") if parte.strip()]
    if not partes:
        return ""
    return partes if len(partes) > 1 else partes[0]


def _texto_id_fijo(raw_value):
    if isinstance(raw_value, list):
        return " | ".join(str(item).strip() for item in raw_value if str(item).strip())
    return str(raw_value or "").strip()


def _ensure_styles():
    """Crea los estilos ttk custom que el refactor nuevo puede no tener definidos."""
    style = ttk.Style()
    try:
        style.layout("Section.Vertical.TScrollbar")
    except Exception:
        try:
            style.layout("Section.Vertical.TScrollbar", style.layout("Vertical.TScrollbar"))
        except Exception:
            pass


# === IDs Excel ===
def _build_tab_ids_excel(popup):

    Label(
        popup,
        text="IDs Excel",
        font=("Segoe UI", 12, "bold"),
        bg=APP_BG_COLOR,
        fg="white",
    ).pack(pady=(16, 8), padx=20, anchor="w")

    descripcion_popup = (
        "Estos IDs corresponden al mapping fijo por país. Elegí un país y ajustá ID, descripción, tipo y columna Excel. "
        "Las columnas A y B siguen fijas para URL y Formulario; desde C el data_index arranca en 0."
    )
    Label(
        popup,
        text=descripcion_popup,
        font=("Segoe UI", 9),
        bg=APP_BG_COLOR,
        fg="#ddd",
        justify="left",
        anchor="w",
        wraplength=700,
    ).pack(fill="x", padx=20, anchor="w")

    ttk.Separator(popup, orient="horizontal").pack(fill="x", padx=20, pady=10)

    paises_disponibles = list_available_fixed_mapping_countries() or list(MAPEO_PAISES.keys())
    pais_var = StringVar(value=paises_disponibles[0] if paises_disponibles else "")
    tipo_var = StringVar(value="Rellenable")
    _TIPO_DISPLAY = {"text": "Rellenable", "select": "Dropdown"}
    _TIPO_STORAGE = {"Rellenable": "text", "Dropdown": "select"}
    id_var = StringVar(value="")
    descripcion_var = StringVar(value="")
    data_index_var = StringVar(value="0")
    columna_excel_var = StringVar(value="Columna Excel: D (data_index 0)")
    btn_guardar_text = StringVar(value="Crear")

    frame_inputs = Frame(popup, bg=APP_BG_COLOR)
    frame_inputs.pack(padx=20, fill="x")

    _estilo_ids = ttk.Style()
    _estilo_ids.configure("IDs.TCombobox", fieldbackground="white", background="white", foreground="black")
    _estilo_ids.map("IDs.TCombobox", fieldbackground=[("readonly", "white")], foreground=[("readonly", "black")])

    Label(frame_inputs, text="País:", bg=APP_BG_COLOR, fg="white", font=("Segoe UI", 10), width=14, anchor="w").grid(row=0, column=0, sticky="w", pady=4)
    combo_pais = ttk.Combobox(frame_inputs, textvariable=pais_var, values=paises_disponibles, state="readonly", width=24, style="IDs.TCombobox")
    combo_pais.grid(row=0, column=1, sticky="w", pady=4, padx=(0, 12))

    Label(frame_inputs, text="Tipo:", bg=APP_BG_COLOR, fg="white", font=("Segoe UI", 10), width=14, anchor="w").grid(row=1, column=0, sticky="w", pady=4)
    combo_tipo = ttk.Combobox(frame_inputs, textvariable=tipo_var, values=["Rellenable", "Dropdown"], state="readonly", width=12, style="IDs.TCombobox")
    combo_tipo.grid(row=1, column=1, sticky="w", pady=4, padx=(0, 12))

    Label(frame_inputs, text="ID:", bg=APP_BG_COLOR, fg="white", font=("Segoe UI", 10), width=14, anchor="w").grid(row=2, column=0, sticky="w", pady=4)
    Entry(frame_inputs, font=("Segoe UI", 10), width=30, textvariable=id_var).grid(row=2, column=1, sticky="w", pady=4, padx=(0, 12))

    Label(frame_inputs, text="Descripción:", bg=APP_BG_COLOR, fg="white", font=("Segoe UI", 10), width=14, anchor="w").grid(row=3, column=0, sticky="w", pady=4)
    Entry(frame_inputs, font=("Segoe UI", 10), width=30, textvariable=descripcion_var).grid(row=3, column=1, sticky="w", pady=4, padx=(0, 12))

    Label(frame_inputs, text="Data index:", bg=APP_BG_COLOR, fg="white", font=("Segoe UI", 10), width=14, anchor="w").grid(row=4, column=0, sticky="w", pady=4)
    Entry(frame_inputs, font=("Segoe UI", 10), width=12, textvariable=data_index_var).grid(row=4, column=1, sticky="w", pady=4, padx=(0, 12))

    Label(frame_inputs, textvariable=columna_excel_var, bg=APP_BG_COLOR, fg="#ffd38a", font=("Segoe UI", 9, "bold"), anchor="w", justify="left").grid(row=5, column=0, columnspan=2, sticky="w", pady=(2, 0))

    frame_ctas_ids_fijos = Frame(frame_inputs, bg=APP_BG_COLOR)
    frame_ctas_ids_fijos.grid(row=0, column=2, rowspan=6, padx=(18, 0), sticky="ne")

    estado_edicion = Label(
        popup,
        text="",
        bg=APP_BG_COLOR,
        fg="#ffd38a",
        font=("Segoe UI", 9, "bold"),
        anchor="w",
        justify="left",
    )
    estado_edicion.pack(fill="x", padx=20, pady=(6, 0), anchor="w")

    ttk.Separator(popup, orient="horizontal").pack(fill="x", padx=20, pady=(10, 6))
    Label(popup, text="Mapping actual:", font=("Segoe UI", 10, "bold"), bg=APP_BG_COLOR, fg="white").pack(padx=20, anchor="w")

    frame_lista_container = Frame(popup, bg=APP_BG_COLOR)
    frame_lista_container.pack(padx=20, pady=(6, 14), fill="both", expand=True)

    canvas_lista = Canvas(frame_lista_container, bg=APP_BG_COLOR, highlightthickness=0, bd=0)
    scroll_lista = ttk.Scrollbar(frame_lista_container, orient="vertical", style="Section.Vertical.TScrollbar", command=canvas_lista.yview)
    canvas_lista.configure(yscrollcommand=lambda f, l, _sb=scroll_lista: _autohide_yscroll(_sb, f, l))
    canvas_lista.pack(side=LEFT, fill="both", expand=True)
    scroll_lista.pack(side=RIGHT, fill="y")

    frame_lista = Frame(canvas_lista, bg=APP_BG_COLOR)
    lista_window = canvas_lista.create_window((0, 0), window=frame_lista, anchor="nw")
    popup._scroll_canvas = canvas_lista  # para scroll con ruedita a nivel popup
    frame_lista.bind("<Configure>", lambda _event=None: canvas_lista.configure(scrollregion=canvas_lista.bbox("all")))
    canvas_lista.bind("<Configure>", lambda event: canvas_lista.itemconfigure(lista_window, width=event.width))

    filas_widgets = {}
    entrada_en_edicion = {"index": None}
    entradas_pais = []
    required_ids_pais = []

    def _actualizar_columna_excel(*_):
        try:
            data_index = int(data_index_var.get().strip())
        except Exception:
            columna_excel_var.set("Columna Excel: índice inválido")
            return

        columna_excel = _columna_excel_desde_data_index(data_index)
        columna_excel_var.set(f"Columna Excel: {columna_excel} (data_index {data_index})")

    def _limpiar_formulario(reset_edicion=True):
        id_var.set("")
        descripcion_var.set("")
        tipo_var.set("Rellenable")
        data_index_var.set("0")
        if reset_edicion:
            entrada_en_edicion["index"] = None
            btn_guardar_text.set("Crear")
            estado_edicion.config(text="")
        _actualizar_columna_excel()

    def _reindexar_entradas_para_excel(entries):
        """Ordena por índice solicitado y reindexa sin huecos para Excel."""
        if not entries:
            return []

        def _as_int(value, default=0):
            try:
                return int(value)
            except Exception:
                return default

        ordenadas = sorted(
            [dict(entry) for entry in entries],
            key=lambda entry: (
                _as_int(entry.get("requested_data_index", entry.get("data_index", 0))),
                str(entry.get("name") or "").lower(),
                str(entry.get("id") or "").lower(),
            ),
        )

        start_index = _as_int(ordenadas[0].get("data_index", 0), 0)
        for offset, entry in enumerate(ordenadas):
            requested = _as_int(entry.get("requested_data_index", entry.get("data_index", start_index + offset)), start_index + offset)
            entry["requested_data_index"] = requested
            entry["data_index"] = start_index + offset

        return ordenadas

    def _cargar_pais_actual():
        nonlocal entradas_pais, required_ids_pais
        config_pais = load_effective_country_form_config(pais_var.get().strip())
        entradas_pais = [dict(entry) for entry in (config_pais.get("field_mapping") or [])]
        for entry in entradas_pais:
            try:
                entry["requested_data_index"] = int(entry.get("requested_data_index", entry.get("data_index", 0)))
            except Exception:
                entry["requested_data_index"] = int(entry.get("data_index", 0)) if str(entry.get("data_index", "")).strip().isdigit() else 0
        entradas_pais = _reindexar_entradas_para_excel(entradas_pais)
        required_ids_pais = list(config_pais.get("country_fields", {}).get("required_fields") or [])

    def _refrescar_lista():
        for widget in list(filas_widgets.values()):
            widget.destroy()
        filas_widgets.clear()

        if not entradas_pais:
            lbl = Label(frame_lista, text="(sin mapping)", bg=APP_BG_COLOR, fg="#aaa", font=("Segoe UI", 9, "italic"))
            lbl.pack(anchor="w")
            filas_widgets["__empty__"] = lbl
            return

        for idx, entry in enumerate(entradas_pais):
            fila = Frame(frame_lista, bg=APP_BG_COLOR)
            fila.pack(fill="x", pady=2)

            data_index = int(entry.get("data_index", 0))
            columna_excel = _columna_excel_desde_data_index(data_index)
            ids_actuales = entry.get("id") if isinstance(entry.get("id"), list) else [entry.get("id")]
            es_requerido = any(str(raw_id or "").strip() in required_ids_pais for raw_id in ids_actuales)
            sufijo_requerido = " | requerido" if es_requerido else ""
            texto_fila = (
                f"{columna_excel} (index {data_index}) | ID: {_texto_id_fijo(entry.get('id'))} | "
                f"{str(entry.get('name') or '').strip()} | {_TIPO_DISPLAY.get(str(entry.get('type') or 'text').strip(), str(entry.get('type') or 'text').strip())}{sufijo_requerido}"
            )

            Label(
                fila,
                text=texto_fila,
                bg=APP_BG_COLOR,
                fg="white",
                font=("Segoe UI", 10),
                anchor="w",
                justify="left",
                wraplength=660,
            ).pack(side=LEFT, expand=True, fill="x")

            def _editar_existente(index=idx):
                if index < 0 or index >= len(entradas_pais):
                    return
                entry_actual = entradas_pais[index]
                id_var.set(_texto_id_fijo(entry_actual.get("id")))
                descripcion_var.set(str(entry_actual.get("name") or "").strip())
                tipo_var.set(_TIPO_DISPLAY.get(str(entry_actual.get("type") or "text").strip(), "Rellenable"))
                data_index_var.set(str(entry_actual.get("requested_data_index", entry_actual.get("data_index", 0))))
                entrada_en_edicion["index"] = index
                btn_guardar_text.set("Editar")
                estado_edicion.config(text=f"Editando {pais_var.get().strip()}: {_texto_id_fijo(entry_actual.get('id'))}")
                _actualizar_columna_excel()
                canvas_lista.yview_moveto(0)

            def _eliminar(index=idx):
                if index < 0 or index >= len(entradas_pais):
                    return
                entradas_pais.pop(index)
                reindexadas = _reindexar_entradas_para_excel(entradas_pais)
                entradas_pais[:] = reindexadas
                pais_nombre = pais_var.get().strip()
                save_country_fixed_field_mapping(pais_nombre, entradas_pais, required_fields=required_ids_pais)
                sincronizar_excels_de_pais(pais_nombre)
                if entrada_en_edicion["index"] == index:
                    _limpiar_formulario()
                elif isinstance(entrada_en_edicion["index"], int) and entrada_en_edicion["index"] > index:
                    entrada_en_edicion["index"] -= 1
                _cargar_pais_actual()
                _refrescar_lista()

            Button(fila, text="Editar", command=_editar_existente, bg=HEADER_BG_COLOR, fg="black", relief="flat", font=("Segoe UI", 9, "bold"), cursor="hand2", padx=6, pady=1).pack(side=RIGHT, padx=(4, 0))
            Button(fila, text="✕", command=_eliminar, bg="#7a2040", fg="white", relief="flat", font=("Segoe UI", 9, "bold"), cursor="hand2", padx=6, pady=1).pack(side=RIGHT)
            filas_widgets[f"row_{idx}"] = fila

    def _guardar_mapping_fijo():
        pais_nombre = pais_var.get().strip()
        if not pais_nombre:
            messagebox.showwarning("IDs Excel", "Seleccioná un país.", parent=popup)
            return

        id_texto = id_var.get().strip()
        if not id_texto:
            messagebox.showwarning("IDs Excel", "Ingresá un ID.", parent=popup)
            return

        try:
            data_index = int(data_index_var.get().strip())
        except Exception:
            messagebox.showwarning("IDs Excel", "El data_index debe ser numérico.", parent=popup)
            return

        if data_index < 0:
            messagebox.showwarning("IDs Excel", "El data_index debe ser mayor o igual a 0.", parent=popup)
            return

        nueva_entry = {
            "id": _parsear_id_fijo_input(id_texto),
            "name": descripcion_var.get().strip() or id_texto,
            "type": _TIPO_STORAGE.get(tipo_var.get().strip(), "text"),
            "requested_data_index": data_index,
            "data_index": data_index,
        }

        idx_edit = entrada_en_edicion["index"]
        if isinstance(idx_edit, int) and 0 <= idx_edit < len(entradas_pais):
            entradas_pais[idx_edit] = nueva_entry
        else:
            entradas_pais.append(nueva_entry)

        reindexadas = _reindexar_entradas_para_excel(entradas_pais)
        entradas_pais[:] = reindexadas

        save_country_fixed_field_mapping(pais_nombre, entradas_pais, required_fields=required_ids_pais)
        sincronizar_excels_de_pais(pais_nombre)
        _cargar_pais_actual()
        _limpiar_formulario()
        _refrescar_lista()

    Button(frame_ctas_ids_fijos, text="Limpiar", command=_limpiar_formulario, bg=HEADER_BG_COLOR, fg="black", relief="flat", font=("Segoe UI", 10, "bold"), cursor="hand2", width=12, padx=10, pady=3).pack(anchor="e", pady=(0, 6))
    Button(frame_ctas_ids_fijos, textvariable=btn_guardar_text, command=_guardar_mapping_fijo, bg=HEADER_BG_COLOR, fg="black", relief="flat", font=("Segoe UI", 10, "bold"), cursor="hand2", width=12, padx=10, pady=3).pack(anchor="e")

    combo_pais.bind("<<ComboboxSelected>>", lambda _event: (_cargar_pais_actual(), _limpiar_formulario(), _refrescar_lista()))
    combo_tipo.set("Rellenable")
    entry_data_index_widget = frame_inputs.grid_slaves(row=4, column=1)
    if entry_data_index_widget:
        entry_data_index_widget[0].bind("<KeyRelease>", _actualizar_columna_excel)

    _cargar_pais_actual()
    _limpiar_formulario()
    _refrescar_lista()


# === IDs únicos ===
def _build_tab_ids_unicos(popup):

    labels_responsivos = []

    def _envolver_texto(texto, max_chars):
        if not texto:
            return ""

        lineas = []
        wrapper = textwrap.TextWrapper(
            width=max_chars,
            break_long_words=True,
            break_on_hyphens=False,
        )
        for linea in str(texto).splitlines():
            chunks = wrapper.wrap(linea) or [""]
            lineas.extend(chunks)
        return "\n".join(lineas)

    def _registrar_label_responsivo(widget, texto_original):
        labels_responsivos.append({"widget": widget, "text": texto_original})

    def _actualizar_labels_responsivos(_event=None):
        ancho_disponible = max(360, popup.winfo_width() - 90)
        max_chars = max(28, int(ancho_disponible / 7))
        for item in labels_responsivos:
            widget = item["widget"]
            if not widget.winfo_exists():
                continue
            texto_envuelto = _envolver_texto(item["text"], max_chars)
            widget.configure(text=texto_envuelto, wraplength=ancho_disponible)

    popup.bind("<Configure>", _actualizar_labels_responsivos)

    # --- Título ---
    Label(
        popup,
        text="IDs únicos",
        font=("Segoe UI", 12, "bold"),
        bg=APP_BG_COLOR,
        fg="white",
    ).pack(pady=(16, 8), padx=20, anchor="w")

    lbl_descripcion = Label(
        popup,
        text="Asigná uno o más valores fijos a un ID de campo no mapeado.\nSirve para inputs, textareas, selects y checkboxes: con ➕ cargás varios valores y en cada envío se elige uno al azar.\nPara checkboxes usá SI o NO como valor (igual que en el Excel) para que se marque o no.",
        font=("Segoe UI", 9),
        bg=APP_BG_COLOR,
        fg="#ddd",
        justify="left",
        anchor="w",
    )
    lbl_descripcion.pack(fill="x", padx=20, anchor="w")
    _registrar_label_responsivo(
        lbl_descripcion,
        "Asigná uno o más valores fijos a un ID de campo no mapeado.\nSirve para inputs, textareas, selects y checkboxes: con ➕ cargás varios valores y en cada envío se elige uno al azar.\nPara checkboxes usá SI o NO como valor (igual que en el Excel) para que se marque o no.",
    )

    ttk.Separator(popup, orient="horizontal").pack(fill="x", padx=20, pady=10)

    # --- Formulario de ingreso ---
    frame_inputs = Frame(popup, bg=APP_BG_COLOR)
    frame_inputs.pack(padx=20, anchor="w")

    Label(
        frame_inputs,
        text="ID:",
        bg=APP_BG_COLOR,
        fg="white",
        font=("Segoe UI", 10),
        width=8,
        anchor="w",
    ).grid(row=0, column=0, sticky="w", pady=4)
    entry_id = Entry(frame_inputs, font=("Segoe UI", 10), width=24)
    entry_id.grid(row=0, column=1, padx=(0, 10), pady=4)

    Label(
        frame_inputs,
        text="Descripción (opcional):",
        bg=APP_BG_COLOR,
        fg="white",
        font=("Segoe UI", 10),
        width=20,
        anchor="w",
    ).grid(row=1, column=0, sticky="w", pady=4)
    entry_nombre_campo = Entry(frame_inputs, font=("Segoe UI", 10), width=24)
    entry_nombre_campo.grid(row=1, column=1, padx=(0, 10), pady=4)

    Label(
        frame_inputs,
        text="Valor:",
        bg=APP_BG_COLOR,
        fg="white",
        font=("Segoe UI", 10),
        width=8,
        anchor="w",
    ).grid(row=2, column=0, sticky="w", pady=4)
    frame_valor = Frame(frame_inputs, bg=APP_BG_COLOR)
    frame_valor.grid(row=2, column=1, padx=(0, 10), pady=4, sticky="w")
    entry_valor = Entry(frame_valor, font=("Segoe UI", 10), width=18)
    entry_valor.pack(side=LEFT)

    # --- Varios valores por ID: se acumulan como etiquetas y se elige uno al azar por envío ---
    valores_chips = []

    frame_chips_unicos = Frame(frame_inputs, bg=APP_BG_COLOR)
    frame_chips_unicos.grid(row=3, column=0, columnspan=3, sticky="w", pady=(2, 0))

    def _render_chips_unicos():
        for w in frame_chips_unicos.winfo_children():
            w.destroy()
        if not valores_chips:
            return
        Label(frame_chips_unicos, text="Valores cargados:", bg=APP_BG_COLOR, fg="white",
              font=("Segoe UI", 9)).pack(side=LEFT, padx=(0, 6))
        for val in valores_chips:
            chip = Frame(frame_chips_unicos, bg="#3A1D52", bd=0)
            chip.pack(side=LEFT, padx=(0, 5), pady=1)
            Label(chip, text=val, bg="#3A1D52", fg="#FFD873",
                  font=("Segoe UI", 9, "bold"), padx=7, pady=1).pack(side=LEFT)

            def _quitar_chip(v=val):
                try:
                    valores_chips.remove(v)
                except ValueError:
                    pass
                _render_chips_unicos()

            Button(chip, text="✕", command=_quitar_chip, bg="#3A1D52", fg="#ff9d9d",
                   relief="flat", bd=0, font=("Segoe UI", 8, "bold"),
                   cursor="hand2", padx=4, pady=0,
                   activebackground="#3A1D52", activeforeground="white").pack(side=LEFT)
        Label(frame_chips_unicos, text="(elige uno al azar en cada envío)", bg=APP_BG_COLOR,
              fg="#c9b3de", font=("Segoe UI", 8, "italic")).pack(side=LEFT, padx=(4, 0))

    def _anadir_valor_unico():
        nuevos = normalizar_valores_id_dinamico(entry_valor.get())
        if not nuevos:
            return
        for v in nuevos:
            if v not in valores_chips:
                valores_chips.append(v)
        entry_valor.delete(0, "end")
        _render_chips_unicos()

    Button(
        frame_valor, text="➕", command=_anadir_valor_unico,
        bg=HEADER_BG_COLOR, fg="black", relief="flat",
        font=("Segoe UI", 9, "bold"), cursor="hand2", padx=7, pady=1,
    ).pack(side=LEFT, padx=(4, 0))
    entry_valor.bind("<Return>", lambda _e: _anadir_valor_unico())

    ids_fijos_mapeados = obtener_ids_mapeados_normales()
    warning_id_var = StringVar(value="")
    label_warning_id = Label(
        frame_inputs,
        textvariable=warning_id_var,
        bg=APP_BG_COLOR,
        fg="#ffd38a",
        font=("Segoe UI", 9),
        anchor="w",
        justify="left",
    )
    label_warning_id.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(2, 0))

    # --- Lista de IDs existentes ---
    ttk.Separator(popup, orient="horizontal").pack(fill="x", padx=20, pady=(12, 6))
    Label(popup, text="IDs configurados:", font=("Segoe UI", 10, "bold"),
          bg=APP_BG_COLOR, fg="white").pack(padx=20, anchor="w")

    frame_lista_container = Frame(popup, bg=APP_BG_COLOR)
    frame_lista_container.pack(padx=20, pady=(6, 16), fill="both", expand=True)

    canvas_lista = Canvas(
        frame_lista_container,
        bg=APP_BG_COLOR,
        highlightthickness=0,
        bd=0,
    )
    scroll_lista = ttk.Scrollbar(
        frame_lista_container,
        orient="vertical",
        style="Section.Vertical.TScrollbar",
        command=canvas_lista.yview,
    )
    canvas_lista.configure(yscrollcommand=lambda f, l, _sb=scroll_lista: _autohide_yscroll(_sb, f, l))

    canvas_lista.pack(side=LEFT, fill="both", expand=True)
    scroll_lista.pack(side=RIGHT, fill="y")

    frame_lista = Frame(canvas_lista, bg=APP_BG_COLOR)
    lista_window = canvas_lista.create_window((0, 0), window=frame_lista, anchor="nw")
    popup._scroll_canvas = canvas_lista  # para scroll con ruedita a nivel popup

    def _actualizar_scroll_lista(_event=None):
        try:
            canvas_lista.configure(scrollregion=canvas_lista.bbox("all"))
        except Exception:
            pass

    def _ajustar_ancho_lista(event):
        try:
            canvas_lista.itemconfigure(lista_window, width=event.width)
        except Exception:
            pass

    def _on_lista_mousewheel(event):
        try:
            delta = int(-1 * (event.delta / 120))
            canvas_lista.yview_scroll(delta, "units")
        except Exception:
            return "break"
        return "break"

    frame_lista.bind("<Configure>", _actualizar_scroll_lista)
    canvas_lista.bind("<Configure>", _ajustar_ancho_lista)
    canvas_lista.bind("<MouseWheel>", _on_lista_mousewheel)

    filas_widgets = {}  # row_key -> frame
    entrada_en_edicion = {"index": None}

    label_modo_edicion = Label(
        popup,
        text="",
        bg=APP_BG_COLOR,
        fg="#ffd38a",
        font=("Segoe UI", 9, "bold"),
        anchor="w",
        justify="left",
    )
    label_modo_edicion.pack(fill="x", padx=20, anchor="w", pady=(0, 4))

    def _actualizar_estado_edicion(id_val="", paises=None):
        if id_val:
            alcance = "Todos" if _es_todos_paises(paises) else ", ".join(paises or [])
            label_modo_edicion.config(text=f"Editando ID: {id_val} (Países: {alcance})")
            btn_crear_text.set("Editar")
        else:
            label_modo_edicion.config(text="")
            btn_crear_text.set("Crear")

    # --- Selector de países (3 columnas de 3 países) ---
    frame_paises = Frame(frame_inputs, bg=APP_BG_COLOR)
    frame_paises.grid(row=0, column=2, rowspan=3, padx=(14, 0), sticky="nw")

    Label(
        frame_paises,
        text="Países (opcional)",
        bg=APP_BG_COLOR,
        fg="white",
        font=("Segoe UI", 10, "bold"),
        anchor="w",
    ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 4))

    pais_vars = {}
    paises_disponibles = list(MAPEO_PAISES.keys())

    def _es_todos_paises(paises):
        seleccionados = normalizar_paises_id_dinamico(paises)
        if not seleccionados:
            return True
        return set(seleccionados) == set(paises_disponibles)

    for idx, pais in enumerate(paises_disponibles):
        row = (idx % 3) + 1
        col = idx // 3
        var = BooleanVar(value=False)
        pais_vars[pais] = var
        Checkbutton(
            frame_paises,
            text=pais,
            variable=var,
            bg=APP_BG_COLOR,
            fg="white",
            activebackground=APP_BG_COLOR,
            activeforeground="white",
            selectcolor=APP_BG_COLOR,
            anchor="w",
        ).grid(row=row, column=col, sticky="w", padx=(0, 10), pady=1)

    btn_crear_text = StringVar(value="Crear")

    def _limpiar_formulario_ids_dinamicos():
        entry_id.delete(0, "end")
        entry_nombre_campo.delete(0, "end")
        entry_valor.delete(0, "end")
        valores_chips.clear()
        _render_chips_unicos()
        for var in pais_vars.values():
            var.set(False)
        filtro_texto_var.set("")
        filtro_pais_var.set("Todos")
        warning_id_var.set("")
        entrada_en_edicion["index"] = None
        _actualizar_estado_edicion()
        btn_crear_text.set("Crear")
        _refrescar_lista()

    # Diccionario de abreviaturas de países
    pais_abreviaturas = {
        "Argentina": "AR",
        "Brasil": "BR",
        "Bolivia": "BO",
        "Chile": "CH",
        "Ecuador": "EC",
        "Colombia": "CO",
        "Paraguay": "PY",
        "Peru": "PE",
        "Uruguay": "UY",
    }
    abreviatura_a_pais = {abrev: pais for pais, abrev in pais_abreviaturas.items()}
    filtro_texto_var = StringVar(value="")
    filtro_pais_var = StringVar(value="Todos")

    def _refrescar_lista():
        for w in list(filas_widgets.values()):
            w.destroy()
        filas_widgets.clear()
        datos = cargar_ids_dinamicos()
        entries = datos.get("entries", []) if isinstance(datos, dict) else []
        texto_filtro = filtro_texto_var.get().strip().lower()
        pais_filtro_abrev = filtro_pais_var.get().strip() or "Todos"
        pais_filtro = abreviatura_a_pais.get(pais_filtro_abrev)

        filas_filtradas = []
        for idx, entry in enumerate(entries):
            if not isinstance(entry, dict):
                continue
            id_val = str(entry.get("id") or "").strip()
            if not id_val:
                continue

            nombre_campo, valor_raw = extraer_datos_id_dinamico(entry)
            valores = normalizar_valores_id_dinamico(valor_raw)
            valor_texto = " | ".join(valores) if valores else "(sin valores)"
            paises_entry = normalizar_paises_id_dinamico(entry.get("paises", entry.get("countries")))
            entry_es_todos = _es_todos_paises(paises_entry)
            if entry_es_todos:
                paises_abrev = "Todos"
            else:
                paises_abrev = ", ".join([pais_abreviaturas.get(p, p) for p in paises_entry])

            # Filtro por ID, descripción, valores o país: se puede escribir más de una
            # palabra y matchean las filas que contengan todas (en cualquier campo).
            texto_busqueda = f"{id_val} {nombre_campo} {valor_texto} {paises_abrev}".lower()
            if texto_filtro and not all(t in texto_busqueda for t in texto_filtro.split()):
                continue

            prioridad = 0
            if pais_filtro:
                if entry_es_todos:
                    prioridad = 1
                elif pais_filtro in paises_entry:
                    prioridad = 0
                else:
                    continue

            filas_filtradas.append((
                prioridad,
                idx,
                entry,
                id_val,
                nombre_campo,
                valor_texto,
                paises_abrev,
            ))

        if pais_filtro:
            filas_filtradas.sort(key=lambda x: (x[0], x[1]))

        if not filas_filtradas:
            lbl = Label(frame_lista, text="(ninguno)", bg=APP_BG_COLOR,
                        fg="#aaa", font=("Segoe UI", 9, "italic"))
            lbl.pack(anchor="w")
            filas_widgets["__empty__"] = lbl
            return

        for _prioridad, idx, entry, id_val, nombre_campo, valor_texto, paises_abrev in filas_filtradas:

            fila = Frame(frame_lista, bg=APP_BG_COLOR)
            fila.pack(fill="x", pady=2)
            if nombre_campo:
                texto_fila = f"ID: {id_val}   Valor: {valor_texto}   Descripción: {nombre_campo}   Países: {paises_abrev}"
            else:
                texto_fila = f"ID: {id_val}   Valor: {valor_texto}   Países: {paises_abrev}"
            label_fila = Label(
                fila,
                text=texto_fila,
                bg=APP_BG_COLOR,
                fg="white",
                font=("Segoe UI", 10),
                anchor="w",
                justify="left",
            )
            label_fila.pack(side=LEFT, expand=True, fill="x")
            _registrar_label_responsivo(label_fila, texto_fila)

            def _editar_existente(index=idx):
                d = cargar_ids_dinamicos()
                entries_edit = d.get("entries", []) if isinstance(d, dict) else []
                if index < 0 or index >= len(entries_edit):
                    return

                entry_actual = entries_edit[index]
                id_actual = str(entry_actual.get("id") or "").strip()
                nombre_actual, valor_raw = extraer_datos_id_dinamico(entry_actual)
                valores_actuales = normalizar_valores_id_dinamico(valor_raw)
                paises_actuales = normalizar_paises_id_dinamico(entry_actual.get("paises", entry_actual.get("countries")))

                entry_id.delete(0, "end")
                entry_id.insert(0, id_actual)
                entry_nombre_campo.delete(0, "end")
                entry_nombre_campo.insert(0, nombre_actual)
                entry_valor.delete(0, "end")
                valores_chips.clear()
                valores_chips.extend(valores_actuales)
                _render_chips_unicos()
                for pais, var in pais_vars.items():
                    var.set(pais in paises_actuales)

                entrada_en_edicion["index"] = index
                _actualizar_estado_edicion(id_actual, paises_actuales)
                _actualizar_warning_id()
                canvas_lista.yview_moveto(0)

            Button(
                fila,
                text="Editar",
                command=_editar_existente,
                bg=HEADER_BG_COLOR,
                fg="black",
                relief="flat",
                font=("Segoe UI", 9, "bold"),
                cursor="hand2",
                padx=6,
                pady=1,
            ).pack(side=RIGHT, padx=(4, 0))

            def _eliminar(index=idx):
                d = cargar_ids_dinamicos()
                entries_del = d.get("entries", []) if isinstance(d, dict) else []
                if index < 0 or index >= len(entries_del):
                    return
                entries_del.pop(index)
                d["entries"] = entries_del

                if entrada_en_edicion["index"] == index:
                    _limpiar_formulario_ids_dinamicos()
                elif isinstance(entrada_en_edicion["index"], int) and entrada_en_edicion["index"] > index:
                    entrada_en_edicion["index"] -= 1

                guardar_ids_dinamicos(d)
                _refrescar_lista()

            Button(fila, text="✕", command=_eliminar,
                   bg="#7a2040", fg="white", relief="flat",
                   font=("Segoe UI", 9, "bold"), cursor="hand2",
                   padx=6, pady=1).pack(side=RIGHT)
            filas_widgets[f"row_{idx}"] = fila

        _actualizar_labels_responsivos()

    _refrescar_lista()

    def _actualizar_warning_id(*_):
        current_id = entry_id.get().strip()
        if current_id and current_id in ids_fijos_mapeados:
            texto_warning = "Aviso: este ID ya existe en los IDs Excel.\nConsultar en Más información para ver el listado completo."
            warning_id_var.set(_envolver_texto(texto_warning, max(28, int(max(360, popup.winfo_width() - 90) / 7))))
        else:
            warning_id_var.set("")

    entry_id.bind("<KeyRelease>", _actualizar_warning_id)

    def _agregar():
        id_val = entry_id.get().strip()
        nombre_campo = entry_nombre_campo.get().strip()
        # Valores = etiquetas cargadas con ➕ más lo que quede escrito en el cuadro
        valores_todos = list(valores_chips)
        for v in normalizar_valores_id_dinamico(entry_valor.get()):
            if v not in valores_todos:
                valores_todos.append(v)
        valor_val = " | ".join(valores_todos)
        paises_seleccionados = [pais for pais, var in pais_vars.items() if var.get()]
        if _es_todos_paises(paises_seleccionados):
            paises_seleccionados = []
        if not id_val or not valor_val:
            messagebox.showwarning("IDs únicos", "Completá el ID y al menos un valor.", parent=popup)
            return
        if id_val in ids_fijos_mapeados:
            continuar = messagebox.askyesno(
                "IDs únicos",
                "Ese ID ya existe en los IDs Excel. Consultar en Más información para ver el listado completo.\n\nSi lo guardás acá, solo actuará como respaldo para campos no mapeados o lógica especial.\n\n¿Querés guardarlo igual?",
                parent=popup,
            )
            if not continuar:
                return

        d = cargar_ids_dinamicos()
        valores_nuevos = normalizar_valores_id_dinamico(valor_val)
        if not valores_nuevos:
            messagebox.showwarning("IDs únicos", "No hay valores válidos para guardar.", parent=popup)
            return

        nueva_entry = {
            "id": id_val,
            "valor": compactar_valores_id_dinamico(valores_nuevos),
            "paises": paises_seleccionados,
        }
        if nombre_campo:
            nueva_entry["nombre_campo"] = nombre_campo

        entries = d.get("entries", []) if isinstance(d, dict) else []
        if not isinstance(entries, list):
            entries = []

        idx_edit = entrada_en_edicion["index"]
        if isinstance(idx_edit, int) and 0 <= idx_edit < len(entries):
            entries[idx_edit] = nueva_entry
        else:
            entries.append(nueva_entry)

        d["entries"] = entries

        guardar_ids_dinamicos(d)
        _limpiar_formulario_ids_dinamicos()
        _refrescar_lista()

    # Fila de filtros (columna izquierda)
    frame_filtros = Frame(frame_inputs, bg=APP_BG_COLOR)
    frame_filtros.grid(row=5, column=0, columnspan=2, pady=(2, 0), sticky="w")

    frame_col_filtro_texto = Frame(frame_filtros, bg=APP_BG_COLOR)
    frame_col_filtro_texto.grid(row=0, column=0, padx=(0, 10), sticky="w")
    Label(
        frame_col_filtro_texto,
        text="Filtro:",
        bg=APP_BG_COLOR,
        fg="white",
        font=("Segoe UI", 10),
    ).pack(side=LEFT, padx=(0, 6))
    entry_filtro = Entry(frame_col_filtro_texto, font=("Segoe UI", 10), width=20, textvariable=filtro_texto_var)
    entry_filtro.pack(side=LEFT)

    frame_col_filtro_pais = Frame(frame_filtros, bg=APP_BG_COLOR)
    frame_col_filtro_pais.grid(row=0, column=1, padx=(0, 10), sticky="w")
    valores_filtro_pais = ["Todos"] + [pais_abreviaturas[p] for p in paises_disponibles if p in pais_abreviaturas]
    estilo_combo_filtro = ttk.Style()
    estilo_combo_filtro.configure(
        "FiltroPais.TCombobox",
        fieldbackground="white",
        background="white",
        foreground="black",
    )
    estilo_combo_filtro.map(
        "FiltroPais.TCombobox",
        fieldbackground=[("readonly", "white")],
        selectbackground=[("readonly", "white")],
        selectforeground=[("readonly", "black")],
        foreground=[("readonly", "black")],
        background=[("readonly", "white")],
    )
    combo_filtro_pais = ttk.Combobox(
        frame_col_filtro_pais,
        textvariable=filtro_pais_var,
        values=valores_filtro_pais,
        state="readonly",
        width=8,
        style="FiltroPais.TCombobox",
    )
    combo_filtro_pais.pack(side=LEFT)
    combo_filtro_pais.set("Todos")

    # CTAs debajo de la columna de países (Limpiar izquierda, Crear derecha)
    frame_ctas_paises = Frame(frame_inputs, bg=APP_BG_COLOR)
    frame_ctas_paises.grid(row=5, column=2, pady=(2, 0), sticky="ew")
    frame_ctas_paises.grid_columnconfigure(0, weight=1)
    frame_ctas_paises.grid_columnconfigure(1, weight=1)
    ancho_cta = 12

    Button(
        frame_ctas_paises,
        text="Limpiar",
        command=_limpiar_formulario_ids_dinamicos,
        bg=HEADER_BG_COLOR,
        fg="black",
        relief="flat",
        font=("Segoe UI", 10, "bold"),
        cursor="hand2",
        width=ancho_cta,
        padx=10,
        pady=3,
    ).grid(row=0, column=0, sticky="w")

    Button(
        frame_ctas_paises,
        textvariable=btn_crear_text,
        command=_agregar,
        bg=HEADER_BG_COLOR,
        fg="black",
        relief="flat",
        font=("Segoe UI", 10, "bold"),
        cursor="hand2",
        width=ancho_cta,
        padx=10,
        pady=3,
    ).grid(row=0, column=1, sticky="e")

    entry_filtro.bind("<KeyRelease>", lambda _event: _refrescar_lista())
    combo_filtro_pais.bind("<<ComboboxSelected>>", lambda _event: _refrescar_lista())

    _actualizar_labels_responsivos()


# === Dependencias ===
def _build_tab_dependencias(popup):
    """Tab para registrar dependencias padre→hijo por país."""

    Label(
        popup,
        text="Dependencias entre IDs",
        font=("Segoe UI", 12, "bold"),
        bg=APP_BG_COLOR,
        fg="white",
    ).pack(pady=(16, 8), padx=20, anchor="w")

    descripcion_popup = (
        "Registrá qué ID hijo depende de un ID padre. El sistema los rellenará en orden respetando la dependencia."
    )
    Label(
        popup,
        text=descripcion_popup,
        font=("Segoe UI", 9),
        bg=APP_BG_COLOR,
        fg="#ddd",
        justify="left",
        anchor="w",
        wraplength=700,
    ).pack(fill="x", padx=20, anchor="w")

    ttk.Separator(popup, orient="horizontal").pack(fill="x", padx=20, pady=10)

    paises_disponibles = [
        "Argentina", "Bolivia", "Brasil", "Chile",
        "Colombia", "Ecuador", "Paraguay", "Peru", "Uruguay",
    ]
    padre_var = StringVar(value="")
    hijo_var = StringVar(value="")
    pais_vars = {pais: BooleanVar(value=False) for pais in paises_disponibles}
    btn_guardar_text = StringVar(value="Crear")
    entrada_en_edicion = {"index": None}

    frame_inputs = Frame(popup, bg=APP_BG_COLOR)
    frame_inputs.pack(padx=20, fill="x")

    # Países primero, en filas de 3
    Label(frame_inputs, text="Países:", bg=APP_BG_COLOR, fg="white", font=("Segoe UI", 10), width=14, anchor="nw").grid(row=0, column=0, sticky="nw", pady=4)
    frame_paises = Frame(frame_inputs, bg=APP_BG_COLOR)
    frame_paises.grid(row=0, column=1, sticky="w", pady=4)
    for idx, pais in enumerate(paises_disponibles):
        row = idx // 3
        col = idx % 3
        Checkbutton(
            frame_paises,
            text=pais,
            variable=pais_vars[pais],
            bg=APP_BG_COLOR,
            fg="white",
            activebackground=APP_BG_COLOR,
            activeforeground="white",
            selectcolor=APP_BG_COLOR,
            anchor="w",
        ).grid(row=row, column=col, sticky="w", padx=(0, 10), pady=1)

    # ID Padre
    Label(frame_inputs, text="ID Padre:", bg=APP_BG_COLOR, fg="white", font=("Segoe UI", 10), width=14, anchor="w").grid(row=1, column=0, sticky="w", pady=4)
    Entry(frame_inputs, font=("Segoe UI", 10), width=24, textvariable=padre_var).grid(row=1, column=1, sticky="w", pady=4, padx=(0, 12))

    # ID Hijo
    Label(frame_inputs, text="ID Hijo:", bg=APP_BG_COLOR, fg="white", font=("Segoe UI", 10), width=14, anchor="w").grid(row=2, column=0, sticky="w", pady=4)
    Entry(frame_inputs, font=("Segoe UI", 10), width=24, textvariable=hijo_var).grid(row=2, column=1, sticky="w", pady=4, padx=(0, 12))

    # CTAs en una fila abajo
    frame_ctas = Frame(frame_inputs, bg=APP_BG_COLOR)
    frame_ctas.grid(row=3, column=0, columnspan=2, pady=(8, 0), sticky="w")
    Button(frame_ctas, text="Limpiar", command=lambda: limpiar_formulario(), bg=HEADER_BG_COLOR, fg="black", relief="flat", font=("Segoe UI", 10, "bold"), cursor="hand2", width=12, padx=10, pady=3).pack(side=LEFT, padx=(0, 8))
    Button(frame_ctas, textvariable=btn_guardar_text, command=lambda: guardar_dependencia(), bg=HEADER_BG_COLOR, fg="black", relief="flat", font=("Segoe UI", 10, "bold"), cursor="hand2", width=12, padx=10, pady=3).pack(side=LEFT)

    ttk.Separator(popup, orient="horizontal").pack(fill="x", padx=20, pady=(10, 6))
    Label(popup, text="Dependencias registradas:", font=("Segoe UI", 10, "bold"), bg=APP_BG_COLOR, fg="white").pack(padx=20, anchor="w")

    frame_lista_container = Frame(popup, bg=APP_BG_COLOR)
    frame_lista_container.pack(padx=20, pady=(6, 14), fill="both", expand=True)

    canvas_lista = Canvas(frame_lista_container, bg=APP_BG_COLOR, highlightthickness=0, bd=0)
    scroll_lista = ttk.Scrollbar(frame_lista_container, orient="vertical", style="Section.Vertical.TScrollbar", command=canvas_lista.yview)
    canvas_lista.configure(yscrollcommand=lambda f, l, _sb=scroll_lista: _autohide_yscroll(_sb, f, l))
    canvas_lista.pack(side=LEFT, fill="both", expand=True)
    scroll_lista.pack(side=RIGHT, fill="y")

    frame_lista = Frame(canvas_lista, bg=APP_BG_COLOR)
    lista_window = canvas_lista.create_window((0, 0), window=frame_lista, anchor="nw")
    popup._scroll_canvas = canvas_lista  # para scroll con ruedita a nivel popup
    frame_lista.bind("<Configure>", lambda _event=None: canvas_lista.configure(scrollregion=canvas_lista.bbox("all")))
    canvas_lista.bind("<Configure>", lambda event: canvas_lista.itemconfigure(lista_window, width=event.width))

    filas_widgets = {}

    def limpiar_formulario():
        padre_var.set("")
        hijo_var.set("")
        for var in pais_vars.values():
            var.set(False)
        entrada_en_edicion["index"] = None
        btn_guardar_text.set("Crear")
        refrescar_lista()

    def guardar_dependencia():
        padre = padre_var.get().strip()
        hijo = hijo_var.get().strip()
        paises = [p for p, v in pais_vars.items() if v.get()]
        if not padre or not hijo:
            messagebox.showwarning("Dependencias", "Completá ID Padre e ID Hijo.", parent=popup)
            return
        nueva_dep = {"padre": padre, "hijo": hijo, "paises": paises}
        deps = cargar_dependencias()
        idx_edit = entrada_en_edicion["index"]
        if isinstance(idx_edit, int) and 0 <= idx_edit < len(deps):
            deps[idx_edit] = nueva_dep
        else:
            deps.append(nueva_dep)
        guardar_dependencias(deps)
        limpiar_formulario()

    def refrescar_lista():
        for widget in list(filas_widgets.values()):
            widget.destroy()
        filas_widgets.clear()
        deps = cargar_dependencias()
        if not deps:
            lbl = Label(frame_lista, text="(sin dependencias)", bg=APP_BG_COLOR, fg="#aaa", font=("Segoe UI", 9, "italic"))
            lbl.pack(anchor="w")
            filas_widgets["__empty__"] = lbl
            return
        for idx, dep in enumerate(deps):
            padre = dep.get("padre", "")
            hijo = dep.get("hijo", "")
            paises = dep.get("paises", [])
            paises_txt = ", ".join(paises) if paises else "Todos"
            texto_fila = f"Padre: {padre}   →   Hijo: {hijo}   |   Países: {paises_txt}"
            fila = Frame(frame_lista, bg=APP_BG_COLOR)
            fila.pack(fill="x", pady=2)
            Label(fila, text=texto_fila, bg=APP_BG_COLOR, fg="white", font=("Segoe UI", 10), anchor="w", justify="left", wraplength=660).pack(side=LEFT, expand=True, fill="x")
            Button(fila, text="Editar", command=lambda i=idx: editar_dependencia(i), bg=HEADER_BG_COLOR, fg="black", relief="flat", font=("Segoe UI", 9, "bold"), cursor="hand2", padx=6, pady=1).pack(side=RIGHT, padx=(4, 0))
            Button(fila, text="✕", command=lambda i=idx: eliminar_dependencia(i), bg="#7a2040", fg="white", relief="flat", font=("Segoe UI", 9, "bold"), cursor="hand2", padx=6, pady=1).pack(side=RIGHT)
            filas_widgets[f"row_{idx}"] = fila

    def editar_dependencia(index):
        deps = cargar_dependencias()
        if index < 0 or index >= len(deps):
            return
        dep = deps[index]
        padre_var.set(dep.get("padre", ""))
        hijo_var.set(dep.get("hijo", ""))
        for pais, var in pais_vars.items():
            var.set(pais in dep.get("paises", []))
        entrada_en_edicion["index"] = index
        btn_guardar_text.set("Editar")

    def eliminar_dependencia(index):
        deps = cargar_dependencias()
        if index < 0 or index >= len(deps):
            return
        deps.pop(index)
        guardar_dependencias(deps)
        limpiar_formulario()

    refrescar_lista()


# === Campos detectados en los formularios ===
def _pais_key(pais):
    return str(pais or "").strip().lower().replace(" ", "_")


def leer_campos_detectados(pais):
    """Lee json/nuevos_campos_<pais>.json y devuelve (lista_campos, fecha_deteccion)."""
    ruta = os.path.join(JSON_DIR, f"nuevos_campos_{_pais_key(pais)}.json")
    if not os.path.exists(ruta):
        return [], ""
    try:
        with open(ruta, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return [], ""
    campos = data.get("campos_nuevos") or []
    if not isinstance(campos, list):
        campos = []
    return campos, str(data.get("ultima_deteccion") or "")


def quitar_campo_detectado(pais, field_id):
    """Saca un campo del reporte nuevos_campos_<pais>.json (ya fue configurado)."""
    ruta = os.path.join(JSON_DIR, f"nuevos_campos_{_pais_key(pais)}.json")
    if not os.path.exists(ruta):
        return
    try:
        with open(ruta, "r", encoding="utf-8") as f:
            data = json.load(f)
        campos = data.get("campos_nuevos") or []
        data["campos_nuevos"] = [
            c for c in campos
            if not (isinstance(c, dict) and str(c.get("id") or "").strip() == str(field_id).strip())
        ]
        tmp = ruta + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, ruta)
    except Exception as e:
        print(f"⚠️ No se pudo actualizar {os.path.basename(ruta)}: {e}")


def consolidar_ids_dinamicos():
    """Fusiona entradas duplicadas (mismo ID + mismo alcance de países) en una sola
    con todos los valores. El backend ya las fusionaba al rellenar; esto alinea la UI."""
    d = cargar_ids_dinamicos()
    entries = d.get("entries", []) if isinstance(d, dict) else []
    consolidadas = []
    indice = {}  # (id, frozenset(paises)) -> posición en consolidadas
    cambio = False
    for entry in entries:
        if not isinstance(entry, dict):
            continue
        entry_id = str(entry.get("id") or "").strip()
        paises = normalizar_paises_id_dinamico(entry.get("paises", entry.get("countries")))
        clave = (entry_id, frozenset(p.lower() for p in paises))
        nombre, valor_raw = extraer_datos_id_dinamico(entry)
        valores = normalizar_valores_id_dinamico(valor_raw)
        if clave in indice:
            destino = consolidadas[indice[clave]]
            _, dest_raw = extraer_datos_id_dinamico(destino)
            dest_vals = normalizar_valores_id_dinamico(dest_raw)
            for v in valores:
                if v not in dest_vals:
                    dest_vals.append(v)
            destino["valor"] = compactar_valores_id_dinamico(dest_vals) if dest_vals else ""
            if nombre and not destino.get("nombre_campo"):
                destino["nombre_campo"] = nombre
            cambio = True
        else:
            indice[clave] = len(consolidadas)
            consolidadas.append(entry)
    if cambio:
        d["entries"] = consolidadas
        guardar_ids_dinamicos(d)
    return cambio


def _buscar_entry_para(entries, pais, field_id):
    """Índice de la entrada de IDs únicos que aplica a (país, id), o None."""
    for i, entry in enumerate(entries):
        if not isinstance(entry, dict):
            continue
        if str(entry.get("id") or "").strip() != str(field_id).strip():
            continue
        paises_entry = normalizar_paises_id_dinamico(entry.get("paises", entry.get("countries")))
        if not paises_entry or pais in paises_entry:
            return i
    return None


def _valores_asignados_para(pais, field_id):
    """Todos los valores asignados en IDs únicos que aplican a (país, id) —
    unión de todas las entradas que matcheen, igual que hace el backend al rellenar."""
    datos = cargar_ids_dinamicos()
    entries = datos.get("entries", []) if isinstance(datos, dict) else []
    valores = []
    for entry in entries:
        if not isinstance(entry, dict):
            continue
        if str(entry.get("id") or "").strip() != str(field_id).strip():
            continue
        paises_entry = normalizar_paises_id_dinamico(entry.get("paises", entry.get("countries")))
        if paises_entry and pais not in paises_entry:
            continue
        _, valor_raw = extraer_datos_id_dinamico(entry)
        for v in normalizar_valores_id_dinamico(valor_raw):
            if v not in valores:
                valores.append(v)
    return valores


def _agregar_valores_campo_detectado(pais, field_id, label, valor_texto):
    """Suma valores (sin duplicar) a la entrada de IDs únicos del campo detectado."""
    valores_nuevos = normalizar_valores_id_dinamico(valor_texto)
    if not valores_nuevos:
        return False

    d = cargar_ids_dinamicos()
    entries = d.get("entries", []) if isinstance(d, dict) else []
    if not isinstance(entries, list):
        entries = []

    idx = _buscar_entry_para(entries, pais, field_id)
    valores = _valores_asignados_para(pais, field_id) if idx is not None else []
    for v in valores_nuevos:
        if v not in valores:
            valores.append(v)

    nueva_entry = {
        "id": field_id,
        "valor": compactar_valores_id_dinamico(valores),
        "paises": [pais] if idx is None else normalizar_paises_id_dinamico(
            entries[idx].get("paises", entries[idx].get("countries"))
        ) or [],
    }
    if label and str(label).strip() and str(label).strip() != str(field_id).strip():
        nueva_entry["nombre_campo"] = str(label).strip()

    if idx is not None:
        entries[idx] = nueva_entry
    else:
        entries.append(nueva_entry)

    d["entries"] = entries
    guardar_ids_dinamicos(d)
    return True


def _quitar_valor_campo_detectado(pais, field_id, valor):
    """Quita un valor puntual de todas las entradas que apliquen a (país, id);
    las entradas que queden sin valores se eliminan."""
    d = cargar_ids_dinamicos()
    entries = d.get("entries", []) if isinstance(d, dict) else []
    nuevas = []
    for entry in entries:
        if isinstance(entry, dict) and str(entry.get("id") or "").strip() == str(field_id).strip():
            paises_entry = normalizar_paises_id_dinamico(entry.get("paises", entry.get("countries")))
            if not paises_entry or pais in paises_entry:
                _, valor_raw = extraer_datos_id_dinamico(entry)
                valores = [v for v in normalizar_valores_id_dinamico(valor_raw) if v != valor]
                if not valores:
                    continue  # entrada sin valores → se elimina
                entry["valor"] = compactar_valores_id_dinamico(valores)
                entry.pop("valores", None)
        nuevas.append(entry)
    d["entries"] = nuevas
    guardar_ids_dinamicos(d)


def _build_tab_campos_detectados(popup):
    """Lista los campos nuevos detectados en los formularios y permite asignarles un valor."""

    Label(
        popup,
        text="Campos detectados",
        font=("Segoe UI", 12, "bold"),
        bg=APP_BG_COLOR,
        fg="white",
    ).pack(pady=(16, 8), padx=20, anchor="w")

    Label(
        popup,
        text="Campos nuevos que la automatización detectó en los formularios durante las corridas. "
             "Escribí un valor y apretá ➕ Añadir valor (o Enter) las veces que quieras: cada valor se guarda "
             "al instante como ID único y, si hay varios, en cada envío la app elige uno al azar. "
             "Cuando termines con un campo apretá ✔ Listo: sale de esta lista y lo seguís editando desde la solapa IDs únicos.",
        font=("Segoe UI", 9),
        bg=APP_BG_COLOR,
        fg="#ddd",
        justify="left",
        anchor="w",
        wraplength=760,
    ).pack(fill="x", padx=20, anchor="w")

    ttk.Separator(popup, orient="horizontal").pack(fill="x", padx=20, pady=10)

    paises_disponibles = list(MAPEO_PAISES.keys())
    pais_var = StringVar(value=paises_disponibles[0] if paises_disponibles else "")
    info_var = StringVar(value="")

    frame_top = Frame(popup, bg=APP_BG_COLOR)
    frame_top.pack(fill="x", padx=20)

    Label(frame_top, text="País:", bg=APP_BG_COLOR, fg="white", font=("Segoe UI", 10), width=8, anchor="w").pack(side=LEFT)

    _estilo_cd = ttk.Style()
    _estilo_cd.configure("CamposDet.TCombobox", fieldbackground="white", background="white", foreground="black")
    _estilo_cd.map("CamposDet.TCombobox", fieldbackground=[("readonly", "white")], foreground=[("readonly", "black")])
    combo_pais = ttk.Combobox(frame_top, textvariable=pais_var, values=paises_disponibles, state="readonly", width=20, style="CamposDet.TCombobox")
    combo_pais.pack(side=LEFT, padx=(0, 12))

    Label(frame_top, textvariable=info_var, bg=APP_BG_COLOR, fg="#ffd38a", font=("Segoe UI", 9), anchor="w").pack(side=LEFT)

    ttk.Separator(popup, orient="horizontal").pack(fill="x", padx=20, pady=(10, 6))

    frame_lista_container = Frame(popup, bg=APP_BG_COLOR)
    frame_lista_container.pack(padx=20, pady=(6, 16), fill="both", expand=True)

    canvas_lista = Canvas(frame_lista_container, bg=APP_BG_COLOR, highlightthickness=0, bd=0)
    scroll_lista = ttk.Scrollbar(frame_lista_container, orient="vertical", style="Section.Vertical.TScrollbar", command=canvas_lista.yview)
    canvas_lista.configure(yscrollcommand=lambda f, l, _sb=scroll_lista: _autohide_yscroll(_sb, f, l))
    canvas_lista.pack(side=LEFT, fill="both", expand=True)
    scroll_lista.pack(side=RIGHT, fill="y")

    frame_lista = Frame(canvas_lista, bg=APP_BG_COLOR)
    lista_window = canvas_lista.create_window((0, 0), window=frame_lista, anchor="nw")
    popup._scroll_canvas = canvas_lista  # para scroll con ruedita a nivel popup
    frame_lista.bind("<Configure>", lambda _event=None: canvas_lista.configure(scrollregion=canvas_lista.bbox("all")))
    canvas_lista.bind("<Configure>", lambda event: canvas_lista.itemconfigure(lista_window, width=event.width))

    def _on_mousewheel(event):
        try:
            canvas_lista.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            return "break"
        return "break"
    canvas_lista.bind("<MouseWheel>", _on_mousewheel)

    filas = []

    def _refrescar():
        for w in filas:
            try:
                w.destroy()
            except Exception:
                pass
        filas.clear()

        pais = pais_var.get().strip()
        campos, fecha = leer_campos_detectados(pais)
        info_var.set(f"Última detección: {fecha}" if fecha else "Sin detecciones registradas para este país.")

        if not campos:
            lbl = Label(frame_lista, text="(no se detectaron campos nuevos para este país)", bg=APP_BG_COLOR,
                        fg="#aaa", font=("Segoe UI", 9, "italic"))
            lbl.pack(anchor="w")
            filas.append(lbl)
            return

        for campo in campos:
            if not isinstance(campo, dict):
                continue
            field_id = str(campo.get("id") or "").strip()
            if not field_id:
                continue
            label = str(campo.get("label") or campo.get("name") or field_id).strip() or field_id
            tipo = str(campo.get("type") or "text").strip()
            requerido = " · requerido" if campo.get("required") else ""

            fila = Frame(frame_lista, bg=APP_BG_COLOR)
            fila.pack(fill="x", pady=(2, 8))
            filas.append(fila)

            Label(
                fila,
                text=f"{label}   [{tipo}{requerido}]",
                bg=APP_BG_COLOR, fg="white", font=("Segoe UI", 10, "bold"),
                anchor="w", justify="left", wraplength=740,
            ).pack(anchor="w")
            Label(
                fila,
                text=f"ID: {field_id}",
                bg=APP_BG_COLOR, fg="#c9b3de", font=("Segoe UI", 8),
                anchor="w", justify="left", wraplength=740,
            ).pack(anchor="w")

            # --- Valores asignados como "chips" (uno por valor, con ✕ para quitar) ---
            chips_frame = Frame(fila, bg=APP_BG_COLOR)
            chips_frame.pack(fill="x", pady=(3, 0), anchor="w")

            def _render_chips(cf=None, fid=field_id, lbl=None):
                cf = cf if cf is not None else chips_frame
                for w in cf.winfo_children():
                    w.destroy()
                pais_actual = pais_var.get().strip()
                valores = _valores_asignados_para(pais_actual, fid)
                Label(cf, text="Valores:", bg=APP_BG_COLOR, fg="white", font=("Segoe UI", 9)).pack(side=LEFT, padx=(0, 6))
                if not valores:
                    Label(cf, text="(ninguno todavía)", bg=APP_BG_COLOR, fg="#aaa",
                          font=("Segoe UI", 9, "italic")).pack(side=LEFT)
                    return
                for val in valores:
                    chip = Frame(cf, bg="#3A1D52", bd=0)
                    chip.pack(side=LEFT, padx=(0, 5), pady=1)
                    Label(chip, text=val, bg="#3A1D52", fg="#FFD873",
                          font=("Segoe UI", 9, "bold"), padx=7, pady=1).pack(side=LEFT)

                    def _quitar(v=val, f=fid, c=cf, l=lbl):
                        _quitar_valor_campo_detectado(pais_var.get().strip(), f, v)
                        _render_chips(c, f, l)

                    Button(chip, text="✕", command=_quitar, bg="#3A1D52", fg="#ff9d9d",
                           relief="flat", bd=0, font=("Segoe UI", 8, "bold"),
                           cursor="hand2", padx=4, pady=0,
                           activebackground="#3A1D52", activeforeground="white").pack(side=LEFT)

            fila_input = Frame(fila, bg=APP_BG_COLOR)
            fila_input.pack(fill="x", pady=(3, 0))

            valor_var = StringVar(value="")
            entry_valor = Entry(fila_input, font=("Segoe UI", 10), width=28, textvariable=valor_var)
            entry_valor.pack(side=LEFT, padx=(0, 8))

            def _anadir(fid=field_id, lbl=label, vv=valor_var, cf=chips_frame):
                pais_actual = pais_var.get().strip()
                if not _agregar_valores_campo_detectado(pais_actual, fid, lbl, vv.get()):
                    messagebox.showwarning("Campos detectados", "Escribí un valor antes de añadir.", parent=popup)
                    return
                vv.set("")
                _render_chips(cf, fid, lbl)

            Button(
                fila_input, text="➕ Añadir valor", command=_anadir,
                bg=HEADER_BG_COLOR, fg="black", relief="flat",
                font=("Segoe UI", 9, "bold"), cursor="hand2", padx=10, pady=2,
            ).pack(side=LEFT)
            entry_valor.bind("<Return>", lambda _e, f=_anadir: f())

            def _listo(fid=field_id, lbl=label):
                pais_actual = pais_var.get().strip()
                if not _valores_asignados_para(pais_actual, fid):
                    seguir = messagebox.askyesno(
                        "Campos detectados",
                        f"\"{lbl}\" no tiene ningún valor asignado.\n\n"
                        "¿Quitarlo de la lista igual? (si es un select, la app va a elegir "
                        "una opción al azar; si es texto, quedará vacío)",
                        parent=popup,
                    )
                    if not seguir:
                        return
                quitar_campo_detectado(pais_actual, fid)
                _refrescar()

            Button(
                fila_input, text="✔ Listo", command=_listo,
                bg="#4c8a5f", fg="white", relief="flat",
                font=("Segoe UI", 9, "bold"), cursor="hand2", padx=10, pady=2,
                activebackground="#5fa876", activeforeground="white",
            ).pack(side=LEFT, padx=(8, 0))

            _render_chips(chips_frame, field_id, label)

    combo_pais.bind("<<ComboboxSelected>>", lambda _e: _refrescar())
    _refrescar()


def abrir_popup_ids_dinamicos(parent):
    """Abre el popup unificado de IDs Dinámicos (Campos detectados, IDs únicos, IDs Excel y Dependencias)."""
    _ensure_styles()
    # Fusionar entradas duplicadas (mismo ID + mismo alcance) para que la UI
    # muestre y edite todos los valores juntos, igual que los usa el backend.
    try:
        consolidar_ids_dinamicos()
    except Exception as e:
        print(f"⚠️ No se pudo consolidar ids_dinamicos: {e}")

    popup = Toplevel(parent)
    popup.title("IDs Dinámicos")
    popup.configure(bg=APP_BG_COLOR)
    popup.geometry("930x620")
    popup.minsize(760, 520)
    popup.resizable(True, True)

    icon_path = os.path.join(ASSET_DIR, "icon.ico") if ASSET_DIR else ""
    if icon_path and os.path.exists(icon_path):
        try:
            popup.iconbitmap(icon_path)
        except Exception:
            try:
                popup.iconbitmap(default=icon_path)
            except Exception:
                pass

    notebook_ids = ttk.Notebook(popup, style="TNotebook")
    notebook_ids.pack(fill="both", expand=True, padx=10, pady=10)

    tab_campos_detectados = Frame(notebook_ids, bg=APP_BG_COLOR)
    tab_ids_unicos = Frame(notebook_ids, bg=APP_BG_COLOR)
    tab_ids_excel = Frame(notebook_ids, bg=APP_BG_COLOR)
    tab_dependencias = Frame(notebook_ids, bg=APP_BG_COLOR)
    notebook_ids.add(tab_campos_detectados, text="Campos detectados")
    notebook_ids.add(tab_ids_unicos, text="IDs únicos")
    notebook_ids.add(tab_ids_excel, text="IDs Excel")
    notebook_ids.add(tab_dependencias, text="Dependencias")

    _build_tab_campos_detectados(tab_campos_detectados)
    _build_tab_ids_unicos(tab_ids_unicos)
    _build_tab_ids_excel(tab_ids_excel)
    _build_tab_dependencias(tab_dependencias)

    # Scroll con ruedita en cualquier parte del popup: scrollea la lista de la
    # solapa activa y frena la propagación al bind_all de la ventana principal.
    def _on_popup_mousewheel(event):
        try:
            tab_actual = notebook_ids.nametowidget(notebook_ids.select())
            canvas = getattr(tab_actual, "_scroll_canvas", None)
            if canvas is not None and canvas.winfo_exists():
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass
        return "break"

    popup.bind("<MouseWheel>", _on_popup_mousewheel)

    return popup
