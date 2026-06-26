"""
main_interface.py — Interfaz gráfica principal de la aplicación (Tkinter).
Pestañas: Carga de datos, Ejecución por país, Validación de campos, Programación y IDs Dinámicos.
Permite cargar Excels, correr formularios, ver resultados y configurar el envío de emails.
"""
import os
import sys
import threading
import subprocess
import importlib
import importlib.util
import time
from tkinter import *
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
from PIL import Image, ImageTk
import shutil
import glob
import pandas as pd
from datetime import datetime, timedelta
import re
import json
import textwrap
import webbrowser

from .helpers_interface import (
    analizar_errores_excel,
    abrir_excel,
    crear_zip_de_carpeta,
    cargar_config_global,
    guardar_config_global,
    obtener_email_destinatario,
    cargar_ids_dinamicos,
    guardar_ids_dinamicos,
    cargar_dependencias,
    guardar_dependencias,
    obtener_ids_mapeados_normales,
    sincronizar_excels_de_pais,
    TEMPORALES_DIR,
)
# === LIMPIEZA DE TEMPORALES ===
def limpiar_temporales():
    """Elimina todo el contenido de la carpeta temporales al iniciar la app."""
    try:
        if not os.path.exists(TEMPORALES_DIR):
            return
        for f in glob.glob(os.path.join(TEMPORALES_DIR, '*')):
            if os.path.isfile(f):
                try:
                    os.remove(f)
                except Exception:
                    pass
            elif os.path.isdir(f):
                try:
                    shutil.rmtree(f)
                except Exception:
                    pass
    except Exception:
        pass
from .field_validation_ui import build_field_validation_tab
from tkinter import StringVar

# Importar utilidades compartidas (import absoluto)
from utils.fixed_field_mapping_store import (
    build_excel_columns_for_country,
    list_available_fixed_mapping_countries,
    load_effective_country_form_config,
    save_country_fixed_field_mapping,
)

# === RUTAS BASE ===
from utils.paths import BASE_DIR, BUNDLE_DIR, FORMS_DIR, DATA_DIR, ASSET_DIR, RESULTS_DIR, JSON_DIR

def abrir_carpeta_resultados():
    """Abre la carpeta de resultados en el explorador."""
    try:
        if not os.path.exists(RESULTS_DIR):
            messagebox.showwarning("Resultados", "No se encontró la carpeta de resultados.")
            return

        if os.name == "nt":
            os.startfile(RESULTS_DIR)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.run(["open", RESULTS_DIR], check=False)
        else:
            subprocess.run(["xdg-open", RESULTS_DIR], check=False)
    except Exception as exc:
        messagebox.showerror("Resultados", f"No se pudo abrir la carpeta de resultados:\n{exc}")


APP_BG_COLOR = "#5D3C7A"
HEADER_BG_COLOR = "#9c6fb4"
SECTION_BG_COLOR = HEADER_BG_COLOR
SECTION_CONTAINER_BG_COLOR = "#7a549a"
SECTION_CTA_BG_COLOR = "#345474"
PRIMARY_TEXT_COLOR = "#000000"


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


def obtener_texto_encabezado(nombre_columna):
    """Devuelve el texto visible a usar en el encabezado del Treeview."""
    if not nombre_columna:
        return nombre_columna

    return nombre_columna

# === MAPEO DE PAÍSES ===
MAPEO_PAISES = {
    "Argentina": "Formulario_Argentina_Main",
    "Bolivia": "Formulario_Bolivia_Main", 
    "Brasil": "Formulario_Brasil_Main",
    "Chile": "Formulario_Chile_Main",
    "Colombia": "Formulario_Colombia_Main",
    "Ecuador": "Formulario_Ecuador_Main",
    "Paraguay": "Formulario_Paraguay_Main",
    "Peru": "Formulario_Peru_Main",
    "Uruguay": "Formulario_Uruguay_Main"
}

# === MAPEO DE ABREVIATURAS DE PAÍSES ===
ABREV_PAISES = {
    "ar": "Argentina",
    "bo": "Bolivia",
    "br": "Brasil",
    "cl": "Chile",
    "co": "Colombia",
    "ec": "Ecuador",
    "py": "Paraguay",
    "pe": "Peru",
    "uy": "Uruguay"
}

# Mapeo inverso para convertir país completo a abreviatura
PAIS_A_ABREV = {v: k for k, v in ABREV_PAISES.items()}


def normalizar_datos_excel(pais_abrev, landing_url, expected_form_url):
    """
    Normaliza datos del Excel y acepta 2 formatos:
    - Nuevo: columna A = país abreviado, B = landing_url, C = expected_form_url (iframe).
      C puede estar vacío: formulario embebido en la landing, sin buscar iframe.
    - Legacy: columna A = landing_url, B = expected_form_url.
      B vacío: mismo criterio (formulario en la misma página).

    Returns: (pais_abrev_valido, landing_url_limpia, expected_form_url_limpia, error)
    """
    error = None

    def limpiar_url(url):
        if not url:
            return ""
        url_limpia = str(url).strip()
        url_limpia = url_limpia.replace('\n', '').replace('\r', '').replace('\t', '')
        return url_limpia.strip()

    def parece_url(texto):
        valor = str(texto or "").strip().lower()
        return valor.startswith("http://") or valor.startswith("https://")

    raw_col_a = limpiar_url(pais_abrev)
    raw_col_b = limpiar_url(landing_url)
    raw_col_c = limpiar_url(expected_form_url)

    # Formato nuevo: A=país, B=landing, C=form.
    pais_normalizado = str(raw_col_a).lower() if raw_col_a else ""
    if pais_normalizado in ABREV_PAISES:
        landing_url_limpia = raw_col_b
        expected_form_url_limpia = raw_col_c
        if not landing_url_limpia:
            error = "Landing URL faltante en formato con país (A/B/C)."
        return pais_normalizado, landing_url_limpia, expected_form_url_limpia, error

    # Formato legacy: A=landing, B=form (iframe). B vacío = formulario embebido en la misma página.
    if parece_url(raw_col_a):
        landing_url_limpia = raw_col_a
        expected_form_url_limpia = raw_col_b if parece_url(raw_col_b) else ""

        # Compatibilidad extra: si B no es URL pero C sí, usar C como form URL.
        if not expected_form_url_limpia and parece_url(raw_col_c):
            expected_form_url_limpia = raw_col_c

        return "", landing_url_limpia, expected_form_url_limpia, error

    # Caso inválido: no parece país válido ni URL legacy.
    if raw_col_a:
        error = f"País '{raw_col_a}' no válido. Use: {', '.join(ABREV_PAISES.keys())} o formato legacy (URL en A)."
    else:
        error = "Columna A vacía: use país abreviado o landing URL (formato legacy)."

    return "", raw_col_b, raw_col_c, error


def normalizar_paises_programacion(raw_paises):
    """Convierte abreviaturas o nombres guardados a nombres de país válidos."""
    paises_normalizados = []
    for pais in normalizar_paises_id_dinamico(raw_paises):
        if pais in MAPEO_PAISES and pais not in paises_normalizados:
            paises_normalizados.append(pais)
            continue

        pais_expandido = ABREV_PAISES.get(pais.strip().lower())
        if pais_expandido and pais_expandido not in paises_normalizados:
            paises_normalizados.append(pais_expandido)

    return paises_normalizados


def _asegurar_paths_modulos():
    """Asegura paths para imports dinámicos en modo fuente y ejecutable."""
    core_dir = os.path.join(BASE_DIR, "core")
    for p in (FORMS_DIR, core_dir, BASE_DIR):
        if p not in sys.path:
            sys.path.insert(0, p)


def _get_run_func(pais_nombre):
    """Devuelve la función run_formularios_<País> usando el runner genérico consolidado."""
    _asegurar_paths_modulos()
    from _runner_common import get_runner
    return get_runner(pais_nombre)


def _get_environments():
    _asegurar_paths_modulos()
    from _runner_common import ENVIRONMENTS
    return ENVIRONMENTS

# === CONFIGURACIONES DISPONIBLES ===
CONFIGURACIONES = [
    {"browser": "chrome", "viewport": "fullscreen"},
    {"browser": "chrome", "viewport": "600x738"},
    {"browser": "edge", "viewport": "fullscreen"},
    {"browser": "edge", "viewport": "600x738"},
    {"browser": "firefox", "viewport": "fullscreen"},
    {"browser": "firefox", "viewport": "600x738"}
]

# Variables globales para ordenamiento
orden_ascendente = {}

# Variable global para programación de tests
programacion_actual = None 

# === CLASE PARA EDITAR CELDAS DIRECTAMENTE ===
class CellEditor:
    def __init__(self, tree_widget):
        self.tree = tree_widget
        self.entry = None
        self.current_item = None
        self.current_column = None
        
    def start_edit(self, event):
        """Inicia la edición de una celda al hacer click"""
        # Identificar el ítem y columna clickeada
        item = self.tree.identify_row(event.y)
        columna = self.tree.identify_column(event.x)
        
        if not item or not columna or columna == '#0':
            return
            
        # No permitir editar en encabezados
        if item == "":
            return
            
        col_index = int(columna.replace('#', '')) - 1
        
        # Verificar que la columna existe
        columnas = self.tree["columns"]
        if col_index >= len(columnas) or columnas[0] in ["info", "error"]:
            return
        
        self.current_item = item
        self.current_column = col_index
        
        # Obtener el valor actual (quitar prefijo de protección si existe)
        valores = list(self.tree.item(item)["values"])
        valor_actual = valores[col_index] if col_index < len(valores) else ""
        valor_actual = str(valor_actual)
        if valor_actual.startswith("\x00"):
            valor_actual = valor_actual[1:]
        
        # Obtener la posición y dimensiones de la celda
        bbox = self.tree.bbox(item, col_index)
        if not bbox:
            return
            
        # Crear el Entry para editar
        if self.entry:
            self.entry.destroy()
            
        self.entry = Entry(self.tree, font=("Segoe UI", 9))
        self.entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        self.entry.insert(0, str(valor_actual))
        self.entry.select_range(0, END)
        self.entry.focus()
        
        # Bind eventos
        self.entry.bind('<Return>', self.finish_edit)
        self.entry.bind('<Escape>', self.cancel_edit)
        self.entry.bind('<FocusOut>', self.finish_edit)
        
    def finish_edit(self, event=None):
        """Termina la edición y guarda el valor"""
        if not self.entry or not self.current_item:
            return
            
        nuevo_valor = self.entry.get()
        # Re-aplicar protección si el nuevo valor empieza con 0 seguido de dígitos
        if len(nuevo_valor) >= 2 and nuevo_valor.startswith("0") and nuevo_valor[1:].isdigit():
            nuevo_valor_tree = "\x00" + nuevo_valor
        else:
            nuevo_valor_tree = nuevo_valor

        # Actualizar los valores en el treeview
        valores = list(self.tree.item(self.current_item)["values"])
        if self.current_column < len(valores):
            valores[self.current_column] = nuevo_valor_tree
        else:
            # Si la columna no existe en los valores, extender la lista
            while len(valores) <= self.current_column:
                valores.append("")
            valores[self.current_column] = nuevo_valor_tree
            
        self.tree.item(self.current_item, values=valores)
        
        self.cleanup()
        
    def cancel_edit(self, event=None):
        """Cancela la edición"""
        self.cleanup()
        
    def cleanup(self):
        """Limpia el editor"""
        if self.entry:
            self.entry.destroy()
            self.entry = None
        self.current_item = None
        self.current_column = None

def safe_str(val):
    if val is None:
        return ""
    
    if isinstance(val, str):
        v = val.strip()
        if v.endswith(".0") and v[:-2].lstrip("-").isdigit() and not v.startswith("0"):
            v = v[:-2]
        return v

    if isinstance(val, float):
        if val == int(val):
            return str(int(val))
        return str(val)

    return str(val).strip()

# === FUNCIONES MEJORADAS PARA MANEJO DE EXCEL EN INTERFAZ ===
def cargar_excel_a_tabla(excel_nombre, tree_widget, cell_editor=None):
    """Carga el contenido del Excel en el widget Treeview mostrando solo 5 filas visibles"""
    try:
        # Limpiar tabla existente
        for item in tree_widget.get_children():
            tree_widget.delete(item)
        
        excel_ruta = os.path.join(DATA_DIR, f"{excel_nombre}.xlsx")
        
        if os.path.exists(excel_ruta):
            # Leer el Excel
            df = pd.read_excel(excel_ruta, dtype=str, keep_default_na=False)

            columnas_normalizadas = {}
            for nombre_columna in df.columns:
                alias = obtener_texto_encabezado(nombre_columna)
                if alias and alias != nombre_columna:
                    columnas_normalizadas[nombre_columna] = alias

            if columnas_normalizadas:
                df.rename(columns=columnas_normalizadas, inplace=True)
            
            # Configurar columnas del Treeview
            columnas = list(df.columns)
            # Forzar nombre de columna 'Modelo' en la posición original si existe, o agregarla si no
            modelo_idx = None
            for idx, col in enumerate(columnas):
                if col.strip().lower() in ("modelo", "model", "models"):
                    modelo_idx = idx
                    columnas[idx] = "Modelo"
            if modelo_idx is None:
                # Si no existe, agregar al final
                columnas.append("Modelo")
                modelo_idx = len(columnas) - 1
                df["Modelo"] = ""
            tree_widget["columns"] = columnas
            tree_widget["show"] = "headings"
            
            # MEJORA 1: Configurar estilo con líneas divisorias
            estilo = ttk.Style()
            estilo.configure(
                "Section.Treeview",
                background="white",
                foreground=PRIMARY_TEXT_COLOR,
                rowheight=25,
                fieldbackground="white",
                borderwidth=0,
                relief="flat",
            )

            estilo.configure(
                "Section.Treeview.Heading",
                background="#d8def2",
                foreground=PRIMARY_TEXT_COLOR,
                relief="flat",
                borderwidth=0,
            )

            estilo.map(
                "Section.Treeview",
                background=[('selected', '#8da9d9')],
                foreground=[('selected', 'white')],
            )
            estilo.map(
                "Section.Treeview.Heading",
                background=[('active', '#c1c9e5')],
            )

            tree_widget.configure(style="Section.Treeview")
            
            # Configurar encabezados con función de ordenamiento
            for col in columnas:
                heading_text = obtener_texto_encabezado(col)
                tree_widget.heading(col, text=heading_text, command=lambda c=col: ordenar_columna(c, tree_widget))

                valores_columna = df[col].dropna()
                max_contenido = 0
                if not valores_columna.empty:
                    max_contenido = valores_columna.astype(str).map(len).max()

                if col.lower() in ("url", "formulario"):
                    caracteres_objetivo = 50
                else:
                    caracteres_objetivo = max(len(col), max_contenido)

                # Convertir a una estimación en píxeles (aprox. 8px por carácter)
                ancho_pixeles = max(80, int(caracteres_objetivo * 8))
                ancho_minimo = max(60, int(len(col) * 8))
                tree_widget.column(col, width=ancho_pixeles, minwidth=ancho_minimo, stretch=False)

            if columnas:
                # Permite que la última columna acompañe el ancho del Treeview sin bloquear el scroll horizontal
                tree_widget.column(columnas[-1], stretch=True)
            
            # Agregar datos (todas las filas, pero solo se mostrarán 5 visualmente)
            # Asegurar que la columna 'Modelo' sea tipo string/object para evitar errores de asignación
            if "Modelo" in df.columns:
                df["Modelo"] = df["Modelo"].astype(str)
            for index, row in df.iterrows():
                valores = [""] * len(columnas)
                # Copiar valores existentes
                for idx, col in enumerate(df.columns):
                    if idx < len(columnas):
                        val = row[col]
                        valores[idx] = safe_str(val)

                # Autocompletar 'Modelo' desde la URL (busca en columna 'Formulario', 'form_url' o 'url')
                col_form = None
                prioridad = ["formulario", "form_url", "url"]
                for nombre in prioridad:
                    for idx, col in enumerate(columnas):
                        if col.strip().lower() == nombre:
                            col_form = idx
                            break
                    if col_form is not None:
                        break
                if col_form is not None and modelo_idx is not None:
                    form_url = valores[col_form] if col_form < len(valores) else ""
                    import re
                    m = re.search(r"[?&]model=([^&#]+)", form_url)
                    if m:
                        from urllib.parse import unquote
                        raw_model = m.group(1)
                        model_val = unquote(raw_model)
                        model_val = re.sub(r"[^a-zA-Z0-9áéíóúÁÉÍÓÚüÜñÑ\s]", " ", model_val)
                        model_val = re.sub(r"\s+", " ", model_val).strip()
                        valores[modelo_idx] = model_val
                        # Actualizar también el DataFrame para que se guarde correctamente
                        df.at[index, "Modelo"] = model_val
                # Proteger valores que empiezan con "0" seguido de dígitos (ej: "09...")
                # para que tkinter Treeview no los convierta a entero y pierda el 0 inicial.
                valores_protegidos = []
                for v in valores:
                    sv = str(v) if v is not None else ""
                    if len(sv) >= 2 and sv.startswith("0") and sv[1:].isdigit():
                        sv = "\x00" + sv  # prefijo invisible para bloquear conversión numérica
                    valores_protegidos.append(sv)
                tree_widget.insert("", "end", values=valores_protegidos)
            
            # Reconfigurar el editor de celdas
            if cell_editor:
                tree_widget.bind('<Button-1>', cell_editor.start_edit)

            aplicar_colores_filas(tree_widget)
            
            return True
        else:
            # Mostrar mensaje de que no existe el archivo
            tree_widget["columns"] = ["info"]
            tree_widget["show"] = "headings"
            tree_widget.heading("info", text="Información")
            tree_widget.column("info", width=400)
            tree_widget.insert("", "end", values=["Archivo Excel no encontrado. Use 'Crear Excel' para generarlo."])
            return False
    except Exception as e:
        print(f" Error al cargar Excel: {e}")
        # Mostrar error en la tabla
        tree_widget["columns"] = ["error"]
        tree_widget["show"] = "headings"
        tree_widget.heading("error", text="Error")
        tree_widget.column("error", width=400)
        tree_widget.insert("", "end", values=[f"Error al cargar Excel: {str(e)}"])
        return False

def ordenar_columna(columna, tree_widget):
    """Ordena la columna al hacer click en el encabezado"""
    global orden_ascendente
    
    # Obtener todos los datos del treeview
    datos = []
    for item in tree_widget.get_children():
        valores = tree_widget.item(item)["values"]
        datos.append(valores)
    
    if not datos:
        return
    
    # Obtener índice de la columna
    columnas = tree_widget["columns"]
    if columna not in columnas:
        return
    
    col_index = columnas.index(columna)
    
    # Determinar dirección de ordenamiento
    if columna not in orden_ascendente:
        orden_ascendente[columna] = True
    else:
        orden_ascendente[columna] = not orden_ascendente[columna]
    
    # Ordenar datos
    try:
        datos_ordenados = sorted(datos, key=lambda x: str(x[col_index]).lower() if x[col_index] is not None else "", 
                                reverse=not orden_ascendente[columna])
    except:
        datos_ordenados = sorted(datos, key=lambda x: str(x[col_index]), 
                                reverse=not orden_ascendente[columna])
    
    # Limpiar y reinsertar datos ordenados
    for item in tree_widget.get_children():
        tree_widget.delete(item)
    
    for fila in datos_ordenados:
        tree_widget.insert("", "end", values=fila)
    
    # Actualizar indicador visual en el encabezado
    indicador = " ▲" if orden_ascendente[columna] else " ▼"
    for col in columnas:
        encabezado = obtener_texto_encabezado(col)
        if col == columna:
            tree_widget.heading(col, text=f"{encabezado}{indicador}")
        else:
            tree_widget.heading(col, text=encabezado)

    aplicar_colores_filas(tree_widget)

def guardar_desde_tabla(excel_nombre, tree_widget):
    """Guarda los datos del Treeview al archivo Excel"""
    try:
        excel_ruta = os.path.join(DATA_DIR, f"{excel_nombre}.xlsx")
        
        # Obtener datos del Treeview
        datos = []
        columnas = tree_widget["columns"]
        
        # Verificar si estamos mostrando mensajes de información/error
        if columnas and columnas[0] in ["info", "error"]:
            messagebox.showerror("Error", "No hay datos válidos para guardar. Cree el Excel primero.")
            return False
        
        for item in tree_widget.get_children():
            valores = tree_widget.item(item)["values"]
            # Limpiar el prefijo \x00 que protege ceros iniciales en el Treeview
            valores = [str(v).lstrip("\x00") if str(v).startswith("\x00") else str(v) for v in valores]
            # Autocompletar model/models si corresponde
            try:
                col_form = None
                col_model = None
                for idx, col in enumerate(columnas):
                    if col.lower() in ("formulario", "form_url", "url"):
                        col_form = idx
                    if col.lower() in ("model", "models"):
                        col_model = idx
                if col_form is not None and col_model is not None:
                    form_url = str(valores[col_form]) if col_form < len(valores) else ""
                    # Buscar model= en la URL
                    import re
                    m = re.search(r"[?&]model=([^&#]+)", form_url)
                    if m:
                        raw_model = m.group(1)
                        # Decodificar %20 y similares
                        from urllib.parse import unquote
                        model_val = unquote(raw_model)
                        # Solo letras y espacios
                        model_val = re.sub(r"[^a-zA-ZáéíóúÁÉÍÓÚüÜñÑ\s]", " ", model_val)
                        model_val = re.sub(r"\s+", " ", model_val).strip()
                        # Reemplazar en la fila
                        valores = list(valores)
                        valores[col_model] = model_val
            except Exception as e:
                print(f"[WARN] Autocompletado model/models falló: {e}")
            datos.append(valores)

        if datos and columnas:
            # Crear DataFrame y guardar
            df = pd.DataFrame(datos, columns=columnas)
            df = df.astype(str)

            with pd.ExcelWriter(excel_ruta, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
            messagebox.showinfo("Éxito", f"Cambios guardados correctamente en:\n{excel_ruta}")
            return True
        else:
            messagebox.showerror("Error", "No hay datos para guardar")
            return False
            
    except Exception as e:
        messagebox.showerror("Error", f"Error al guardar cambios: {str(e)}")
        print(f"Error al guardar: {e}")
        return False

def actualizar_botones_excel(frame_pais, fila, excel_nombre, btn_actualizar, btn_guardar, tree_widget, cell_editor=None):
    """Actualiza el estado de los botones de Excel"""
    excel_ruta = os.path.join(DATA_DIR, f"{excel_nombre}.xlsx")
    
    if os.path.exists(excel_ruta):
        # Habilitar botones de actualizar y guardar
        btn_actualizar.config(state="normal")
        btn_guardar.config(state="normal")
        # Cargar el contenido del Excel
        cargar_excel_a_tabla(excel_nombre, tree_widget, cell_editor)
    else:
        # Deshabilitar botones si no existe el Excel
        btn_actualizar.config(state="disabled")
        btn_guardar.config(state="disabled")
        # Limpiar tabla y mostrar mensaje
        for item in tree_widget.get_children():
            tree_widget.delete(item)
        tree_widget["columns"] = ["info"]
        tree_widget["show"] = "headings"
        tree_widget.heading("info", text="Información")
        tree_widget.column("info", width=400)
        tree_widget.insert("", "end", values=["Cree el Excel primero usando el botón 'Crear Excel'"])

def agregar_fila(tree_widget):
    """Agrega una nueva fila vacía a la tabla"""
    try:
        columnas = tree_widget["columns"]
        if columnas and columnas[0] not in ["info", "error"]:
            nueva_fila = [""] * len(columnas)
            tree_widget.insert("", "end", values=nueva_fila)
            aplicar_colores_filas(tree_widget)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo agregar fila: {e}")

def eliminar_fila(tree_widget):
    """Elimina la fila seleccionada"""
    try:
        seleccion = tree_widget.selection()
        if seleccion:
            tree_widget.delete(seleccion[0])
            aplicar_colores_filas(tree_widget)
        else:
            messagebox.showwarning("Advertencia", "Seleccione una fila para eliminar")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo eliminar fila: {e}")

# MEJORA 2: Función para clonar fila
def clonar_fila(tree_widget):
    """Clona la fila seleccionada y la inserta debajo"""
    try:
        seleccion = tree_widget.selection()
        if not seleccion:
            messagebox.showwarning("Advertencia", "Seleccione una fila para clonar")
            return
            
        item_seleccionado = seleccion[0]
        valores_originales = tree_widget.item(item_seleccionado)["values"]
        
        # Insertar nueva fila con los mismos valores debajo de la seleccionada
        tree_widget.insert("", tree_widget.index(item_seleccionado) + 1, values=valores_originales)
        aplicar_colores_filas(tree_widget)
        
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo clonar la fila: {e}")


def aplicar_colores_filas(tree_widget):
    """Alterna colores de filas para mejorar legibilidad."""
    try:
        tree_widget.tag_configure("odd", background="white")
        tree_widget.tag_configure("even", background="#f1edf6")

        for indice, item in enumerate(tree_widget.get_children()):
            tag = "odd" if indice % 2 == 0 else "even"
            tree_widget.item(item, tags=(tag,))
    except Exception:
        pass

# ══════════════════════════════════════════════════════════════════════════════
# ESTADO GLOBAL DE EJECUCIÓN — bloqueo UI + botón DETENER
# ══════════════════════════════════════════════════════════════════════════════

_run_state = {
    "running": False,
    "stop_event": None,
    "enviar_btn": None,   # botón "Enviar Leads"
    "lt_btn": None,       # botón "Ejecutar en LambdaTest"
    "stop_btn": None,     # botón "Detener"
    "root": None,
}


def _set_running(running: bool):
    """Habilita/deshabilita botones según si hay una ejecución activa."""
    _run_state["running"] = running
    _set_manual_input_callback(running)
    root = _run_state.get("root")

    def _apply():
        state = "disabled" if running else "normal"
        for key in ("enviar_btn", "lt_btn"):
            btn = _run_state.get(key)
            if btn:
                try:
                    btn.configure(state=state)
                except Exception:
                    pass
        stop_btn = _run_state.get("stop_btn")
        if stop_btn:
            try:
                if running:
                    stop_btn.pack(side=RIGHT, padx=6)
                else:
                    stop_btn.pack_forget()
            except Exception:
                pass

    if root:
        root.after(0, _apply)
    else:
        _apply()


def _request_stop():
    ev = _run_state.get("stop_event")
    if ev:
        ev.set()
        print("⛔ Detención solicitada — la ejecución se detendrá al finalizar el lead actual.")


# ── Broker de input manual — conecta el hilo selenium con la UI de tkinter ──

class _ManualInputBroker:
    """Muestra un diálogo de tkinter desde el hilo principal cuando el filler lo solicita."""

    def __init__(self, root_widget):
        self._root = root_widget
        self._request = None          # (field_id, label)
        self._response = [None]
        self._req_ready = threading.Event()
        self._resp_ready = threading.Event()
        self._active = False

    def start_polling(self):
        self._active = True
        self._root.after(400, self._poll)

    def stop_polling(self):
        self._active = False

    def _poll(self):
        if self._req_ready.is_set():
            self._req_ready.clear()
            fid, label = self._request
            self._show_dialog(fid, label)
        if self._active:
            self._root.after(400, self._poll)

    def _show_dialog(self, field_id, label):
        dialog = Toplevel(self._root)
        dialog.title("Campo no mapeado")
        dialog.geometry("480x230")
        dialog.configure(bg=APP_BG_COLOR)
        dialog.resizable(False, False)
        dialog.grab_set()
        dialog.lift()
        dialog.focus_force()

        Label(dialog,
              text="Se encontró un campo no mapeado en el formulario",
              font=("Segoe UI", 11, "bold"), bg=APP_BG_COLOR, fg="white").pack(pady=(16, 2))
        Label(dialog, text=f"Campo: {label}   (id: {field_id})",
              font=("Segoe UI", 9), bg=APP_BG_COLOR, fg="#ddd").pack()
        Label(dialog, text="Ingresá el valor a completar (vacío = omitir este campo):",
              font=("Segoe UI", 9), bg=APP_BG_COLOR, fg="#ddd").pack(pady=(12, 4))

        entry_var = StringVar()
        entry = Entry(dialog, textvariable=entry_var, font=("Segoe UI", 11), width=36)
        entry.pack(pady=(0, 12))
        entry.focus()

        def _ok(_ev=None):
            self._response[0] = entry_var.get().strip() or None
            self._resp_ready.set()
            dialog.destroy()

        def _skip():
            self._response[0] = None
            self._resp_ready.set()
            dialog.destroy()

        entry.bind("<Return>", _ok)
        frame_btns = Frame(dialog, bg=APP_BG_COLOR)
        frame_btns.pack()
        ttk.Button(frame_btns, text="Usar este valor", command=_ok,
                   style="Section.TButton").pack(side=LEFT, padx=4)
        ttk.Button(frame_btns, text="Omitir campo", command=_skip,
                   style="Section.TButton").pack(side=LEFT, padx=4)

    def request_value(self, field_id, label):
        """Llamado desde hilo selenium. Bloquea hasta que el usuario responde (máx. 3 min)."""
        self._response[0] = None
        self._resp_ready.clear()
        self._request = (field_id, label)
        self._req_ready.set()
        self._resp_ready.wait(timeout=180)
        return self._response[0]


# Al iniciar/detener ejecución: activa/desactiva el callback global en base_form_filler
def _set_manual_input_callback(active: bool):
    try:
        _asegurar_paths_modulos()
        from base_form_filler import set_global_manual_input_callback
        broker = _run_state.get("broker")
        if active and broker:
            set_global_manual_input_callback(broker.request_value)
        else:
            set_global_manual_input_callback(None)
    except Exception:
        pass


# === FUNCIONES DE INTERFAZ ORIGINALES ===
def ejecutar_script_configurable(nombre_script_base, selected_browser, selected_viewport, headless=None):
    """Ejecuta el script base con la configuración seleccionada de forma dinámica."""
    if headless is None:
        headless = False

    # --- CONFIGURACIONES DE SUFIJOS ---
    config_map = {
        'chrome': {'fullscreen': '_chrome_desktop', '600x738': '_chrome_mobile'},
        'firefox': {'fullscreen': '_firefox_desktop', '600x738': '_firefox_mobile'},
        'edge': {'fullscreen': '_edge_desktop', '600x738': '_edge_mobile'}
    }

    browser = selected_browser.get()
    viewport = selected_viewport.get()

    sufijo = config_map.get(browser, {}).get(viewport)
    if not sufijo:
        messagebox.showerror("Error", f"Configuración no soportada: {browser} - {viewport}")
        return

    # --- DETERMINAR PAÍS ---
    nombre_base_sin_ext = nombre_script_base.replace('.py', '')
    pais_encontrado = None

    for pais_nombre, base_nombre in MAPEO_PAISES.items():
        if base_nombre == nombre_base_sin_ext:  # Cambiado de 'in' a '=='
            pais_encontrado = pais_nombre
            break

    if not pais_encontrado:
        messagebox.showerror("Error",
            f"No se pudo determinar el país del script: {nombre_script_base}\n"
            f"Países soportados: {', '.join(MAPEO_PAISES.keys())}")
        return

    # --- USAR MAIN UNIFICADO ---
    script_main = f"Formulario_{pais_encontrado}_Main.py"
    env_param = f"{selected_browser.get()}_{'desktop' if selected_viewport.get() == 'fullscreen' else 'mobile'}"

    def _run_form():
        try:
            env_config = _get_environments().get(env_param)
            if env_config is None:
                messagebox.showerror("Error", f"Entorno '{env_param}' no reconocido.")
                return
            run_func = _get_run_func(pais_encontrado)
            run_func(browser=env_config["browser"], viewport=env_config["viewport"], headless=headless)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo ejecutar {pais_encontrado}:\n{e}")

    threading.Thread(target=_run_form, daemon=True).start()

def ejecutar_script(nombre_script):
    """Función original para compatibilidad con scripts existentes"""
    script_path = os.path.join(FORMS_DIR, nombre_script)

    if os.path.exists(script_path):
        threading.Thread(target=lambda: subprocess.run([sys.executable, script_path], check=False), daemon=True).start()
        messagebox.showinfo("Ejecución", f"Ejecutando script:\n{nombre_script}")
        return

    # Fallback para ejecutable empaquetado sin scripts .py externos.
    nombre_base_sin_ext = nombre_script.replace('.py', '')
    pais_encontrado = None
    for pais_nombre, base_nombre in MAPEO_PAISES.items():
        if base_nombre == nombre_base_sin_ext:
            pais_encontrado = pais_nombre
            break

    if not pais_encontrado:
        messagebox.showerror("Error", f"No se encontró el archivo ni se pudo mapear país:\n{nombre_script}")
        return

    def _run_fallback():
        try:
            environments = _get_environments()
            env_config = environments.get("chrome_desktop") or next(iter(environments.values()), None)
            if env_config is None:
                messagebox.showerror("Error", f"No hay ENVIRONMENTS válidos.")
                return
            run_func = _get_run_func(pais_encontrado)
            run_func(browser=env_config["browser"], viewport=env_config["viewport"], headless=False)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo ejecutar {nombre_script}:\n{e}")

    threading.Thread(target=_run_fallback, daemon=True).start()
    messagebox.showinfo("Ejecución", f"Ejecutando (modo empaquetado):\n{nombre_script}")

def crear_y_actualizar_excel(excel_nombre, fila, frame_pais, btn_actualizar, btn_guardar, tree_widget, cell_editor=None):
    """Crea el Excel y actualiza los botones"""
    try:
        abrir_excel(excel_nombre)
        # Esperar un poco para que se cree el archivo y luego actualizar
        frame_pais.after(2000, lambda: actualizar_botones_excel(frame_pais, fila, excel_nombre, btn_actualizar, btn_guardar, tree_widget, cell_editor))
    except Exception as e:
        print(f" Error al crear Excel: {e}")

def _autohide_yscroll(sb, first, last):
    """Oculta el scrollbar cuando todo el contenido cabe; lo muestra cuando hay overflow.
    Funciona con pack y grid — detecta el geometry manager automáticamente.
    NO chequea winfo_ismapped() antes de ocultar: en Windows el widget puede estar
    colocado (winfo_manager='grid') pero aún no pintado (winfo_ismapped=False)."""
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


def _build_generar_excels_tab(parent):
    """Tab para generar Excels con datos aleatorios por país a partir de URLs."""
    from utils.data_generator import generar_fila_datos

    # Keywords de países en URLs
    _COUNTRY_URL_KEYWORDS = {
        "argentina": "Argentina",
        "bolivia": "Bolivia",
        "brasil": "Brasil",
        "brazil": "Brasil",
        "chile": "Chile",
        "colombia": "Colombia",
        "ecuador": "Ecuador",
        "paraguay": "Paraguay",
        "peru": "Peru",
        "uruguay": "Uruguay",
    }

    def _detectar_pais_desde_urls(texto):
        texto_lower = texto.lower()
        for kw, pais in _COUNTRY_URL_KEYWORDS.items():
            if kw in texto_lower:
                return pais
        return None

    frame = Frame(parent, bg=APP_BG_COLOR, padx=30, pady=20)
    frame.pack(fill="both", expand=True)

    Label(frame, text="Generar Excels con Datos",
          font=("Segoe UI", 14, "bold"), bg=APP_BG_COLOR, fg="white").pack(anchor="w", pady=(0, 4))
    Label(frame,
          text="Pegá las URLs (landing+form en pares, o solo forms). Detecta país automáticamente. Genera datos aleatorios válidos.",
          font=("Segoe UI", 9), bg=APP_BG_COLOR, fg="#ddd").pack(anchor="w", pady=(0, 10))

    # ── Selector de país — RadioButtons (estilo checkboxes) ─────────────────
    Label(frame, text="País:", bg=APP_BG_COLOR, fg="white",
          font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 4))

    paises_disponibles = list(MAPEO_PAISES.keys())
    pais_var = StringVar(value="")

    frame_pais_checks = Frame(frame, bg=APP_BG_COLOR)
    frame_pais_checks.pack(anchor="w", pady=(0, 8))
    for _idx, _pnombre in enumerate(paises_disponibles):
        Radiobutton(
            frame_pais_checks,
            text=_pnombre,
            variable=pais_var,
            value=_pnombre,
            bg=APP_BG_COLOR,
            fg="white",
            selectcolor=APP_BG_COLOR,
            activebackground=APP_BG_COLOR,
            activeforeground="white",
            font=("Segoe UI", 9),
        ).grid(row=_idx // 5, column=_idx % 5, sticky="w", padx=6, pady=1)

    # ── Modo de URLs ─────────────────────────────────────────────────────────
    url_mode_var = StringVar(value="landing_form")
    frame_url_mode = Frame(frame, bg=APP_BG_COLOR)
    frame_url_mode.pack(anchor="w", pady=(0, 6))
    Label(frame_url_mode, text="Tipo de URLs:", bg=APP_BG_COLOR, fg="white",
          font=("Segoe UI", 9, "bold")).pack(side=LEFT, padx=(0, 10))
    Radiobutton(frame_url_mode, text="Landing + Form (pares)",
                variable=url_mode_var, value="landing_form",
                bg=APP_BG_COLOR, fg="white", selectcolor=APP_BG_COLOR,
                activebackground=APP_BG_COLOR, activeforeground="white",
                font=("Segoe UI", 9)).pack(side=LEFT, padx=(0, 14))
    Radiobutton(frame_url_mode, text="Solo Forms",
                variable=url_mode_var, value="solo_forms",
                bg=APP_BG_COLOR, fg="white", selectcolor=APP_BG_COLOR,
                activebackground=APP_BG_COLOR, activeforeground="white",
                font=("Segoe UI", 9)).pack(side=LEFT)

    # Warning de discrepancia país seleccionado vs URLs
    url_country_warning_var = StringVar(value="")
    url_country_warning_lbl = Label(
        frame, textvariable=url_country_warning_var,
        bg=APP_BG_COLOR, fg="#f4a261",
        font=("Segoe UI", 9, "italic"), wraplength=700, justify="left",
    )
    url_country_warning_lbl.pack(anchor="w", pady=(0, 4))

    # ── Área de URLs ────────────────────────────────────────────────────────
    _url_label_var = StringVar(value="URLs (landing + form — de a pares, una por línea):")
    url_area_lbl = Label(frame, textvariable=_url_label_var,
                         bg=APP_BG_COLOR, fg="white", font=("Segoe UI", 10))
    url_area_lbl.pack(anchor="w", pady=(0, 4))

    def _actualizar_label_urls(*_):
        if url_mode_var.get() == "solo_forms":
            _url_label_var.set("URLs de forms (una por línea):")
        else:
            _url_label_var.set("URLs (landing + form — de a pares, una por línea):")

    url_mode_var.trace_add("write", _actualizar_label_urls)

    frame_text_borde = Frame(frame, bg=SECTION_CONTAINER_BG_COLOR, padx=1, pady=1)
    frame_text_borde.pack(fill="x", pady=(0, 8))
    frame_text_inner = Frame(frame_text_borde, bg="white")
    frame_text_inner.pack(fill="both")

    url_scroll = ttk.Scrollbar(frame_text_inner, orient="vertical")
    url_text = Text(frame_text_inner, height=8, font=("Consolas", 9),
                    bg="white", fg="#111", relief="flat", bd=4,
                    yscrollcommand=lambda f, l: _autohide_yscroll(url_scroll, f, l), wrap="none")
    url_scroll.config(command=url_text.yview)
    url_text.pack(side=LEFT, fill="both", expand=True)
    url_scroll.pack(side=RIGHT, fill="y")

    def _on_url_change(event=None):
        texto = url_text.get("1.0", END)
        detectado = _detectar_pais_desde_urls(texto)
        seleccionado = pais_var.get()
        if detectado:
            if not seleccionado:
                pais_var.set(detectado)
                _toggle_brasil_panel()
                url_country_warning_var.set("")
            elif detectado != seleccionado:
                url_country_warning_var.set(
                    f"⚠ Las URLs sugieren {detectado} pero tenés seleccionado {seleccionado}. "
                    "Verificá antes de generar."
                )
            else:
                url_country_warning_var.set("")
        else:
            url_country_warning_var.set("")

    url_text.bind("<KeyRelease>", _on_url_change)

    # ── Status + botón carpeta ───────────────────────────────────────────────
    status_var = StringVar(value="")
    status_lbl = Label(frame, textvariable=status_var,
                       bg=APP_BG_COLOR, fg="#a8e6a3",
                       font=("Segoe UI", 9, "italic"), wraplength=700, justify="left")
    status_lbl.pack(anchor="w", pady=(4, 0))

    _ultima_ruta = [None]

    def _abrir_carpeta():
        ruta = _ultima_ruta[0]
        if not ruta:
            return
        carpeta = os.path.dirname(ruta)
        try:
            if os.name == "nt":
                os.startfile(carpeta)
            elif sys.platform == "darwin":
                import subprocess
                subprocess.run(["open", carpeta], check=False)
            else:
                import subprocess
                subprocess.run(["xdg-open", carpeta], check=False)
        except Exception:
            pass

    btn_abrir_carpeta = ttk.Button(frame, text="📁 Abrir carpeta del Excel",
                                   command=_abrir_carpeta, style="FolderCTA.TButton")

    # ── Botones ─────────────────────────────────────────────────────────────
    frame_btns = Frame(frame, bg=APP_BG_COLOR)
    frame_btns.pack(anchor="w", pady=(8, 0))

    def _generar():
        import pandas as pd

        pais = pais_var.get()
        if not pais:
            status_var.set("Seleccioná un país.")
            status_lbl.config(fg="#f4a261")
            return

        raw_lines = [ln.strip() for ln in url_text.get("1.0", END).splitlines()]
        lines = [ln for ln in raw_lines if ln]
        if not lines:
            status_var.set("Ingresá al menos una URL.")
            status_lbl.config(fg="#f4a261")
            return

        # Construir pares según modo seleccionado
        pares = []
        if url_mode_var.get() == "solo_forms":
            pares = [("", url) for url in lines]
        else:
            if len(lines) % 2 != 0:
                status_var.set(
                    f"En modo Landing+Form las URLs deben ser pares. "
                    f"Tenés {len(lines)} línea(s) — falta una URL."
                )
                status_lbl.config(fg="#f4a261")
                return
            pares = [(lines[i], lines[i + 1]) for i in range(0, len(lines), 2)]

        columnas = build_excel_columns_for_country(pais)

        # Para Brasil: recargar config en el momento de generar
        _br_docs_actual = cargar_config_global().get("lambdatest", {}).get("brasil_docs", {})
        _br_cpf_rows  = set(_br_docs_actual.get("cpf_rows",  []))
        _br_cep_rows  = set(_br_docs_actual.get("cep_rows",  []))
        _br_cnpj_rows = set(_br_docs_actual.get("cnpj_rows", []))

        filas = []
        _form_counter = 0
        for landing_url, form_url in pares:
            _form_counter += 1
            datos = generar_fila_datos(pais)
            if pais == "Brasil":
                from utils.data_generator import generar_documento_brasil_4devs
                if _form_counter in _br_cnpj_rows:
                    _tipo_doc = "cnpj"
                elif _form_counter in _br_cep_rows:
                    _tipo_doc = "cep"
                else:
                    _tipo_doc = "cpf"
                datos["Documento"] = generar_documento_brasil_4devs(_tipo_doc)
            fila = []
            for col in columnas:
                if col == "URL":
                    fila.append(landing_url)
                elif col == "Formulario":
                    fila.append(form_url)
                else:
                    fila.append(datos.get(col, ""))
            filas.append(fila)

        if not filas:
            status_var.set("No se generaron filas.")
            status_lbl.config(fg="#f4a261")
            return

        if pais == "Chile" and "Documento" in columnas:
            from utils.data_generator import generar_rut_chile_con_k
            doc_col_idx = columnas.index("Documento")
            filas[0][doc_col_idx] = generar_rut_chile_con_k()

        excel_nombre = f"Lead_information_Formulario_{pais}_Main.xlsx"
        try:
            _asegurar_paths_modulos()
            from core.country_configs import COUNTRY_CONFIGS
            excel_nombre = COUNTRY_CONFIGS.get(pais, {}).get("excel_file", excel_nombre)
        except Exception:
            pass

        excel_ruta = os.path.join(DATA_DIR, excel_nombre)
        df_new = pd.DataFrame(filas, columns=columnas).astype(str)

        if os.path.exists(excel_ruta):
            reemplazar = messagebox.askyesno(
                "Excel existente",
                f"Ya existe un Excel para {pais}:\n{excel_nombre}\n\n¿Reemplazarlo con los nuevos datos?"
            )
            if not reemplazar:
                status_var.set("Generación cancelada.")
                status_lbl.config(fg="#f4a261")
                return

        try:
            with pd.ExcelWriter(excel_ruta, engine="openpyxl") as writer:
                df_new.to_excel(writer, index=False, sheet_name="Sheet1")
            status_var.set(f"✓ {len(filas)} fila(s) generada(s) → {excel_ruta}")
            status_lbl.config(fg="#a8e6a3")
            _ultima_ruta[0] = excel_ruta
            btn_abrir_carpeta.pack(anchor="w", pady=(6, 0))
        except PermissionError:
            status_var.set("Cerrá el archivo Excel y volvé a intentar.")
            status_lbl.config(fg="#f4a261")
        except Exception as exc:
            status_var.set(f"Error: {exc}")
            status_lbl.config(fg="#f4a261")

    def _limpiar():
        import pandas as pd

        pais = pais_var.get()
        if not pais:
            status_var.set("Seleccioná un país primero.")
            status_lbl.config(fg="#f4a261")
            return
        excel_nombre = f"Lead_information_Formulario_{pais}_Main.xlsx"
        try:
            _asegurar_paths_modulos()
            from core.country_configs import COUNTRY_CONFIGS
            excel_nombre = COUNTRY_CONFIGS.get(pais, {}).get("excel_file", excel_nombre)
        except Exception:
            pass
        excel_ruta = os.path.join(DATA_DIR, excel_nombre)
        if not os.path.exists(excel_ruta):
            status_var.set("No existe el archivo para limpiar.")
            status_lbl.config(fg="#f4a261")
            return
        if not messagebox.askyesno("Limpiar Excel", f"Borrar todas las filas de datos de:\n{excel_nombre}?"):
            return
        columnas = build_excel_columns_for_country(pais)
        try:
            with pd.ExcelWriter(excel_ruta, engine="openpyxl") as writer:
                pd.DataFrame(columns=columnas).to_excel(writer, index=False, sheet_name="Sheet1")
            status_var.set(f"✓ Excel limpiado: {excel_nombre}")
            status_lbl.config(fg="#a8e6a3")
        except PermissionError:
            status_var.set("Cerrá el archivo Excel antes de limpiar.")
            status_lbl.config(fg="#f4a261")

    def _borrar_urls():
        url_text.delete("1.0", END)
        url_country_warning_var.set("")
        status_var.set("URLs borradas.")
        status_lbl.config(fg="white")

    def _generar_con_urls_existentes():
        import pandas as pd

        pais = pais_var.get()
        if not pais:
            status_var.set("Seleccioná un país.")
            status_lbl.config(fg="#f4a261")
            return

        excel_nombre = f"Lead_information_Formulario_{pais}_Main.xlsx"
        try:
            _asegurar_paths_modulos()
            from core.country_configs import COUNTRY_CONFIGS
            excel_nombre = COUNTRY_CONFIGS.get(pais, {}).get("excel_file", excel_nombre)
        except Exception:
            pass

        excel_ruta = os.path.join(DATA_DIR, excel_nombre)
        if not os.path.exists(excel_ruta):
            status_var.set("No existe el Excel para este país. Generá uno primero con URLs.")
            status_lbl.config(fg="#f4a261")
            return

        try:
            df_existente = pd.read_excel(excel_ruta, dtype=str, keep_default_na=False)
        except Exception as exc:
            status_var.set(f"Error al leer el Excel: {exc}")
            status_lbl.config(fg="#f4a261")
            return

        col_url = next((c for c in df_existente.columns if c.strip().lower() == "url"), None)
        col_form = next((c for c in df_existente.columns if c.strip().lower() == "formulario"), None)
        n_filas = len(df_existente)

        if n_filas == 0:
            status_var.set("El Excel no tiene filas. Ingresá URLs primero.")
            status_lbl.config(fg="#f4a261")
            return

        ok = messagebox.askyesno(
            "Regenerar datos",
            f"Se regenerarán los datos de {n_filas} fila(s) para {pais}\n"
            f"manteniendo las URLs existentes en el Excel.\n\n"
            f"¿Continuar?"
        )
        if not ok:
            return

        columnas = build_excel_columns_for_country(pais)
        _br_docs_actual = cargar_config_global().get("lambdatest", {}).get("brasil_docs", {})
        _br_cpf_rows  = set(_br_docs_actual.get("cpf_rows",  []))
        _br_cep_rows  = set(_br_docs_actual.get("cep_rows",  []))
        _br_cnpj_rows = set(_br_docs_actual.get("cnpj_rows", []))

        filas = []
        for i, (_, row) in enumerate(df_existente.iterrows(), start=1):
            datos = generar_fila_datos(pais)
            if pais == "Brasil":
                from utils.data_generator import generar_documento_brasil_4devs
                if i in _br_cnpj_rows:
                    _tipo_doc = "cnpj"
                elif i in _br_cep_rows:
                    _tipo_doc = "cep"
                else:
                    _tipo_doc = "cpf"
                datos["Documento"] = generar_documento_brasil_4devs(_tipo_doc)
            fila = []
            for col in columnas:
                if col == "URL":
                    fila.append(row.get(col_url, "") if col_url else "")
                elif col == "Formulario":
                    fila.append(row.get(col_form, "") if col_form else "")
                else:
                    fila.append(datos.get(col, ""))
            filas.append(fila)

        if pais == "Chile" and "Documento" in columnas:
            from utils.data_generator import generar_rut_chile_con_k
            doc_col_idx = columnas.index("Documento")
            filas[0][doc_col_idx] = generar_rut_chile_con_k()

        try:
            df_new = pd.DataFrame(filas, columns=columnas).astype(str)
            with pd.ExcelWriter(excel_ruta, engine="openpyxl") as writer:
                df_new.to_excel(writer, index=False, sheet_name="Sheet1")
            status_var.set(f"✓ {len(filas)} fila(s) regeneradas con URLs existentes → {excel_ruta}")
            status_lbl.config(fg="#a8e6a3")
            _ultima_ruta[0] = excel_ruta
            btn_abrir_carpeta.pack(anchor="w", pady=(6, 0))
        except PermissionError:
            status_var.set("Cerrá el archivo Excel y volvé a intentar.")
            status_lbl.config(fg="#f4a261")
        except Exception as exc:
            status_var.set(f"Error: {exc}")
            status_lbl.config(fg="#f4a261")

    ttk.Button(frame_btns, text="Generar Excel", command=_generar,
               style="Section.TButton").pack(side=LEFT, padx=(0, 8))
    ttk.Button(frame_btns, text="Regenerar datos (URLs actuales)", command=_generar_con_urls_existentes,
               style="Section.TButton").pack(side=LEFT, padx=(0, 8))
    ttk.Button(frame_btns, text="Limpiar Excel", command=_limpiar,
               style="Section.TButton").pack(side=LEFT, padx=(0, 8))
    ttk.Button(frame_btns, text="Borrar URLs", command=_borrar_urls,
               style="Section.TButton").pack(side=LEFT)

    # ── Config Brasil (tipo de documento) — visible solo cuando se selecciona Brasil ──
    _br_cfg_global = cargar_config_global()
    brasil_cfg = _br_cfg_global.get("lambdatest", {}).get("brasil_docs", {})

    frame_brasil_outer = Frame(frame, bg=APP_BG_COLOR)

    Label(frame_brasil_outer, text="Brasil — tipo de documento a generar:",
          bg=APP_BG_COLOR, fg="white",
          font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(8, 2))
    Label(frame_brasil_outer,
          text="Si no se marca ningún tipo, se genera CPF para todos los forms.",
          font=("Segoe UI", 8, "italic"), bg=APP_BG_COLOR, fg="#aaa").pack(anchor="w", pady=(0, 4))

    frame_brasil = Frame(frame_brasil_outer, bg=APP_BG_COLOR)
    frame_brasil.pack(anchor="w", pady=(0, 4))

    br_cpf_var  = BooleanVar(value=bool(brasil_cfg.get("cpf_rows")))
    br_cep_var  = BooleanVar(value=bool(brasil_cfg.get("cep_rows")))
    br_cnpj_var = BooleanVar(value=bool(brasil_cfg.get("cnpj_rows")))

    frame_br_checks = Frame(frame_brasil, bg=APP_BG_COLOR)
    frame_br_checks.pack(anchor="w")
    Checkbutton(frame_br_checks, text="CPF",  variable=br_cpf_var,  bg=APP_BG_COLOR, fg="white",
                selectcolor=APP_BG_COLOR, activebackground=APP_BG_COLOR, activeforeground="white",
                font=("Segoe UI", 9)).pack(side=LEFT, padx=(0, 12))
    Checkbutton(frame_br_checks, text="CEP",  variable=br_cep_var,  bg=APP_BG_COLOR, fg="white",
                selectcolor=APP_BG_COLOR, activebackground=APP_BG_COLOR, activeforeground="white",
                font=("Segoe UI", 9)).pack(side=LEFT, padx=(0, 12))
    Checkbutton(frame_br_checks, text="CNPJ", variable=br_cnpj_var, bg=APP_BG_COLOR, fg="white",
                selectcolor=APP_BG_COLOR, activebackground=APP_BG_COLOR, activeforeground="white",
                font=("Segoe UI", 9)).pack(side=LEFT)

    frame_br_rows = Frame(frame_brasil, bg=APP_BG_COLOR)
    frame_br_rows.pack(anchor="w", pady=(4, 0))

    def _br_row_gen(parent, label, default):
        f = Frame(parent, bg=APP_BG_COLOR)
        f.pack(anchor="w", pady=1)
        Label(f, text=label, bg=APP_BG_COLOR, fg="white",
              font=("Segoe UI", 9), width=16, anchor="w").pack(side=LEFT)
        v = StringVar(value=default)
        Entry(f, textvariable=v, width=24, font=("Segoe UI", 9)).pack(side=LEFT)
        return v

    br_cpf_rows_var  = _br_row_gen(frame_br_rows, "Filas CPF  (ej: 1,2):",  ",".join(str(x) for x in brasil_cfg.get("cpf_rows",  [])))
    br_cep_rows_var  = _br_row_gen(frame_br_rows, "Filas CEP  (ej: 3):",    ",".join(str(x) for x in brasil_cfg.get("cep_rows",  [])))
    br_cnpj_rows_var = _br_row_gen(frame_br_rows, "Filas CNPJ (ej: 4,5):",  ",".join(str(x) for x in brasil_cfg.get("cnpj_rows", [])))

    def _parse_br_rows(s):
        result = []
        for tok in s.split(","):
            tok = tok.strip()
            if tok.isdigit():
                result.append(int(tok))
        return result

    def _guardar_brasil_cfg():
        cfg_actual = cargar_config_global()
        cfg_actual.setdefault("lambdatest", {})["brasil_docs"] = {
            "cpf_rows":  _parse_br_rows(br_cpf_rows_var.get())  if br_cpf_var.get()  else [],
            "cep_rows":  _parse_br_rows(br_cep_rows_var.get())  if br_cep_var.get()  else [],
            "cnpj_rows": _parse_br_rows(br_cnpj_rows_var.get()) if br_cnpj_var.get() else [],
        }
        guardar_config_global(cfg_actual)
        print("✓ Config Brasil guardada.")

    ttk.Button(frame_brasil, text="Guardar config Brasil",
               command=_guardar_brasil_cfg,
               style="Section.TButton").pack(anchor="w", pady=(8, 0))

    # ── Leyenda de reglas ────────────────────────────────────────────────────
    frame_reglas = Frame(frame, bg=SECTION_CONTAINER_BG_COLOR, padx=2, pady=2)
    frame_reglas.pack(fill="x", pady=(16, 0))

    def _toggle_brasil_panel(*_):
        if pais_var.get() == "Brasil":
            frame_brasil_outer.pack(anchor="w", pady=(8, 4), before=frame_reglas)
        else:
            frame_brasil_outer.pack_forget()

    pais_var.trace_add("write", _toggle_brasil_panel)
    _toggle_brasil_panel()  # estado inicial
    frame_reglas_inner = Frame(frame_reglas, bg=SECTION_BG_COLOR, padx=12, pady=10)
    frame_reglas_inner.pack(fill="both")
    Label(frame_reglas_inner, text="Reglas de generación",
          font=("Segoe UI", 9, "bold"), bg=SECTION_BG_COLOR, fg="white").pack(anchor="w")
    reglas = (
        "• Modelo → vacío (selección aleatoria en el formulario)\n"
        "• Nombre/Apellido → 1 o 2 palabras aleatorias variadas\n"
        "• Documento → Chile: RUT válido  ·  Ecuador: CI válida  ·  Brasil: CPF válido  ·  Resto: número aleatorio\n"
        "• Celular → prefijo y cantidad de dígitos válidos por país\n"
        "  Bolivia: 6/7+7  ·  Perú/Chile: 9+8  ·  Colombia: 3+9  ·  Ecuador/Paraguay: 09+8  ·  Uruguay: 09+7\n"
        "• Región / Ciudad / Concesionario / Fecha → vacíos (selección aleatoria en el formulario)"
    )
    Label(frame_reglas_inner, text=reglas, font=("Segoe UI", 8), bg=SECTION_BG_COLOR,
          fg="#ddd", justify="left").pack(anchor="w", pady=(4, 0))

    Label(frame, text="Added By Elian", font=("Segoe UI", 7, "italic"),
          bg=APP_BG_COLOR, fg="#7a5a95", anchor="e").pack(fill="x", pady=(12, 2))


def _build_lambdatest_tab(parent):
    """Construye el tab de LambdaTest con sub-tabs Mac y Android."""

    frame = Frame(parent, bg=APP_BG_COLOR, padx=20, pady=8)
    frame.pack(fill="both", expand=True)

    Label(frame, text="LambdaTest", font=("Segoe UI", 12, "bold"),
          bg=APP_BG_COLOR, fg="white").pack(anchor="w", pady=(0, 2))
    Label(frame, text="Ejecutá formularios en la nube de LambdaTest.",
          font=("Segoe UI", 9), bg=APP_BG_COLOR, fg="#ddd").pack(anchor="w", pady=(0, 6))

    # ── Credenciales compartidas ───────────────────────────────────────────────
    def _leer_creds_txt():
        path = os.path.join(BASE_DIR, "lambdatest_credentials.txt")
        u = ak = ""
        if os.path.exists(path):
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
        return u, ak

    frame_creds = Frame(frame, bg=APP_BG_COLOR)
    frame_creds.pack(anchor="w", pady=(0, 12))
    Label(frame_creds, text="Credenciales:", bg=APP_BG_COLOR, fg="white",
          font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 2))
    Label(frame_creds, text="Se guardan en lambdatest_credentials.txt (compartidas para Mac y Android)",
          bg=APP_BG_COLOR, fg="#aaa", font=("Segoe UI", 8, "italic")).pack(anchor="w", pady=(0, 4))

    _u, _ak = _leer_creds_txt()

    frame_user = Frame(frame_creds, bg=APP_BG_COLOR)
    frame_user.pack(anchor="w", pady=2)
    Label(frame_user, text="Username:", bg=APP_BG_COLOR, fg="white", width=12, anchor="w").pack(side=LEFT)
    lt_username_var = StringVar(value=_u)
    Entry(frame_user, textvariable=lt_username_var, width=34).pack(side=LEFT)

    frame_key = Frame(frame_creds, bg=APP_BG_COLOR)
    frame_key.pack(anchor="w", pady=2)
    Label(frame_key, text="Access Key:", bg=APP_BG_COLOR, fg="white", width=12, anchor="w").pack(side=LEFT)
    lt_key_var = StringVar(value=_ak)
    Entry(frame_key, textvariable=lt_key_var, width=34, show="*").pack(side=LEFT)

    def _guardar_creds():
        path = os.path.join(BASE_DIR, "lambdatest_credentials.txt")
        with open(path, "w", encoding="utf-8") as f:
            f.write(f"username={lt_username_var.get().strip()}\n")
            f.write(f"access_key={lt_key_var.get().strip()}\n")
        print(f"✓ Credenciales LambdaTest guardadas en {path}")
        messagebox.showinfo("LambdaTest", f"Credenciales guardadas en:\n{path}")

    ttk.Button(frame_creds, text="Guardar credenciales", command=_guardar_creds,
               style="Section.TButton").pack(anchor="w", pady=(8, 0))

    # ── Sub-tabs Mac / Android ─────────────────────────────────────────────────
    inner_nb = ttk.Notebook(frame)
    inner_nb.pack(fill="both", expand=True, pady=(8, 0))

    tab_mac     = Frame(inner_nb, bg=APP_BG_COLOR, padx=12, pady=6)
    tab_android = Frame(inner_nb, bg=APP_BG_COLOR, padx=12, pady=6)
    inner_nb.add(tab_mac,     text="  Mac  ")
    inner_nb.add(tab_android, text="  Android  ")

    # ── Sub-tab MAC ────────────────────────────────────────────────────────────
    Label(tab_mac, text="Mac + Safari", font=("Segoe UI", 11, "bold"),
          bg=APP_BG_COLOR, fg="white").pack(anchor="w", pady=(0, 4))
    Label(tab_mac, text="macOS Sonoma · Safari · 1920×1080",
          font=("Segoe UI", 9), bg=APP_BG_COLOR, fg="#aaa").pack(anchor="w", pady=(0, 10))

    Label(tab_mac, text="Países a ejecutar:", bg=APP_BG_COLOR, fg="white",
          font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 4))
    frame_paises_mac = Frame(tab_mac, bg=APP_BG_COLOR)
    frame_paises_mac.pack(anchor="w", pady=(0, 12))
    lt_mac_vars = {}
    for idx, pais in enumerate(MAPEO_PAISES.keys()):
        var = BooleanVar(value=False)
        lt_mac_vars[pais] = var
        Checkbutton(frame_paises_mac, text=pais, variable=var,
                    bg=APP_BG_COLOR, fg="white", selectcolor=APP_BG_COLOR,
                    activebackground=APP_BG_COLOR, activeforeground="white",
                    font=("Segoe UI", 9)).grid(row=idx // 5, column=idx % 5, sticky="w", padx=6, pady=2)

    frame_btns_mac = Frame(tab_mac, bg=APP_BG_COLOR)
    frame_btns_mac.pack(anchor="w", pady=(8, 0))

    def _ejecutar_lt_mac():
        if _run_state["running"]:
            messagebox.showwarning("En ejecución", "Ya hay una ejecución en curso.")
            return
        paises_sel = [p for p, v in lt_mac_vars.items() if v.get()]
        if not paises_sel:
            messagebox.showwarning("LambdaTest Mac", "Seleccioná al menos un país.")
            return

        cfg_actual = cargar_config_global()
        cfg_actual["lambdatest"] = {
            "username": lt_username_var.get().strip(),
            "access_key": lt_key_var.get().strip(),
        }
        guardar_config_global(cfg_actual)

        stop_ev = threading.Event()
        _run_state["stop_event"] = stop_ev
        _set_running(True)

        import sys as _sys
        _lt_dir = os.path.join(BASE_DIR, "lambdatest_mac")
        if _lt_dir not in _sys.path:
            _sys.path.insert(0, _lt_dir)

        remaining = [len(paises_sel)]
        remaining_lock = threading.Lock()

        def _lt_log(msg):
            msg = str(msg).strip()
            if not msg:
                return
            if (
                "LEAD " in msg
                or msg.lstrip().startswith("→")
                or msg.lstrip().startswith("✗")
                or msg.lstrip().startswith("⛔")
                or msg.lstrip().startswith("⚠ Error")
            ):
                print(msg.strip())

        def _run_pais_mac(pais):
            try:
                import lt_controller  # type: ignore[import]
                summary = lt_controller.run(
                    pais=pais,
                    platform="mac",
                    log_fn=_lt_log,
                    stop_event=stop_ev,
                )
                if summary.get("error"):
                    print(f"✗ {pais}: {summary['error']}")
                else:
                    ok = summary.get("ok", 0)
                    total = summary.get("total", 0)
                    print(f"✓ Leads de {pais} enviados ({ok}/{total})")
            except Exception as e:
                print(f"✗ Error Mac [{pais}]: {e}")
            finally:
                with remaining_lock:
                    remaining[0] -= 1
                    if remaining[0] == 0:
                        _set_running(False)

        print(f"▶ Enviando leads de: {', '.join(paises_sel)}")
        for pais in paises_sel:
            threading.Thread(target=_run_pais_mac, args=(pais,), daemon=True).start()

    btn_lt_mac = ttk.Button(frame_btns_mac, text="Ejecutar en Mac",
                            command=_ejecutar_lt_mac, style="Section.TButton")
    btn_lt_mac.pack(side=LEFT, pady=(4, 0))
    _run_state["lt_btn"] = btn_lt_mac

    Label(tab_mac, text="Los logs aparecen en la consola global (parte inferior).",
          font=("Segoe UI", 8, "italic"), bg=APP_BG_COLOR, fg="#aaa").pack(anchor="w", pady=(10, 0))

    # ── Sub-tab ANDROID ────────────────────────────────────────────────────────
    Label(tab_android, text="Android + Chrome", font=("Segoe UI", 11, "bold"),
          bg=APP_BG_COLOR, fg="white").pack(anchor="w", pady=(0, 4))

    frame_device = Frame(tab_android, bg=APP_BG_COLOR)
    frame_device.pack(anchor="w", pady=(0, 10))
    Label(frame_device, text="Dispositivo:", bg=APP_BG_COLOR, fg="white",
          font=("Segoe UI", 10, "bold"), width=12, anchor="w").pack(side=LEFT)
    ANDROID_DEVICES = ["Galaxy S23", "Galaxy S24", "Galaxy S22", "Galaxy S21", "Galaxy A54"]
    lt_android_device_var = StringVar(value=ANDROID_DEVICES[0])
    ttk.Combobox(frame_device, textvariable=lt_android_device_var,
                 values=ANDROID_DEVICES, state="readonly", width=22).pack(side=LEFT)

    frame_screenshots_android = Frame(tab_android, bg=APP_BG_COLOR)
    frame_screenshots_android.pack(anchor="w", pady=(0, 10))
    lt_android_screenshots_var = BooleanVar(value=False)
    Radiobutton(frame_screenshots_android, text="Sin capturas  ", variable=lt_android_screenshots_var,
                value=False, bg=APP_BG_COLOR, fg="white", selectcolor=APP_BG_COLOR,
                activebackground=APP_BG_COLOR, activeforeground="white",
                font=("Segoe UI", 9)).pack(side=LEFT)
    Radiobutton(frame_screenshots_android, text="Con capturas", variable=lt_android_screenshots_var,
                value=True, bg=APP_BG_COLOR, fg="white", selectcolor=APP_BG_COLOR,
                activebackground=APP_BG_COLOR, activeforeground="white",
                font=("Segoe UI", 9)).pack(side=LEFT)

    Label(tab_android, text="Países a ejecutar:", bg=APP_BG_COLOR, fg="white",
          font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 4))
    frame_paises_android = Frame(tab_android, bg=APP_BG_COLOR)
    frame_paises_android.pack(anchor="w", pady=(0, 12))
    lt_android_vars = {}
    for idx, pais in enumerate(MAPEO_PAISES.keys()):
        var = BooleanVar(value=False)
        lt_android_vars[pais] = var
        Checkbutton(frame_paises_android, text=pais, variable=var,
                    bg=APP_BG_COLOR, fg="white", selectcolor=APP_BG_COLOR,
                    activebackground=APP_BG_COLOR, activeforeground="white",
                    font=("Segoe UI", 9)).grid(row=idx // 5, column=idx % 5, sticky="w", padx=6, pady=2)

    frame_btns_android = Frame(tab_android, bg=APP_BG_COLOR)
    frame_btns_android.pack(anchor="w", pady=(8, 0))

    def _ejecutar_lt_android():
        if _run_state["running"]:
            messagebox.showwarning("En ejecución", "Ya hay una ejecución en curso.")
            return
        paises_sel = [p for p, v in lt_android_vars.items() if v.get()]
        if not paises_sel:
            messagebox.showwarning("LambdaTest Android", "Seleccioná al menos un país.")
            return

        cfg_actual = cargar_config_global()
        cfg_actual["lambdatest"] = {
            "username": lt_username_var.get().strip(),
            "access_key": lt_key_var.get().strip(),
        }
        guardar_config_global(cfg_actual)

        stop_ev = threading.Event()
        _run_state["stop_event"] = stop_ev
        _set_running(True)

        import sys as _sys
        _lt_android_dir = os.path.join(BASE_DIR, "lambdatest_android")
        _lt_mac_dir     = os.path.join(BASE_DIR, "lambdatest_mac")
        for _d in (_lt_android_dir, _lt_mac_dir):
            if _d not in _sys.path:
                _sys.path.insert(0, _d)

        device_name = lt_android_device_var.get()
        remaining = [len(paises_sel)]
        remaining_lock = threading.Lock()

        def _lt_log_android(msg):
            msg = str(msg).strip()
            if not msg:
                return
            if (
                "LEAD " in msg
                or msg.lstrip().startswith("→")
                or msg.lstrip().startswith("✗")
                or msg.lstrip().startswith("⛔")
                or msg.lstrip().startswith("⚠ Error")
            ):
                print(msg.strip())

        def _run_pais_android(pais):
            try:
                import lt_android_controller  # type: ignore[import]
                summary = lt_android_controller.run(
                    pais=pais,
                    device_name=device_name,
                    with_screenshots=lt_android_screenshots_var.get(),
                    log_fn=_lt_log_android,
                    stop_event=stop_ev,
                )
                if summary.get("error"):
                    print(f"✗ {pais}: {summary['error']}")
                else:
                    ok = summary.get("ok", 0)
                    total = summary.get("total", 0)
                    print(f"✓ Leads de {pais} enviados ({ok}/{total})")
            except Exception as e:
                print(f"✗ Error Android [{pais}]: {e}")
            finally:
                with remaining_lock:
                    remaining[0] -= 1
                    if remaining[0] == 0:
                        _set_running(False)

        print(f"▶ Enviando leads de: {', '.join(paises_sel)} ({device_name})")
        for pais in paises_sel:
            threading.Thread(target=_run_pais_android, args=(pais,), daemon=True).start()

    btn_lt_android = ttk.Button(frame_btns_android, text="Ejecutar en Android",
                                command=_ejecutar_lt_android, style="Section.TButton")
    btn_lt_android.pack(side=LEFT, pady=(4, 0))

    Label(tab_android, text="Los logs aparecen en la consola global (parte inferior).",
          font=("Segoe UI", 8, "italic"), bg=APP_BG_COLOR, fg="#aaa").pack(anchor="w", pady=(10, 0))

    Label(frame, text="Added By Elian", font=("Segoe UI", 7, "italic"),
          bg=APP_BG_COLOR, fg="#7a5a95", anchor="e").pack(fill="x", pady=(8, 2))


def iniciar_interfaz():
    limpiar_temporales()
    root = Tk()
    root.title("Osocio - Form Automation")
    # Tamaño optimizado: 980x768
    root.geometry("1200x800")
    root.minsize(1100, 700)
    root.configure(bg=APP_BG_COLOR)

    # Uniformar estilos de tabs y botones para respetar la paleta definida
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass

    style.configure(
        "TNotebook",
        background=APP_BG_COLOR,
        borderwidth=0,
    )
    style.configure(
        "TNotebook.Tab",
        background=HEADER_BG_COLOR,
        foreground="white",
        padding=(8, 4),
        font=("Segoe UI", 10, "bold"),
    )
    style.map(
        "TNotebook.Tab",
        background=[("selected", HEADER_BG_COLOR), ("active", "#b897d8")],
        foreground=[("selected", "white"), ("disabled", "#e6dcf2")],
        padding=[("selected", (16, 8))],
    )

    style.configure(
        "TButton",
        background=HEADER_BG_COLOR,
        foreground="white",
        borderwidth=0,
        padding=(10, 6),
    )
    style.map(
        "TButton",
        background=[
            ("active", "#b897d8"),
            ("pressed", "#7f5798"),
            ("disabled", "#bda7d6"),
        ],
        foreground=[("disabled", "#f3eff8")],
    )

    style.configure(
        "Section.TNotebook",
        background=SECTION_BG_COLOR,
        borderwidth=0,
    )
    style.configure(
        "Section.TNotebook.Tab",
        background=SECTION_CTA_BG_COLOR,
        foreground=PRIMARY_TEXT_COLOR,
        padding=(8, 4),
    )
    style.map(
        "Section.TNotebook.Tab",
        background=[("selected", "#6990BA"), ("active", "#4a6a8d")],
        foreground=[("selected", PRIMARY_TEXT_COLOR), ("disabled", "#4a4a4a")],
        padding=[("selected", (16, 8))],
    )

    style.configure(
        "Section.TButton",
        background="#6990BA",
        foreground=PRIMARY_TEXT_COLOR,
        borderwidth=0,
        padding=(10, 6),
        relief="flat",
    )
    style.map(
        "Section.TButton",
        background=[
            ("active", "#4a6a8d"),
            ("pressed", "#2c435b"),
            ("disabled", "#9cb6d0"),
        ],
        foreground=[("disabled", "#4a4a4a")],
    )

    style.configure(
        "FolderCTA.TButton",
        background="#2E8B57",
        foreground="white",
        font=("Segoe UI", 9, "bold"),
        borderwidth=0,
        padding=(10, 6),
        relief="flat",
    )
    style.map(
        "FolderCTA.TButton",
        background=[("active", "#246B43"), ("pressed", "#1A5230"), ("disabled", "#7ab89a")],
        foreground=[("disabled", "#ccc")],
    )

    style.configure(
        "Section.Vertical.TScrollbar",
        troughcolor="#6b4890",
        background="#1a0d2e",
        bordercolor="#1a0d2e",
        arrowcolor="#c8a0e8",
        gripcount=0,
        relief="flat",
    )
    style.map(
        "Section.Vertical.TScrollbar",
        background=[("active", "#2d1b44"), ("pressed", "#0f0820")],
    )
    style.configure(
        "Section.Horizontal.TScrollbar",
        troughcolor="#6b4890",
        background="#1a0d2e",
        bordercolor="#1a0d2e",
        arrowcolor="#c8a0e8",
        gripcount=0,
        relief="flat",
    )
    style.map(
        "Section.Horizontal.TScrollbar",
        background=[("active", "#2d1b44"), ("pressed", "#0f0820")],
    )

    style.configure(
        "Results.TButton",
        background="#6990BA",
        foreground=PRIMARY_TEXT_COLOR,
        borderwidth=0,
        padding=(10, 4),
        relief="flat",
    )
    style.map(
        "Results.TButton",
        background=[
            ("active", "#4a6a8d"),
            ("pressed", "#2c435b"),
            ("disabled", "#9cb6d0"),
        ],
        foreground=[("disabled", "#4a4a4a")],
    )

    style.configure(
        "EnviarLeads.TButton",
        background="white",
        foreground="black",
        font=("Segoe UI", 11, "bold"),
        borderwidth=0,
        padding=(14, 8),
        relief="flat",
    )
    style.map(
        "EnviarLeads.TButton",
        background=[
            ("active", "#e0e0e0"),
            ("pressed", "#c0c0c0"),
            ("disabled", "#cccccc"),
        ],
        foreground=[("disabled", "#888888")],
    )

    style.configure(
        "HeaderSeparator.TSeparator",
        background=HEADER_BG_COLOR,
        bordercolor=HEADER_BG_COLOR,
        darkcolor=HEADER_BG_COLOR,
        lightcolor=HEADER_BG_COLOR,
    )
    style.configure(
        "App.TSeparator",
        background=APP_BG_COLOR,
        bordercolor=APP_BG_COLOR,
        darkcolor=APP_BG_COLOR,
        lightcolor=APP_BG_COLOR,
    )
    style.configure(
        "Section.TSeparator",
        background=SECTION_BG_COLOR,
        bordercolor=SECTION_BG_COLOR,
        darkcolor=SECTION_BG_COLOR,
        lightcolor=SECTION_BG_COLOR,
    )

    icon_path = os.path.join(ASSET_DIR, "icon.ico")
    if os.path.exists(icon_path):
        try:
            root.iconbitmap(icon_path)
        except Exception:
            pass

    # === CABECERA FIJA ===
    _fullheader_path = os.path.join(ASSET_DIR, "Fullheader.png")
    _fullheader_orig = Image.open(_fullheader_path) if os.path.exists(_fullheader_path) else None
    _fh_orig_w, _fh_orig_h = (_fullheader_orig.size if _fullheader_orig else (900, 120))

    _FH_HEIGHT = 200
    _FH_WIDTH = max(1, int(round((_fh_orig_w * _FH_HEIGHT) / _fh_orig_h)))
    _HEADER_SPLIT_LEFT_COLOR = "#110830"
    _HEADER_SPLIT_RIGHT_COLOR = "#28164B"
    header_canvas = Canvas(root, bd=0, highlightthickness=0, height=_FH_HEIGHT)
    header_canvas.pack(fill="x", side="top")

    _fh_photo_ref = [None]
    if _fullheader_orig is not None:
        _fh_photo_ref[0] = ImageTk.PhotoImage(_fullheader_orig.resize((_FH_WIDTH, _FH_HEIGHT), Image.LANCZOS))

    def _render_fixed_header(_event=None):
        header_canvas.delete("all")
        canvas_w = max(1, header_canvas.winfo_width())
        split_x = canvas_w // 2
        header_canvas.create_rectangle(0, 0, split_x, _FH_HEIGHT, fill=_HEADER_SPLIT_LEFT_COLOR, outline="")
        header_canvas.create_rectangle(split_x, 0, canvas_w, _FH_HEIGHT, fill=_HEADER_SPLIT_RIGHT_COLOR, outline="")
        if _fh_photo_ref[0] is None:
            return
        x = (canvas_w - _FH_WIDTH) // 2
        header_canvas.create_image(x, 0, anchor="nw", image=_fh_photo_ref[0])

    header_canvas.bind("<Configure>", _render_fixed_header)
    root.after(0, _render_fixed_header)
    ttk.Separator(root, orient="horizontal", style="App.TSeparator").pack(fill="x")

    # === PANEL DE CONSOLA DESPLEGABLE (bottom) ===
    _run_state["root"] = root
    _broker = _ManualInputBroker(root)
    _broker.start_polling()
    _run_state["broker"] = _broker
    _console_expanded = [False]   # estado: cerrado por defecto

    # Franja inferior fija (siempre visible) — contiene el toggle y, si está abierto, el panel
    _console_wrapper = Frame(root, bg="#2b1d3a")
    _console_wrapper.pack(fill="x", side="bottom")

    # ── Toggle bar ──────────────────────────────────────────────────────────
    _toggle_bar = Frame(_console_wrapper, bg="#3d2a52", cursor="hand2")
    _toggle_bar.pack(fill="x", padx=20, pady=(2, 0))

    _toggle_lbl = Label(
        _toggle_bar, text="▶  Consola", bg="#3d2a52", fg="#c9b8e8",
        font=("Segoe UI", 9), padx=8, pady=4, cursor="hand2",
    )
    _toggle_lbl.pack(side=LEFT)

    # Botón DETENER (oculto hasta que haya ejecución)
    _stop_btn = Button(
        _toggle_bar, text="⛔ Detener", bg="#c0392b", fg="white",
        font=("Segoe UI", 9, "bold"), relief="flat", bd=0, cursor="hand2",
        command=_request_stop,
        activebackground="#e74c3c", activeforeground="white",
        padx=8, pady=3,
    )
    _run_state["stop_btn"] = _stop_btn

    # ── Panel de texto (oculto por defecto) ─────────────────────────────────
    _console_panel = Frame(_console_wrapper, bg="#3d2a52", padx=1, pady=1)
    # NO se empaca aquí — se muestra/oculta con toggle

    _console_panel_header = Frame(_console_panel, bg="#3d2a52")
    _console_panel_header.pack(fill="x")

    def _clear_console():
        _console_text.configure(state="normal")
        _console_text.delete("1.0", END)
        _console_text.configure(state="disabled")

    Button(_console_panel_header, text="Limpiar", bg="#3d2a52", fg="#aaa",
           font=("Consolas", 8), relief="flat", bd=0, cursor="hand2",
           command=_clear_console,
           activebackground="#3d2a52", activeforeground="white").pack(side=RIGHT, padx=6, pady=2)

    _console_inner = Frame(_console_panel, bg="#1a1a2e")
    _console_inner.pack(fill="both", expand=True, padx=1, pady=(0, 1))
    _console_vscroll = ttk.Scrollbar(_console_inner, orient="vertical",
                                     style="Section.Vertical.TScrollbar")
    _console_vscroll.pack(side=RIGHT, fill="y")
    _console_text = Text(
        _console_inner, height=8, bg="#1a1a2e", fg="#d4d4d4",
        font=("Consolas", 9), relief="flat", bd=0,
        yscrollcommand=_console_vscroll.set,
        state="disabled", wrap="word",
    )
    _console_text.pack(fill="both", expand=True, padx=4, pady=4)
    _console_vscroll.config(command=_console_text.yview)

    def _toggle_console(_event=None):
        if _console_expanded[0]:
            _console_panel.pack_forget()
            _toggle_lbl.configure(text="▶  Consola")
            _console_expanded[0] = False
        else:
            _console_panel.pack(fill="both", expand=True, padx=20, pady=(0, 8))
            _toggle_lbl.configure(text="▼  Consola")
            _console_expanded[0] = True

    _toggle_bar.bind("<Button-1>", _toggle_console)
    _toggle_lbl.bind("<Button-1>", _toggle_console)

    # Redirigir sys.stdout al widget de consola (TeeStream thread-safe)
    _original_stdout = sys.stdout

    class _TeeStream:
        def write(self, text):
            if text:
                try:
                    root.after(0, _append_console, text)
                except Exception:
                    pass
                try:
                    _original_stdout.write(text)
                except Exception:
                    pass
        def flush(self):
            try:
                _original_stdout.flush()
            except Exception:
                pass
        def isatty(self):
            return False

    def _append_console(text):
        try:
            _console_text.configure(state="normal")
            _console_text.insert(END, text)
            _console_text.see(END)
            _console_text.configure(state="disabled")
        except Exception:
            pass

    sys.stdout = _TeeStream()

    # === CONTENEDOR SCROLLEABLE (evita colapso al achicar ventana) ===
    outer_container = Frame(root, bg=APP_BG_COLOR, bd=0, highlightthickness=0)
    outer_container.pack(fill="both", expand=True)

    canvas = Canvas(outer_container, bg=APP_BG_COLOR, bd=0, highlightthickness=0)
    v_scroll = ttk.Scrollbar(outer_container, orient="vertical", style="Section.Vertical.TScrollbar")

    # Configurar grid: canvas (row 0), footer (row 1)
    outer_container.grid_rowconfigure(0, weight=1)  # Canvas se expande
    outer_container.grid_rowconfigure(1, weight=0)  # Footer tiene altura fija
    outer_container.grid_columnconfigure(0, weight=1)
    
    canvas.grid(row=0, column=0, sticky="nsew")
    v_scroll.grid(row=0, column=1, sticky="ns")

    ui_root = Frame(canvas, bg=APP_BG_COLOR, bd=0, highlightthickness=0)
    content_window = canvas.create_window((0, 0), window=ui_root, anchor="nw")

    # Debounce para evitar "estiramientos" al scrollear rápido (muchos <Configure>)
    _scroll_update_job = {"id": None}

    def _update_scroll_region(_event=None):
        try:
            canvas.configure(scrollregion=canvas.bbox("all"))
            _update_main_scrollbar_state()
        except Exception:
            pass

    def _canvas_has_vertical_overflow():
        try:
            bbox = canvas.bbox("all")
            if not bbox:
                return False
            content_height = max(0, bbox[3] - bbox[1])
            viewport_height = max(0, canvas.winfo_height())
            return content_height > (viewport_height + 1)
        except Exception:
            return False

    def _canvas_has_horizontal_overflow():
        try:
            bbox = canvas.bbox("all")
            if not bbox:
                return False
            content_width = max(0, bbox[2] - bbox[0])
            viewport_width = max(0, canvas.winfo_width())
            return content_width > (viewport_width + 1)
        except Exception:
            return False

    def _update_main_scrollbar_state(first=None, last=None):
        try:
            has_overflow = _canvas_has_vertical_overflow()
            if has_overflow:
                v_scroll.grid()
                if first is None or last is None:
                    first, last = canvas.yview()
                first, last = float(first), float(last)
                # Si scrollregion no está seteado aún, yview devuelve (0,1)
                # aunque haya overflow real → calcular thumb desde bbox
                if first == 0.0 and last >= 1.0:
                    try:
                        bbox = canvas.bbox("all")
                        if bbox:
                            content_h = max(bbox[3] - bbox[1], 1)
                            viewport_h = max(canvas.winfo_height(), 1)
                            if content_h > viewport_h:
                                last = viewport_h / content_h
                    except Exception:
                        pass
                v_scroll.set(first, last)
            else:
                v_scroll.grid_remove()
        except Exception:
            pass

    def _on_main_scrollbar(*args):
        if not _canvas_has_vertical_overflow():
            _update_main_scrollbar_state(0.0, 1.0)
            return
        canvas.yview(*args)

    def _schedule_scroll_region_update(_event=None):
        try:
            if _scroll_update_job["id"] is not None:
                canvas.after_cancel(_scroll_update_job["id"])
        except Exception:
            pass
        # ~60 FPS (16ms). Reduce jitter sin perder fluidez.
        _scroll_update_job["id"] = canvas.after(16, _update_scroll_region)

    def _sync_window_width(event):
        try:
            # Fijar el ancho del contenido al ancho visible del canvas para que
            # componentes anchos (ej. Treeview) NO estiren la ventana completa.
            # El Treeview ya tiene su propio scroll horizontal.
            canvas.itemconfigure(content_window, width=event.width)
            _update_main_scrollbar_state()
        except Exception:
            pass

    v_scroll.configure(command=_on_main_scrollbar)
    canvas.configure(yscrollcommand=_update_main_scrollbar_state)

    ui_root.bind("<Configure>", _schedule_scroll_region_update)
    canvas.bind("<Configure>", _sync_window_width)
    # Inicializar scrollregion una vez que se renderiza todo
    canvas.after(0, _schedule_scroll_region_update)

    # === Scroll con rueda del mouse (vertical) y Shift+rueda (horizontal) ===
    def _on_mousewheel(event):
        # Windows/macOS: event.delta (multiplo de 120). Linux: usar Button-4/5 abajo.
        try:
            if getattr(event, "state", 0) & 0x0001:  # Shift presionado
                # horizontal
                if not _canvas_has_horizontal_overflow():
                    return "break"
                delta = int(-1 * (event.delta / 120))
                canvas.xview_scroll(delta, "units")
            else:
                if not _canvas_has_vertical_overflow():
                    return "break"
                delta = int(-1 * (event.delta / 120))
                canvas.yview_scroll(delta, "units")
        except Exception:
            return "break"
        return "break"

    def _on_linux_wheel_up(_event):
        if not _canvas_has_vertical_overflow():
            return "break"
        canvas.yview_scroll(-1, "units")
        return "break"

    def _on_linux_wheel_down(_event):
        if not _canvas_has_vertical_overflow():
            return "break"
        canvas.yview_scroll(1, "units")
        return "break"

    def _bind_mousewheel(_event=None):
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        canvas.bind_all("<Button-4>", _on_linux_wheel_up)
        canvas.bind_all("<Button-5>", _on_linux_wheel_down)

    def _unbind_mousewheel(_event=None):
        try:
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")
        except Exception:
            pass

    canvas.bind("<Enter>", _bind_mousewheel)
    canvas.bind("<Leave>", _unbind_mousewheel)

    # === INICIALIZAR VARIABLES DE CONFIGURACIÓN ===
    global chrome_var, firefox_var, edge_var, desktop_var, mobile_var
    global lt_mac_var, lt_android_var, visible_browser_var, email_modo_var
    global modo_ejecucion_var, pais_single_var

    _ui_prefs = cargar_config_global().get("ui_prefs", {})
    _navs = _ui_prefs.get("navegadores", ["chrome"])
    _vps  = _ui_prefs.get("viewports",   ["fullscreen"])
    chrome_var          = BooleanVar(value="chrome"             in _navs)
    firefox_var         = BooleanVar(value="firefox"            in _navs)
    edge_var            = BooleanVar(value="edge"               in _navs)
    lt_mac_var          = BooleanVar(value="lambdatest_mac"     in _navs)
    lt_android_var      = BooleanVar(value="lambdatest_android" in _navs)
    visible_browser_var = BooleanVar(value=bool(_ui_prefs.get("visible_browser", False)))
    desktop_var         = BooleanVar(value="fullscreen" in _vps)
    mobile_var          = BooleanVar(value="600x738"    in _vps)
    modo_ejecucion_var  = StringVar(value="consecutive")
    pais_single_var     = StringVar(value="")

    email_modo_var = StringVar(value=_ui_prefs.get("email_modo", "por_pais"))

    app_tabs_container = Frame(ui_root, bg=APP_BG_COLOR, bd=0, highlightthickness=0)
    app_tabs_container.pack(fill="both", expand=True, padx=20, pady=(0, 0))

    app_notebook = ttk.Notebook(app_tabs_container, style="TNotebook")
    app_notebook.pack(fill="both", expand=True)

    testing_tab = Frame(app_notebook, bg=APP_BG_COLOR, bd=0, highlightthickness=0)
    validation_tab = Frame(app_notebook, bg=APP_BG_COLOR, bd=0, highlightthickness=0)
    lambdatest_tab = Frame(app_notebook, bg=APP_BG_COLOR, bd=0, highlightthickness=0)
    generar_excels_tab = Frame(app_notebook, bg=APP_BG_COLOR, bd=0, highlightthickness=0)
    app_notebook.add(testing_tab, text="Envio de Leads")
    app_notebook.add(validation_tab, text="Validación de Campos")
    app_notebook.add(lambdatest_tab, text="LambdaTest")
    app_notebook.add(generar_excels_tab, text="Generar Excels con Datos")

    # Configuración global compartida entre tabs
    cfg_global = cargar_config_global()
    cfg_global["enviar_mail"] = False  # siempre arranca desactivado, sin importar sesión anterior
    if "adjuntar_resultados" not in cfg_global:
        cfg_global["adjuntar_resultados"] = True
    if "adjuntar_screenshots" not in cfg_global:
        cfg_global["adjuntar_screenshots"] = True
    guardar_config_global(cfg_global)

    email_var = StringVar(value=obtener_email_destinatario())
    enviar_mail_var = BooleanVar(value=False)  # siempre inicia apagado por seguridad
    adjuntar_resultados_var = BooleanVar(value=bool(cfg_global.get("adjuntar_resultados", True)))
    adjuntar_screenshots_var = BooleanVar(value=bool(cfg_global.get("adjuntar_screenshots", True)))

    build_field_validation_tab(
        validation_tab,
        {
            "app_bg": APP_BG_COLOR,
            "container_bg": SECTION_CONTAINER_BG_COLOR,
            "section_bg": SECTION_BG_COLOR,
            "text_color": PRIMARY_TEXT_COLOR,
            "button_bg": HEADER_BG_COLOR,
            "button_fg": PRIMARY_TEXT_COLOR,
        },
        {
            "browser_vars": {
                "chrome": chrome_var,
                "firefox": firefox_var,
                "edge": edge_var,
            },
            "viewport_vars": {
                "fullscreen": desktop_var,
                "600x738": mobile_var,
            },
            "email_var": email_var,
            "enviar_mail_var": enviar_mail_var,
            "adjuntar_resultados_var": adjuntar_resultados_var,
        },
    )

    _build_lambdatest_tab(lambdatest_tab)
    _build_generar_excels_tab(generar_excels_tab)

    # === FILA SUPERIOR: Configuración Global + Botón Ejecutar Todos ===
    frame_superior = Frame(ui_root, bg=APP_BG_COLOR)
    frame_superior.pack(fill="x", padx=20, pady=(0, 4), before=app_tabs_container)

    # Configuración Global - Título arriba, controles abajo
    # TÍTULO PRINCIPAL arriba contra margen izquierdo
    Label(
        frame_superior,
        text="Configuración Global",
        font=("Segoe UI", 12, "bold"),
        bg=APP_BG_COLOR,
        fg="white",
    ).pack(anchor="w", pady=(0, 6))

    # CONTROLES por debajo del título
    frame_controles = Frame(frame_superior, bg=APP_BG_COLOR)
    frame_controles.pack(fill="x")

    # IZQUIERDA: Navegador y Viewport en columna vertical
    frame_izquierda = Frame(frame_controles, bg=APP_BG_COLOR)
    frame_izquierda.pack(side=LEFT, anchor="n", padx=(0, 20))

    # Navegador (primera fila)
    frame_navegador = Frame(frame_izquierda, bg=APP_BG_COLOR)
    frame_navegador.pack(anchor="w", pady=2)
    Label(
        frame_navegador,
        text="Navegador:",
        bg=APP_BG_COLOR,
        fg="white",
        width=10,
        anchor="w",
    ).pack(side=LEFT)

    Checkbutton(
        frame_navegador,
        text="Chrome",
        variable=chrome_var,
        bg=APP_BG_COLOR,
        fg="white",
        activebackground=APP_BG_COLOR,
        activeforeground="white",
        selectcolor=APP_BG_COLOR,
    ).pack(side=LEFT, padx=5)
    Checkbutton(
        frame_navegador,
        text="Firefox",
        variable=firefox_var,
        bg=APP_BG_COLOR,
        fg="white",
        activebackground=APP_BG_COLOR,
        activeforeground="white",
        selectcolor=APP_BG_COLOR,
    ).pack(side=LEFT, padx=5)
    Checkbutton(
        frame_navegador,
        text="Edge",
        variable=edge_var,
        bg=APP_BG_COLOR,
        fg="white",
        activebackground=APP_BG_COLOR,
        activeforeground="white",
        selectcolor=APP_BG_COLOR,
    ).pack(side=LEFT, padx=5)

    # Viewport
    frame_viewport = Frame(frame_izquierda, bg=APP_BG_COLOR)
    frame_viewport.pack(anchor="w", pady=2)
    Label(
        frame_viewport,
        text="Viewport:",
        bg=APP_BG_COLOR,
        fg="white",
        width=10,
        anchor="w",
    ).pack(side=LEFT)

    Checkbutton(
        frame_viewport,
        text="Desktop",
        variable=desktop_var,
        bg=APP_BG_COLOR,
        fg="white",
        activebackground=APP_BG_COLOR,
        activeforeground="white",
        selectcolor=APP_BG_COLOR,
    ).pack(side=LEFT, padx=5)
    Checkbutton(
        frame_viewport,
        text="Mobile emulado",
        variable=mobile_var,
        bg=APP_BG_COLOR,
        fg="white",
        activebackground=APP_BG_COLOR,
        activeforeground="white",
        selectcolor=APP_BG_COLOR,
    ).pack(side=LEFT, padx=5)

    # Visibilidad del browser
    frame_visibilidad = Frame(frame_izquierda, bg=APP_BG_COLOR)
    frame_visibilidad.pack(anchor="w", pady=2)
    Label(
        frame_visibilidad,
        text="Modo:",
        bg=APP_BG_COLOR,
        fg="white",
        width=10,
        anchor="w",
    ).pack(side=LEFT)
    Checkbutton(
        frame_visibilidad,
        text="Ver navegador mientras corre",
        variable=visible_browser_var,
        bg=APP_BG_COLOR,
        fg="white",
        activebackground=APP_BG_COLOR,
        activeforeground="white",
        selectcolor=APP_BG_COLOR,
    ).pack(side=LEFT, padx=5)
    Label(
        frame_visibilidad,
        text="ℹ️ Por defecto corre en segundo plano sin interrumpirte",
        bg=APP_BG_COLOR,
        fg="#888888",
        font=("Segoe UI", 8),
    ).pack(side=LEFT, padx=4)

    # Email (derecha contra el margen)
    frame_email_derecha = Frame(frame_controles, bg=APP_BG_COLOR)
    frame_email_derecha.pack(side=RIGHT, anchor="e")

    # === Email destinatario (persistente) ===
    frame_email = Frame(frame_email_derecha, bg=APP_BG_COLOR)
    frame_email.pack(anchor="e", pady=2)

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

    # === IDs únicos (contenido de tab dentro del popup unificado) ===
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
            text="Asigná uno o más valores fijos a un ID de campo no mapeado.\nSirve para inputs, textareas y selects especiales. También podés volver a cargar el mismo ID para sumar más valores.",
            font=("Segoe UI", 9),
            bg=APP_BG_COLOR,
            fg="#ddd",
            justify="left",
            anchor="w",
        )
        lbl_descripcion.pack(fill="x", padx=20, anchor="w")
        _registrar_label_responsivo(
            lbl_descripcion,
            "Asigná uno o más valores fijos a un ID de campo no mapeado.\nSirve para inputs, textareas y selects especiales.\nTambién podés volver a cargar el mismo ID para sumar más valores.",
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
        entry_valor = Entry(frame_inputs, font=("Segoe UI", 10), width=24)
        entry_valor.grid(row=2, column=1, padx=(0, 10), pady=4)

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
        label_warning_id.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(2, 0))

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

                texto_busqueda = f"{id_val} {nombre_campo} {valor_texto} {paises_abrev}".lower()
                if texto_filtro and texto_filtro not in texto_busqueda:
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
                    entry_valor.insert(0, " | ".join(valores_actuales))
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
            valor_val = entry_valor.get().strip()
            paises_seleccionados = [pais for pais, var in pais_vars.items() if var.get()]
            if _es_todos_paises(paises_seleccionados):
                paises_seleccionados = []
            if not id_val or not valor_val:
                messagebox.showwarning("IDs únicos", "Completá ambos campos.", parent=popup)
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
        frame_filtros.grid(row=4, column=0, columnspan=2, pady=(2, 0), sticky="w")

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
        frame_ctas_paises.grid(row=4, column=2, pady=(2, 0), sticky="ew")
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

    def _build_tab_dependencias(popup):
        """Tab para registrar dependencias padre→hijo por país, con estilo y estructura igual a IDs Excel."""

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
        frame_lista.bind("<Configure>", lambda _event=None: canvas_lista.configure(scrollregion=canvas_lista.bbox("all")))
        canvas_lista.bind("<Configure>", lambda event: canvas_lista.itemconfigure(lista_window, width=event.width))

        filas_widgets = {}
        dependencias = cargar_dependencias()

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

    def abrir_popup_ids_dinamicos():
        popup = Toplevel(ui_root)
        popup.title("IDs Dinámicos")
        popup.configure(bg=APP_BG_COLOR)
        popup.geometry("930x620")
        popup.minsize(760, 520)
        popup.resizable(True, True)

        icon_path = os.path.join(ASSET_DIR, "icon.ico")
        if os.path.exists(icon_path):
            try:
                popup.iconbitmap(icon_path)
            except Exception:
                try:
                    popup.iconbitmap(default=icon_path)
                except Exception:
                    pass

        notebook_ids = ttk.Notebook(popup, style="TNotebook")
        notebook_ids.pack(fill="both", expand=True, padx=10, pady=10)

        tab_ids_unicos = Frame(notebook_ids, bg=APP_BG_COLOR)
        tab_ids_excel = Frame(notebook_ids, bg=APP_BG_COLOR)
        tab_dependencias = Frame(notebook_ids, bg=APP_BG_COLOR)
        notebook_ids.add(tab_ids_unicos, text="IDs únicos")
        notebook_ids.add(tab_ids_excel, text="IDs Excel")
        notebook_ids.add(tab_dependencias, text="Dependencias")

        _build_tab_ids_unicos(tab_ids_unicos)
        _build_tab_ids_excel(tab_ids_excel)
        _build_tab_dependencias(tab_dependencias)

    frame_btn_dinamicos = Frame(testing_tab, bg=APP_BG_COLOR)
    frame_btn_dinamicos.pack(fill="x", padx=20, pady=(10, 8))
    Button(
        frame_btn_dinamicos,
        text=" IDs Dinámicos",
        command=abrir_popup_ids_dinamicos,
        bg=HEADER_BG_COLOR,
        fg="black",
        relief="flat",
        font=("Segoe UI", 10, "bold"),
        cursor="hand2",
        padx=12,
        pady=4,
    ).pack(anchor="e")

    Label(
        frame_email,
        text="Email:",
        bg=APP_BG_COLOR,
        fg="white",
        width=10,
        anchor="w",
    ).pack(side=LEFT)

    entry_email = Entry(frame_email, font=("Segoe UI", 10), width=50, textvariable=email_var)
    entry_email.pack(side=LEFT, padx=5)
    Label(frame_email, text="(varios emails separados por coma)",
          font=("Segoe UI", 8), bg=APP_BG_COLOR, fg="#888").pack(side=LEFT)

    def _persistir_email_destinatario(*_):
        config = cargar_config_global()
        config["email_destinatario"] = (email_var.get() or "").strip()
        guardar_config_global(config)

    # Guardar cuando se sale del campo o cuando cambia (por ejemplo pegado)
    entry_email.bind("<FocusOut>", _persistir_email_destinatario)
    email_var.trace_add("write", _persistir_email_destinatario)

    # === Opciones de email (persistentes) ===
    frame_email_opts = Frame(frame_email_derecha, bg=APP_BG_COLOR)
    frame_email_opts.pack(anchor="e", pady=(4, 2))

    chk_enviar_mail = Checkbutton(
        frame_email_opts,
        text="Enviar mail",
        variable=enviar_mail_var,
        bg=APP_BG_COLOR,
        fg="white",
        activebackground=APP_BG_COLOR,
        activeforeground="white",
        selectcolor=APP_BG_COLOR,
    )
    chk_enviar_mail.pack(side=LEFT, padx=(0, 10))

    chk_adj_resultados = Checkbutton(
        frame_email_opts,
        text="Adjuntar resultados",
        variable=adjuntar_resultados_var,
        bg=APP_BG_COLOR,
        fg="white",
        activebackground=APP_BG_COLOR,
        activeforeground="white",
        selectcolor=APP_BG_COLOR,
    )
    chk_adj_resultados.pack(side=LEFT, padx=5)

    chk_adj_screens = Checkbutton(
        frame_email_opts,
        text="Adjuntar screenshots",
        variable=adjuntar_screenshots_var,
        bg=APP_BG_COLOR,
        fg="white",
        activebackground=APP_BG_COLOR,
        activeforeground="white",
        selectcolor=APP_BG_COLOR,
    )
    chk_adj_screens.pack(side=LEFT, padx=5)

    # Modo de envío: un email por país vs. uno consolidado al final
    frame_email_modo = Frame(frame_email_derecha, bg=APP_BG_COLOR)
    frame_email_modo.pack(anchor="e", pady=(2, 0))
    Label(frame_email_modo, text="Modo email:", bg=APP_BG_COLOR, fg="white",
          font=("Segoe UI", 9)).pack(side=LEFT, padx=(0, 4))
    rb_por_pais = Radiobutton(frame_email_modo, text="1 por país", variable=email_modo_var,
                value="por_pais", bg=APP_BG_COLOR, fg="white",
                activebackground=APP_BG_COLOR, activeforeground="white",
                selectcolor=APP_BG_COLOR, font=("Segoe UI", 9))
    rb_por_pais.pack(side=LEFT)
    rb_consolidado = Radiobutton(frame_email_modo, text="Consolidado al final", variable=email_modo_var,
                value="consolidado", bg=APP_BG_COLOR, fg="white",
                activebackground=APP_BG_COLOR, activeforeground="white",
                selectcolor=APP_BG_COLOR, font=("Segoe UI", 9))
    rb_consolidado.pack(side=LEFT, padx=(6, 0))

    def _persistir_opciones_email(*_):
        cfg = cargar_config_global()
        cfg["enviar_mail"] = bool(enviar_mail_var.get())  # sincroniza archivo con checkbox
        cfg["adjuntar_resultados"] = bool(adjuntar_resultados_var.get())
        cfg["adjuntar_screenshots"] = bool(adjuntar_screenshots_var.get())

        # Guardar prefs de UI (browser, viewport, modo browser, modo email)
        navs = [n for n, v in [
            ("chrome", chrome_var), ("firefox", firefox_var), ("edge", edge_var),
            ("lambdatest_mac", lt_mac_var), ("lambdatest_android", lt_android_var),
        ] if v.get()]
        vps = [vp for vp, v in [
            ("fullscreen", desktop_var), ("600x738", mobile_var),
        ] if v.get()]
        cfg["ui_prefs"] = {
            "navegadores": navs if navs else ["chrome"],
            "viewports": vps if vps else ["fullscreen"],
            "visible_browser": bool(visible_browser_var.get()),
            "email_modo": email_modo_var.get(),
        }
        guardar_config_global(cfg)

        estado = "normal" if enviar_mail_var.get() else "disabled"
        entry_email.config(state=estado)
        chk_adj_resultados.config(state=estado)
        chk_adj_screens.config(state=estado)
        rb_por_pais.config(state=estado)
        rb_consolidado.config(state=estado)

    for _var in (enviar_mail_var, adjuntar_resultados_var, adjuntar_screenshots_var,
                 chrome_var, firefox_var, edge_var, lt_mac_var, lt_android_var,
                 desktop_var, mobile_var, visible_browser_var, email_modo_var):
        _var.trace_add("write", _persistir_opciones_email)
    _persistir_opciones_email()

    # === Estado de envío de email (muestra ⏳ / ✅ / ❌ en tiempo real) ===
    email_status_var = StringVar(value="")
    email_status_label = Label(
        frame_email_derecha,
        textvariable=email_status_var,
        font=("Segoe UI", 9, "italic"),
        bg=APP_BG_COLOR,
        fg="#7FFF7F",
        anchor="e",
    )
    email_status_label.pack(anchor="e", pady=(0, 2))

    def _set_email_status(success, error_msg=""):
        if success:
            email_status_var.set("✅ Email enviado correctamente")
            email_status_label.config(fg="#7FFF7F")
        else:
            email_status_var.set(f"❌ Error al enviar: {error_msg}")
            email_status_label.config(fg="#FF7F7F")

    from interface.helpers_interface import registrar_callback_ui_email

    def _global_email_ui_handler(estado, err_msg):
        def _update():
            if estado == "pending":
                email_status_var.set("⏳ Enviando email...")
                email_status_label.config(fg="#FFFF99")
            elif estado == "success":
                _set_email_status(True)
            else:
                _set_email_status(False, err_msg)
        root.after(0, _update)

    registrar_callback_ui_email(_global_email_ui_handler)

    #ttk.Separator(root, orient="horizontal", style="App.TSeparator").pack(fill="x", pady=10)

    # === PREVISUALIZACIÓN DE LEADS ===
    section_container = Frame(
        testing_tab,
        bg=SECTION_CONTAINER_BG_COLOR,
        highlightthickness=0,
        bd=0,
    )
    section_container.pack(fill="both", padx=20, pady=(10, 5))

    section_frame = Frame(
        section_container,
        bg=SECTION_BG_COLOR,
        highlightthickness=0,
        bd=0,
    )
    section_frame.pack(fill="both", padx=4, pady=4)

    Label(
        section_frame,
        text="Elija el país, tendrá una previsualización al Excel con la información del Lead.",
        font=("Segoe UI", 11, "bold"),
        bg=SECTION_BG_COLOR,
        fg=PRIMARY_TEXT_COLOR,
    ).pack(anchor="w", padx=12, pady=(12, 6))

    # === SELECTOR DE MODO DE EJECUCIÓN ===
    frame_selector = Frame(section_frame, bg=SECTION_BG_COLOR)
    frame_selector.pack(fill="x", padx=10, pady=(0, 4))

    Label(
        frame_selector,
        text="Modo de envío:",
        font=("Segoe UI", 9, "bold"),
        bg=SECTION_BG_COLOR,
        fg="white",
    ).pack(anchor="w", pady=(0, 4))

    _lt_modo_var = StringVar(value="")

    frame_modos = Frame(frame_selector, bg=SECTION_BG_COLOR)
    frame_modos.pack(anchor="w")
    for _modo_val, _modo_txt in [
        ("consecutive", "Múltiples países (consecutivo por sesión)"),
        ("parallel", "Múltiples países (en paralelo por sesión)"),
    ]:
        Radiobutton(
            frame_modos,
            text=_modo_txt,
            variable=_lt_modo_var,
            value=_modo_val,
            bg=SECTION_BG_COLOR,
            fg="white",
            selectcolor=SECTION_BG_COLOR,
            activebackground=SECTION_BG_COLOR,
            activeforeground="white",
            font=("Segoe UI", 9),
            cursor="hand2",
        ).pack(side=LEFT, padx=(0, 16))

    # Contenedor dinámico de selectores de país
    frame_seleccion_pais = Frame(frame_selector, bg=SECTION_BG_COLOR)
    frame_seleccion_pais.pack(anchor="w", pady=(6, 0))

    _paises_multi_vars = {}

    def _actualizar_estado_btn_enviar(*_):
        modo_ok = _lt_modo_var.get() != ""
        paises_ok = any(v.get() for v in _paises_multi_vars.values())
        btn_enviar_leads.config(state="normal" if (modo_ok and paises_ok) else "disabled")

    def _render_pais_selector(*_):
        for widget in frame_seleccion_pais.winfo_children():
            widget.destroy()
        _paises_multi_vars.clear()
        for _col_idx, _pais_nombre in enumerate(MAPEO_PAISES.keys()):
            _var = BooleanVar(value=False)
            _paises_multi_vars[_pais_nombre] = _var
            _var.trace_add("write", _actualizar_estado_btn_enviar)
            Checkbutton(
                frame_seleccion_pais,
                text=_pais_nombre,
                variable=_var,
                bg=SECTION_BG_COLOR,
                fg="white",
                activebackground=SECTION_BG_COLOR,
                activeforeground="white",
                selectcolor=SECTION_BG_COLOR,
                font=("Segoe UI", 8),
            ).grid(row=_col_idx // 5, column=_col_idx % 5, sticky="w", padx=6, pady=1)
        _actualizar_estado_btn_enviar()

    _lt_modo_var.trace_add("write", _render_pais_selector)
    _lt_modo_var.trace_add("write", _actualizar_estado_btn_enviar)

    def _ejecutar_envio():
        from concurrent.futures import ThreadPoolExecutor

        if _run_state["running"]:
            messagebox.showwarning("En ejecución", "Ya hay una ejecución en curso.")
            return

        modo = _lt_modo_var.get()
        navegadores = [b for b, v in [("chrome", chrome_var), ("firefox", firefox_var), ("edge", edge_var)] if v.get()]
        viewports = [vp for vp, v in [("fullscreen", desktop_var), ("600x738", mobile_var)] if v.get()]
        if not navegadores or not viewports:
            messagebox.showwarning("Configuración", "Seleccioná al menos un navegador y un viewport.")
            return

        paises_sel = [p for p, v in _paises_multi_vars.items() if v.get()]
        if not paises_sel:
            messagebox.showwarning("Selección vacía", "Marcá al menos un país para ejecutar.")
            return

        stop_ev = threading.Event()
        _run_state["stop_event"] = stop_ev
        _set_running(True)

        _background = not visible_browser_var.get()
        _consolidado = email_modo_var.get() == "consolidado"
        _resultados_consolidado = []
        _resultados_lock = threading.Lock()

        def _run_combo(pais, nav, vp):
            if stop_ev.is_set():
                return
            try:
                run_func = _get_run_func(pais)
                formulario = run_func(
                    browser=nav, viewport=vp, headless=False,
                    background=_background,
                    enviar_email=not _consolidado,
                )
                if _consolidado and formulario is not None:
                    excel = getattr(formulario, "RESULTADOS_PATH", None)
                    shots = getattr(formulario, "SCREENSHOT_DIR", None)
                    if excel and os.path.exists(excel):
                        with _resultados_lock:
                            _resultados_consolidado.append({
                                "pais": pais,
                                "navegador": nav,
                                "viewport": vp,
                                "estado": "completado",
                                "excel_path": excel,
                                "screenshots_dir": shots,
                            })
            except Exception as exc:
                print(f"Error ejecutando {pais} ({nav}/{vp}): {exc}")

        def _run_pais_con_browsers(pais):
            browser_combos = [(nav, vp) for nav in navegadores for vp in viewports]
            if len(browser_combos) <= 1:
                if browser_combos:
                    _run_combo(pais, *browser_combos[0])
                return
            with ThreadPoolExecutor(max_workers=len(browser_combos)) as ex:
                for f in [ex.submit(_run_combo, pais, nav, vp) for nav, vp in browser_combos]:
                    try:
                        f.result()
                    except Exception:
                        pass

        def _done():
            _set_running(False)
            if _consolidado and _resultados_consolidado:
                from interface.helpers_interface import enviar_email_resultados_consolidados
                threading.Thread(
                    target=lambda: enviar_email_resultados_consolidados(_resultados_consolidado),
                    daemon=True,
                ).start()

        if modo == "parallel":
            def _run_parallel():
                with ThreadPoolExecutor(max_workers=min(len(paises_sel), 9)) as ex:
                    for f in [ex.submit(_run_pais_con_browsers, p) for p in paises_sel]:
                        try:
                            f.result()
                        except Exception:
                            pass
                _done()
            threading.Thread(target=_run_parallel, daemon=True).start()
        else:
            def _run_sequential():
                for pais in paises_sel:
                    if stop_ev.is_set():
                        break
                    _run_pais_con_browsers(pais)
                _done()
            threading.Thread(target=_run_sequential, daemon=True).start()

    # Botón "Enviar Leads" unificado (blanco/negro, disabled hasta selección)
    btn_enviar_leads = ttk.Button(
        frame_selector,
        text="Enviar Leads",
        command=_ejecutar_envio,
        style="EnviarLeads.TButton",
        state="disabled",
    )
    btn_enviar_leads.pack(anchor="w", pady=(10, 4))
    _run_state["enviar_btn"] = btn_enviar_leads

    # Renderizar selector inicial
    _render_pais_selector()

    # === SECCIÓN DE BOTONES POR PAÍS ===
    tabs_container = Frame(section_frame, bg=SECTION_BG_COLOR)
    tabs_container.pack(fill="x", padx=10, pady=(0, 12))

    notebook = ttk.Notebook(tabs_container, style="Section.TNotebook")
    notebook.pack(fill="x", expand=True)

    btn_resultados = ttk.Button(
        section_container,
        text="Resultados",
        style="Results.TButton",
        command=abrir_carpeta_resultados,
    )
    btn_resultados.place(in_=tabs_container, relx=1.0, x=-4, y=0, anchor="ne")
    # Definición de países y sus scripts
    paises = [
        {
            "nombre": "Argentina",
            "scripts": [
                {"texto": "Completar Formularios", "script": "Formulario_Argentina_Main.py"},
            ]
        },
        {
            "nombre": "Bolivia",
            "scripts": [
                {"texto": "Completar Formularios", "script": "Formulario_Bolivia_Main.py"},
            ]
        },
        {
            "nombre": "Brasil",
            "scripts": [
                {"texto": "Completar Formularios", "script": "Formulario_Brasil_Main.py"},
            ]
        },
        {
            "nombre": "Chile",
            "scripts": [
                {"texto": "Completar Formularios", "script": "Formulario_Chile_Main.py"},
            ]
        },
        {
            "nombre": "Colombia",
            "scripts": [
                {"texto": "Completar Formularios", "script": "Formulario_Colombia_Main.py"},
            ]
        },
        {
            "nombre": "Ecuador",
            "scripts": [
                {"texto": "Completar Formularios", "script": "Formulario_Ecuador_Main.py"},
            ]
        },
        {
            "nombre": "Paraguay",
            "scripts": [
                {"texto": "Completar Formularios", "script": "Formulario_Paraguay_Main.py"},
            ]
        },
        {
            "nombre": "Peru",
            "scripts": [
                {"texto": "Completar Formularios", "script": "Formulario_Peru_Main.py"},
            ]
        },
        {
            "nombre": "Uruguay",
            "scripts": [
                {"texto": "Completar Formularios", "script": "Formulario_Uruguay_Main.py"},
            ]
        }
    ]

    # Crear pestañas para cada país
    for pais in paises:
        frame_pais = Frame(notebook, bg=SECTION_BG_COLOR)
        notebook.add(frame_pais, text=pais["nombre"])
        
        # Configurar grid para expansión
        frame_pais.grid_rowconfigure(2, weight=1)
        frame_pais.grid_columnconfigure(0, weight=1)
        frame_pais.grid_columnconfigure(1, weight=1)
                    
        # Área de tabla para mostrar/editar el Excel (una sola por pestaña)
        frame_tabla = Frame(
            frame_pais,
            bg=SECTION_BG_COLOR,
            highlightthickness=0,
            bd=0,
        )
        frame_tabla.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
        
        # Frame para controles de tabla - NUEVA ESTRUCTURA
        frame_controles = Frame(frame_tabla, bg=SECTION_BG_COLOR)
        frame_controles.pack(fill="x", pady=5)

        # Contenedor para botones de la izquierda (operaciones de tabla)
        frame_botones_izquierda = Frame(frame_controles, bg=SECTION_BG_COLOR)
        frame_botones_izquierda.pack(side=LEFT)

        # Contenedor para botones de la derecha (Excel y Formularios)
        frame_botones_derecha = Frame(frame_controles, bg=SECTION_BG_COLOR)
        frame_botones_derecha.pack(side=RIGHT)

        # Botones para controlar la tabla (IZQUIERDA)
        btn_agregar_fila = ttk.Button(
            frame_botones_izquierda,
            text="Agregar Fila",
            command=lambda tw=None: agregar_fila(tw) if tw else None,
            style="Section.TButton",
        )
        btn_agregar_fila.pack(side=LEFT, padx=5)

        btn_eliminar_fila = ttk.Button(
            frame_botones_izquierda,
            text="Eliminar Fila",
            command=lambda tw=None: eliminar_fila(tw) if tw else None,
            style="Section.TButton",
        )
        btn_eliminar_fila.pack(side=LEFT, padx=5)

        # Botón para clonar fila
        btn_clonar_fila = ttk.Button(
            frame_botones_izquierda,
            text="Clonar Fila",
            command=lambda tw=None: clonar_fila(tw) if tw else None,
            style="Section.TButton",
        )
        btn_clonar_fila.pack(side=LEFT, padx=5)

        btn_actualizar_tabla = ttk.Button(
            frame_botones_izquierda,
            text="Actualizar",
            command=lambda n=None, tw=None, ce=None: cargar_excel_a_tabla(n, tw, ce) if n and tw else None,
            style="Section.TButton",
        )
        btn_actualizar_tabla.pack(side=LEFT, padx=5)

        btn_guardar_tabla = ttk.Button(
            frame_botones_izquierda,
            text="Guardar Cambios",
            command=lambda n=None, tw=None: guardar_desde_tabla(n, tw) if n and tw else None,
            style="Section.TButton",
        )
        btn_guardar_tabla.pack(side=LEFT, padx=5)

        # Crear Treeview con scrollbars y altura limitada
        tree_frame = Frame(frame_tabla, bg=SECTION_BG_COLOR, bd=0, highlightthickness=0)
        tree_frame.pack(fill="both", expand=True)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # Scrollbars
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", style="Section.Vertical.TScrollbar")
        h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", style="Section.Horizontal.TScrollbar")

        # Treeview con altura limitada a 5 filas (25px de altura por fila = 125px)
        tree_excel = ttk.Treeview(
            tree_frame,
            yscrollcommand=lambda f, l, _sb=v_scroll: _autohide_yscroll(_sb, f, l),
            xscrollcommand=h_scroll.set,
            selectmode="browse",
            height=5,  # ← ESTA ES LA LÍNEA IMPORTANTE: limita a 5 filas visibles
            style="Section.Treeview",
        )
        v_scroll.config(command=tree_excel.yview)
        h_scroll.config(command=tree_excel.xview)

        # Posicionar con grid para mantener scrollbars visibles
        tree_excel.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        # Evitar que el espacio bajo el scroll horizontal colapse
        tree_frame.grid_rowconfigure(1, weight=0)
        
        # Crear editor de celdas
        cell_editor = CellEditor(tree_excel)
        
        # Botones de formulario y Excel en la barra superior derecha
        for i, script in enumerate(pais["scripts"], start=1):
            # Determinar nombre del archivo Excel
            excel_nombre = f"Lead_information_{script['script'].replace('.py', '')}"
            
            # === Botones en la barra superior derecha ===
            excel_ruta = os.path.join(DATA_DIR, f"{excel_nombre}.xlsx")

            if os.path.exists(excel_ruta):
                # Si existe, mostrar botón "Abrir Excel"
                btn_abrir_excel_superior = ttk.Button(
                    frame_botones_derecha,
                    text="Abrir Excel", 
                    command=lambda n=excel_nombre: abrir_excel(n),
                    style="Section.TButton",
                )
                btn_abrir_excel_superior.pack(side=RIGHT, padx=5)
            else:
                # Si no existe, mostrar botón "Crear Excel"
                btn_crear_excel_superior = ttk.Button(
                    frame_botones_derecha,
                    text="Crear Excel",
                    command=lambda n=excel_nombre, f=i, fp=frame_pais, ba=None, bg=None, tw=tree_excel, ce=cell_editor: crear_y_actualizar_excel(n, f, fp, ba, bg, tw, ce),
                    style="Section.TButton",
                )
                btn_crear_excel_superior.pack(side=RIGHT, padx=5)
            
            # Configurar los comandos de los botones de tabla con las referencias correctas
            btn_agregar_fila.config(command=lambda tw=tree_excel: agregar_fila(tw))
            btn_eliminar_fila.config(command=lambda tw=tree_excel: eliminar_fila(tw))
            btn_clonar_fila.config(command=lambda tw=tree_excel: clonar_fila(tw))
            btn_actualizar_tabla.config(command=lambda n=excel_nombre, tw=tree_excel, ce=cell_editor: cargar_excel_a_tabla(n, tw, ce))
            btn_guardar_tabla.config(command=lambda n=excel_nombre, tw=tree_excel: guardar_desde_tabla(n, tw))
            
            # Cargar el Excel si existe
            if os.path.exists(excel_ruta):
                cargar_excel_a_tabla(excel_nombre, tree_excel, cell_editor)
            else:
                # Mostrar mensaje indicando que no existe el Excel
                tree_excel["columns"] = ["info"]
                tree_excel["show"] = "headings"
                tree_excel.heading("info", text="Información")
                tree_excel.column("info", width=400)
                tree_excel.insert("", "end", values=["Cree el Excel primero usando el botón 'Crear Excel'"])

    #ttk.Separator(root, orient="horizontal", style="App.TSeparator").pack(fill="x", pady=10)

    # === CALENDARIO SEMANAL DE AUTOMATIZACIÓN ===
    programar_container = Frame(
        testing_tab,
        bg=SECTION_CONTAINER_BG_COLOR,
        highlightthickness=0,
        bd=0,
    )
    programar_container.pack(fill="x", padx=20, pady=15)

    from interface.weekly_scheduler import WeeklySchedulerPanel

    def _get_navegadores_seleccionados():
        navs = []
        if chrome_var.get():     navs.append("chrome")
        if firefox_var.get():    navs.append("firefox")
        if edge_var.get():       navs.append("edge")
        if lt_mac_var.get():     navs.append("lambdatest_mac")
        if lt_android_var.get(): navs.append("lambdatest_android")
        return navs

    def _get_viewports_seleccionados():
        vps = []
        if desktop_var.get(): vps.append("fullscreen")
        if mobile_var.get():  vps.append("600x738")
        return vps

    # Referencia forward para que actualizar_estado_botones funcione antes de crear el panel
    _scheduler_ref = {}

    def actualizar_estado_botones():
        panel = _scheduler_ref.get("panel")
        activo = panel.is_active() if panel else False
        estado = "disabled" if activo else "normal"
        for widget in root.winfo_children():
            if isinstance(widget, ttk.Notebook):
                for tab_id in widget.tabs():
                    tab = widget.nametowidget(tab_id)
                    for child in tab.winfo_children():
                        if isinstance(child, ttk.Button) and "Enviar Leads" in child.cget("text"):
                            child.config(state=estado)

    def _ejecutar_programacion(programacion, stop_event=None):
        """Ejecuta todos los tests del schedule semanal y retorna lista de resultados."""
        from glob import glob as glob_search

        def _stopped():
            return stop_event is not None and stop_event.is_set()

        resultados = []
        _LT_NAV = ("lambdatest_mac", "lambdatest_android")
        # LambdaTest no compatible con ejecución programada por ahora — se omite
        navegadores_prog = [n for n in programacion.get("navegadores", []) if n not in _LT_NAV]

        for pais_nombre in programacion.get("paises", []):
            if _stopped():
                print("⛔ Ejecución detenida por el usuario.")
                break
            for navegador in navegadores_prog:
                if _stopped():
                    break
                if navegador in _LT_NAV:
                    lt_type = "mac" if navegador == "lambdatest_mac" else "android"
                    try:
                        lt_dir = os.path.join(BASE_DIR,
                            "lambdatest_mac" if lt_type == "mac" else "lambdatest_android")
                        if lt_dir not in sys.path:
                            sys.path.insert(0, lt_dir)
                        lt_results_dir = os.path.join(BASE_DIR,
                            "resultados_lambdatestmac" if lt_type == "mac"
                            else "resultados_lambdatest_android")
                        antes = set(os.listdir(lt_results_dir)) if os.path.isdir(lt_results_dir) else set()
                        if lt_type == "mac":
                            import lt_controller  # type: ignore[import]
                            summary = lt_controller.run(pais=pais_nombre, build_name=f"Automatizado — {pais_nombre}", stop_event=stop_event)
                        else:
                            import lt_android_controller  # type: ignore[import]
                            summary = lt_android_controller.run(pais=pais_nombre, build_name=f"Automatizado — {pais_nombre}", stop_event=stop_event)
                        if summary.get("error"):
                            print(f"⚠️ LambdaTest {lt_type} — {pais_nombre}: {summary['error']}")
                        else:
                            excel_file = summary.get("results_excel")
                            if not excel_file or not os.path.exists(excel_file):
                                # Fallback: scan directory for new xlsx
                                if os.path.isdir(lt_results_dir):
                                    despues = set(os.listdir(lt_results_dir))
                                    nuevos = [f for f in (despues - antes) if f.endswith(".xlsx")]
                                    if nuevos:
                                        excel_file = os.path.join(lt_results_dir, sorted(nuevos)[-1])
                            if excel_file and os.path.exists(excel_file):
                                resultados.append({
                                    "pais": pais_nombre, "navegador": navegador,
                                    "viewport": lt_type, "estado": "completado",
                                    "excel_path": excel_file,
                                    "screenshots_dir": None,
                                })
                    except Exception as ex:
                        print(f"⚠️ Error LambdaTest {lt_type} — {pais_nombre}: {ex}")
                    time.sleep(2)
                else:
                    for viewport in programacion.get("viewports", []):
                        if _stopped():
                            break
                        env_param = f"{navegador}_{'desktop' if viewport == 'fullscreen' else 'mobile'}"
                        try:
                            pattern = os.path.join(RESULTS_DIR, f"resultados_{pais_nombre}*.xlsx")
                            antes_m = glob_search(pattern)
                            max_antes = max(
                                (int(os.path.basename(m).replace(f"resultados_{pais_nombre}", "").replace(".xlsx", ""))
                                 for m in antes_m
                                 if os.path.basename(m).replace(f"resultados_{pais_nombre}", "").replace(".xlsx", "").isdigit()),
                                default=0)
                            env_config = _get_environments().get(env_param)
                            if env_config is None:
                                print(f"⚠️ Entorno no reconocido: {env_param}")
                                continue
                            run_func = _get_run_func(pais_nombre)
                            run_func(browser=env_config["browser"], viewport=env_config["viewport"],
                                     headless=False, enviar_email=False, background=True)
                            despues_m = glob_search(pattern)
                            max_despues = max(
                                (int(os.path.basename(m).replace(f"resultados_{pais_nombre}", "").replace(".xlsx", ""))
                                 for m in despues_m
                                 if os.path.basename(m).replace(f"resultados_{pais_nombre}", "").replace(".xlsx", "").isdigit()),
                                default=0)
                            if max_despues > max_antes:
                                excel_file = os.path.join(RESULTS_DIR, f"resultados_{pais_nombre}{max_despues}.xlsx")
                                if os.path.exists(excel_file):
                                    resultados.append({
                                        "pais": pais_nombre, "navegador": navegador,
                                        "viewport": viewport, "estado": "completado",
                                        "excel_path": excel_file,
                                        "screenshots_dir": os.path.join(
                                            RESULTS_DIR, f"screenshots_{pais_nombre}{max_despues}"),
                                    })
                        except Exception as ex:
                            print(f"⚠️ Error ejecutando {pais_nombre} ({env_param}): {ex}")
                        time.sleep(2)
        return resultados

    from interface.helpers_interface import enviar_email_resultados_consolidados as _send_consolidated

    scheduler_panel = WeeklySchedulerPanel(
        programar_container,
        get_navegadores_cb=_get_navegadores_seleccionados,
        get_viewports_cb=_get_viewports_seleccionados,
        on_scheduling_change=lambda active: actualizar_estado_botones(),
        execute_cb=lambda prog, ev: _ejecutar_programacion(prog, ev),
        send_email_cb=lambda r: _send_consolidated(r) if enviar_mail_var.get() else None,
        root=root,
    )
    scheduler_panel.pack(fill="x", padx=4, pady=4)
    _scheduler_ref["panel"] = scheduler_panel
    actualizar_estado_botones()

    # === FOOTER STICKY (siempre visible abajo) ===
    frame_footer = Frame(outer_container, bg=APP_BG_COLOR)
    frame_footer.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 5))

    # Variable para rastrear la ventana de "Cómo se usa"
    ventana_indicaciones = None

    # Función para el popup de "Cómo se usa"
    def mostrar_indicaciones():
        nonlocal ventana_indicaciones
        
        # Si ya hay una ventana abierta, cerrarla
        if ventana_indicaciones is not None and ventana_indicaciones.winfo_exists():
            ventana_indicaciones.destroy()
            ventana_indicaciones = None
            return
        
        # Crear nueva ventana
        ventana_indicaciones = Toplevel(root)
        ventana_indicaciones.title("Cómo se usa")
        ventana_indicaciones.geometry("980x700")
        ventana_indicaciones.minsize(700, 500)
        ventana_indicaciones.configure(bg=APP_BG_COLOR)

        icon_path = os.path.join(ASSET_DIR, "icon.ico")
        if os.path.exists(icon_path):
            try:
                ventana_indicaciones.iconbitmap(icon_path)
            except Exception:
                pass

        container = Frame(ventana_indicaciones, bg=APP_BG_COLOR)
        container.pack(fill="both", expand=True)

        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        canvas = Canvas(container, bg=APP_BG_COLOR, highlightthickness=0)
        v_scroll = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        h_scroll = ttk.Scrollbar(container, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=lambda f, l, _sb=v_scroll: _autohide_yscroll(_sb, f, l),
                         xscrollcommand=h_scroll.set)

        canvas.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")

        content_frame = Frame(canvas, bg=APP_BG_COLOR)
        content_window = canvas.create_window((0, 0), window=content_frame, anchor="nw")
        text_labels = []
        image_blocks = []

        def _refresh_scrollregion(_event=None):
            try:
                canvas.configure(scrollregion=canvas.bbox("all"))
            except Exception:
                pass

        def _apply_responsive_layout():
            try:
                content_width = max(480, canvas.winfo_width() - 32)
                wrap = max(420, content_width - 40)

                for lbl in text_labels:
                    lbl.configure(wraplength=wrap)

                max_img_w = max(420, content_width - 44)
                new_refs = []
                for block in image_blocks:
                    original_img = block["image"]
                    img_label = block["label"]

                    if original_img.width > max_img_w:
                        ratio = max_img_w / float(original_img.width)
                        resized = original_img.resize(
                            (int(original_img.width * ratio), int(original_img.height * ratio)),
                            Image.LANCZOS,
                        )
                    else:
                        resized = original_img

                    photo = ImageTk.PhotoImage(resized)
                    img_label.configure(image=photo)
                    new_refs.append(photo)

                ventana_indicaciones._image_refs = new_refs
                _refresh_scrollregion()
            except Exception:
                pass

        def _on_canvas_resize(event):
            try:
                canvas.itemconfigure(content_window, width=max(480, event.width))
                _apply_responsive_layout()
            except Exception:
                pass

        def _on_mousewheel(event):
            try:
                if getattr(event, "state", 0) & 0x0001:
                    delta = int(-1 * (event.delta / 120))
                    canvas.xview_scroll(delta, "units")
                else:
                    delta = int(-1 * (event.delta / 120))
                    canvas.yview_scroll(delta, "units")
            except Exception:
                return "break"
            return "break"

        def _on_linux_wheel_up(_event):
            canvas.yview_scroll(-1, "units")
            return "break"

        def _on_linux_wheel_down(_event):
            canvas.yview_scroll(1, "units")
            return "break"

        content_frame.bind("<Configure>", _refresh_scrollregion)
        canvas.bind("<Configure>", _on_canvas_resize)
        ventana_indicaciones.bind("<MouseWheel>", _on_mousewheel)
        ventana_indicaciones.bind("<Shift-MouseWheel>", _on_mousewheel)
        ventana_indicaciones.bind("<Button-4>", _on_linux_wheel_up)
        ventana_indicaciones.bind("<Button-5>", _on_linux_wheel_down)

        Label(
            content_frame,
            text="Cómo se usa",
            font=("Segoe UI", 16, "bold"),
            bg=APP_BG_COLOR,
            fg="white",
            anchor="w",
            justify="left",
        ).pack(fill="x", padx=20, pady=(16, 8))

        url_sharepoint = "https://interpublic.sharepoint.com/sites/buemrmqa/_layouts/15/Doc.aspx?sourcedoc=%7B72455d97-49e3-4f3b-a71f-50b7d081e863%7D&action=edit&wd=target%28Automatizaci%C3%B3n%20de%20form.one%7C9fc98087-e699-46fd-92a4-d6d85119fc81%2FGu%C3%ADa%20de%20uso%20Osocio%20automatizaci%C3%B3n%20de%20form.%7C04d96e18-7215-48a3-af98-cd00777e8426%2F%29&wdorigin=703&wdpartid=%7Be9c543a1-820c-44a8-89d0-7b129cf12549%7D%7B16%7D&wdsectionfileid=%7Ba2134ac4-168c-4613-8694-cef6b74f3db7%7D"
        url_repositorio_git = "https://interpublic.sharepoint.com/sites/buemrmqa/_layouts/15/Doc.aspx?sourcedoc=%7B72455d97-49e3-4f3b-a71f-50b7d081e863%7D&action=edit&wd=target%28Automatizaci%C3%B3n%20de%20form.one%7C9fc98087-e699-46fd-92a4-d6d85119fc81%2FGu%C3%ADa%20de%20uso%20Osocio%20automatizaci%C3%B3n%20de%20form.%7C04d96e18-7215-48a3-af98-cd00777e8426%2F%29&wdorigin=703&wdpartid=%7Be9c543a1-820c-44a8-89d0-7b129cf12549%7D%7B16%7D&wdsectionfileid=%7Ba2134ac4-168c-4613-8694-cef6b74f3db7%7D"
        url_documentacion = "https://interpublic.sharepoint.com/sites/buemrmqa/_layouts/15/Doc.aspx?sourcedoc=%7B72455d97-49e3-4f3b-a71f-50b7d081e863%7D&action=edit&wd=target%28Automatizaci%C3%B3n%20de%20form.one%7C9fc98087-e699-46fd-92a4-d6d85119fc81%2FGu%C3%ADa%20de%20uso%20Osocio%20automatizaci%C3%B3n%20de%20form.%7C04d96e18-7215-48a3-af98-cd00777e8426%2F%29&wdorigin=703&wdpartid=%7Be9c543a1-820c-44a8-89d0-7b129cf12549%7D%7B16%7D&wdsectionfileid=%7Ba2134ac4-168c-4613-8694-cef6b74f3db7%7D"

        def _abrir_enlace(url):
            def _handler(_event=None):
                try:
                    webbrowser.open(url, new=2)
                except Exception as exc:
                    messagebox.showerror("Cómo se usa", f"No se pudo abrir el enlace:\n{exc}", parent=ventana_indicaciones)

            return _handler

        mensajes_ayuda = [
            "En el siguiente SharePoint podrás encontrar una guía completa, con explicaciones paso a paso, recomendaciones de uso y contexto funcional para aprovechar Osocio en profundidad.",
            "Acceder a la guía en SharePoint",
            "Si necesitás una referencia más técnica, orientada al código y al mantenimiento de la herramienta, también podés consultar el proyecto completo en el siguiente repositorio de Git.",
            "Abrir repositorio de Git",
            "Además, contás con una documentación adicional donde se describe la estructura general del proyecto, su lógica principal y varios detalles útiles para comprender mejor cómo está organizado.",
            "Ver documentación del proyecto",
        ]

        lbl_mensaje_sharepoint = Label(
            content_frame,
            text=mensajes_ayuda[0],
            font=("Segoe UI", 12),
            bg=APP_BG_COLOR,
            fg="white",
            anchor="w",
            justify="left",
            wraplength=920,
        )
        lbl_mensaje_sharepoint.pack(fill="x", padx=20, pady=(8, 6))
        text_labels.append(lbl_mensaje_sharepoint)

        lbl_link_sharepoint = Label(
            content_frame,
            text=mensajes_ayuda[1],
            font=("Segoe UI", 12, "underline"),
            bg=APP_BG_COLOR,
            fg="#9fd4ff",
            anchor="w",
            justify="left",
            wraplength=920,
            cursor="hand2",
        )
        lbl_link_sharepoint.pack(fill="x", padx=20, pady=(0, 12))
        lbl_link_sharepoint.bind("<Button-1>", _abrir_enlace(url_sharepoint))
        text_labels.append(lbl_link_sharepoint)

        lbl_mensaje_git = Label(
            content_frame,
            text=mensajes_ayuda[2],
            font=("Segoe UI", 12),
            bg=APP_BG_COLOR,
            fg="white",
            anchor="w",
            justify="left",
            wraplength=920,
        )
        lbl_mensaje_git.pack(fill="x", padx=20, pady=(0, 6))
        text_labels.append(lbl_mensaje_git)

        lbl_link_git = Label(
            content_frame,
            text=mensajes_ayuda[3],
            font=("Segoe UI", 12, "underline"),
            bg=APP_BG_COLOR,
            fg="#9fd4ff",
            anchor="w",
            justify="left",
            wraplength=920,
            cursor="hand2",
        )
        lbl_link_git.pack(fill="x", padx=20, pady=(0, 12))
        lbl_link_git.bind("<Button-1>", _abrir_enlace(url_repositorio_git))
        text_labels.append(lbl_link_git)

        lbl_mensaje_docs = Label(
            content_frame,
            text=mensajes_ayuda[4],
            font=("Segoe UI", 12),
            bg=APP_BG_COLOR,
            fg="white",
            anchor="w",
            justify="left",
            wraplength=920,
        )
        lbl_mensaje_docs.pack(fill="x", padx=20, pady=(0, 6))
        text_labels.append(lbl_mensaje_docs)

        lbl_link_docs = Label(
            content_frame,
            text=mensajes_ayuda[5],
            font=("Segoe UI", 12, "underline"),
            bg=APP_BG_COLOR,
            fg="#9fd4ff",
            anchor="w",
            justify="left",
            wraplength=920,
            cursor="hand2",
        )
        lbl_link_docs.pack(fill="x", padx=20, pady=(0, 12))
        lbl_link_docs.bind("<Button-1>", _abrir_enlace(url_documentacion))
        text_labels.append(lbl_link_docs)

        ventana_indicaciones.after(10, _apply_responsive_layout)

    # Variable para rastrear la ventana de "Más información"
    ventana_ejemplos = None

    # Función para el popup de "Más información"
    def mostrar_ejemplos_formularios():
        nonlocal ventana_ejemplos
        
        # Si ya hay una ventana abierta, cerrarla
        if ventana_ejemplos is not None and ventana_ejemplos.winfo_exists():
            ventana_ejemplos.destroy()
            ventana_ejemplos = None
            return
        
        # Crear nueva ventana
        ventana_ejemplos = Toplevel(root)
        ventana_ejemplos.title(" ")
        ventana_ejemplos.geometry("980x700")
        ventana_ejemplos.minsize(700, 500)
        ventana_ejemplos.configure(bg=APP_BG_COLOR)

        icon_path = os.path.join(ASSET_DIR, "icon.ico")
        if os.path.exists(icon_path):
            try:
                ventana_ejemplos.iconbitmap(icon_path)
            except Exception:
                pass

        container = Frame(ventana_ejemplos, bg=APP_BG_COLOR)
        container.pack(fill="both", expand=True)

        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        canvas = Canvas(container, bg=APP_BG_COLOR, highlightthickness=0)
        v_scroll = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        h_scroll = ttk.Scrollbar(container, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=lambda f, l, _sb=v_scroll: _autohide_yscroll(_sb, f, l),
                         xscrollcommand=h_scroll.set)

        canvas.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")

        content_frame = Frame(canvas, bg=APP_BG_COLOR)
        content_window = canvas.create_window((0, 0), window=content_frame, anchor="nw")
        text_labels = []

        def _refresh_scrollregion(_event=None):
            try:
                canvas.configure(scrollregion=canvas.bbox("all"))
            except Exception:
                pass

        def _apply_responsive_layout():
            try:
                content_width = max(480, canvas.winfo_width() - 32)
                wrap = max(420, content_width - 40)
                for lbl in text_labels:
                    lbl.configure(wraplength=wrap)
                _refresh_scrollregion()
            except Exception:
                pass

        def _on_canvas_resize(event):
            try:
                canvas.itemconfigure(content_window, width=max(480, event.width))
                _apply_responsive_layout()
            except Exception:
                pass

        def _on_mousewheel(event):
            try:
                if getattr(event, "state", 0) & 0x0001:
                    delta = int(-1 * (event.delta / 120))
                    canvas.xview_scroll(delta, "units")
                else:
                    delta = int(-1 * (event.delta / 120))
                    canvas.yview_scroll(delta, "units")
            except Exception:
                return "break"
            return "break"

        def _on_linux_wheel_up(_event):
            canvas.yview_scroll(-1, "units")
            return "break"

        def _on_linux_wheel_down(_event):
            canvas.yview_scroll(1, "units")
            return "break"

        content_frame.bind("<Configure>", _refresh_scrollregion)
        canvas.bind("<Configure>", _on_canvas_resize)
        ventana_ejemplos.bind("<MouseWheel>", _on_mousewheel)
        ventana_ejemplos.bind("<Shift-MouseWheel>", _on_mousewheel)
        ventana_ejemplos.bind("<Button-4>", _on_linux_wheel_up)
        ventana_ejemplos.bind("<Button-5>", _on_linux_wheel_down)

        titulo = Label(
            content_frame,
            text=" ",
            font=("Segoe UI", 16, "bold"),
            bg=APP_BG_COLOR,
            fg="white",
            anchor="w",
            justify="left",
            wraplength=920,
        )
        titulo.pack(fill="x", padx=20, pady=(16, 8))
        text_labels.append(titulo)

        # Contenido intencionalmente vacío por ahora.
        spacer = Frame(content_frame, bg=APP_BG_COLOR, height=640)
        spacer.pack(fill="both", expand=True)

        ventana_ejemplos.after(10, _apply_responsive_layout)

    # 5a. Izquierda: "Cómo se usa" (clickeable)
    label_como_se_usa = Label(
        frame_footer,
        text="Cómo se usa",
        font=("Segoe UI", 11, "underline bold"),
        bg=APP_BG_COLOR,
        fg="white",
        cursor="hand2",
    )
    label_como_se_usa.pack(side=LEFT, padx=20)
    label_como_se_usa.bind("<Button-1>", lambda e: mostrar_indicaciones())

    # 5b. Centro: "Hecho por Ariel Melgratti"
    label_autor = Label(
        frame_footer,
        text="Made by Ariel Melgratti",
        font=("Segoe UI", 11, "italic"),
        bg=APP_BG_COLOR,
        fg="white",
    )
    label_autor.pack(side=LEFT, expand=True)

    # 5c. Derecha: "Más información" (clickeable)
    label_ejemplos = Label(
        frame_footer,
        text="              ",
        font=("Segoe UI", 11),
        bg=APP_BG_COLOR,
        fg="white",
        #cursor="hand2",
    )
    label_ejemplos.pack(side=RIGHT, padx=20)
    #label_ejemplos.bind("<Button-1>", lambda e: mostrar_ejemplos_formularios())

    def _on_close():
        sys.stdout = _original_stdout
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", _on_close)
    root.mainloop()

if __name__ == "__main__":
    iniciar_interfaz()