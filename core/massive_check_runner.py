import os
import time
import requests
import datetime
import openpyxl
from openpyxl.styles import PatternFill, Font
from selenium.webdriver.common.by import By

# Import base classes dynamically to ensure project compatibility
import sys
_base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
for p in (_base, os.path.join(_base, "core")):
    if p not in sys.path:
        sys.path.insert(0, p)

from core.generic_country_base import GenericCountryBase
from utils.data_generator import generar_fila_datos

class MassiveFormFiller(GenericCountryBase):
    """
    Subclase especializada de GenericCountryBase para el chequeo masivo sin envíos.
    Configura de forma dura el modo rellenado parcial / no enviar lead,
    y escribe las capturas en la ruta específica solicitada.
    """
    def __init__(self, country_name, browser="chrome", headless=True, background=True):
        super().__init__(country_name, browser=browser, headless=headless, background=background)
        self.config["solo_verificar_visual"] = True
        self.config["no_enviar_lead"] = True

    def setup_directories_and_files(self, custom_dir=None):
        """Sobrescribir para guardar las capturas en la subcarpeta resultados/resultado_urlsinsertas/Capturas"""
        if custom_dir:
            self.SCREENSHOT_DIR = os.path.normpath(custom_dir)
        else:
            self.SCREENSHOT_DIR = os.path.normpath(os.path.join(self.BASE_DIR, "resultados", "resultado_urlsinsertas", "Capturas"))
        os.makedirs(self.SCREENSHOT_DIR, exist_ok=True)
        self.RUN_NUMBER = 1
        return 1

    def fill_minimum_fields_massive(self):
        """
        Completa al menos un campo de texto (Nombre/Apellido) con caracteres aleatorios (mínimo 4)
        y un dropdown (Select o personalizado) seleccionando una opción aleatoria no placeholder.
        """
        import random
        nombres_db = ["Carlos", "Daniel", "Mariana", "Alejandro", "Sofia", "Mateo", "Camila", "Nicolas", "Valentina", "Sebastian"]
        apellidos_db = ["Gomez", "Rodriguez", "Lopez", "Fernandez", "Martinez", "Perez", "Gonzalez", "Sanchez", "Alvarez", "Romero"]
        
        # Generar nombre y apellido aleatorios de mínimo 4 caracteres
        num_names = random.randint(1, 2)
        first_val = " ".join(random.choices(nombres_db, k=num_names))
        while len(first_val) < 4:
            first_val = " ".join(random.choices(nombres_db, k=num_names))
            
        last_val = " ".join(random.choices(apellidos_db, k=random.randint(1, 2)))
        while len(last_val) < 4:
            last_val = " ".join(random.choices(apellidos_db, k=random.randint(1, 2)))

        # 1. Rellenar Campo de Texto (Nombre o Apellido)
        inputs = self.driver.find_elements(By.TAG_NAME, "input")
        name_inputs = []
        lastname_inputs = []
        generic_text_inputs = []
        
        for inp in inputs:
            try:
                if not inp.is_displayed() or not inp.is_enabled():
                    continue
                itype = str(inp.get_attribute("type") or "").lower()
                if itype not in ["text", "", "email", "tel"]:
                    continue
                    
                name_attr = str(inp.get_attribute("name") or "").lower()
                id_attr = str(inp.get_attribute("id") or "").lower()
                placeholder = str(inp.get_attribute("placeholder") or "").lower()
                class_attr = str(inp.get_attribute("class") or "").lower()
                
                if any(x in name_attr or x in id_attr or x in placeholder or x in class_attr for x in ["firstname", "first_name", "nombre", "name"]):
                    if "last" not in name_attr and "last" not in id_attr and "apellido" not in name_attr and "apellido" not in id_attr:
                        name_inputs.append(inp)
                        continue
                
                if any(x in name_attr or x in id_attr or x in placeholder or x in class_attr for x in ["lastname", "last_name", "apellido"]):
                    lastname_inputs.append(inp)
                    continue
                    
                generic_text_inputs.append(inp)
            except Exception:
                pass

        filled_text = False
        target_inp = None
        target_val = first_val
        
        if name_inputs:
            target_inp = name_inputs[0]
            target_val = first_val
        elif lastname_inputs:
            target_inp = lastname_inputs[0]
            target_val = last_val
        elif generic_text_inputs:
            target_inp = generic_text_inputs[0]
            target_val = first_val

        if target_inp:
            try:
                target_inp.clear()
                target_inp.send_keys(target_val)
                filled_text = True
            except Exception as e:
                print(f"Error rellenando campo de texto masivo: {e}")

        # Rellenar también el segundo campo de texto si existe
        if filled_text:
            try:
                if target_inp in name_inputs and lastname_inputs:
                    lastname_inputs[0].clear()
                    lastname_inputs[0].send_keys(last_val)
                elif target_inp in lastname_inputs and name_inputs:
                    name_inputs[0].clear()
                    name_inputs[0].send_keys(first_val)
            except Exception:
                pass

        # 2. Seleccionar Dropdown (Select o custom)
        selects = self.driver.find_elements(By.TAG_NAME, "select")
        filled_dropdown = False
        
        for sel in selects:
            try:
                if not sel.is_displayed() or not sel.is_enabled():
                    continue
                from selenium.webdriver.support.ui import Select
                select_obj = Select(sel)
                valid_options = []
                for opt in select_obj.options:
                    otxt = (opt.text or "").strip()
                    oval = (opt.get_attribute("value") or "").strip()
                    otxt_lower = otxt.lower()
                    
                    if any(p in otxt_lower for p in ["selecciona", "seleccione", "selecione", "select", "elige", "elegir", "choose", "--", "placeholder"]):
                        continue
                    if not otxt or not oval:
                        continue
                    valid_options.append(opt)
                    
                if valid_options:
                    chosen_opt = random.choice(valid_options)
                    select_obj.select_by_visible_text(chosen_opt.text)
                    filled_dropdown = True
                    break
            except Exception as e:
                print(f"Error dropdown nativo: {e}")

        # Fallback para custom dropdowns (AEM / custom divs)
        if not filled_dropdown:
            try:
                dropdown_triggers = self.driver.find_elements(By.XPATH, "//*[contains(@class, 'select') or contains(@class, 'dropdown') or @role='combobox' or @role='listbox']")
                for trigger in dropdown_triggers:
                    if not trigger.is_displayed() or not trigger.is_enabled():
                        continue
                    try:
                        trigger.click()
                        time.sleep(0.5)
                        options = self.driver.find_elements(By.XPATH, "//li[not(contains(translate(text(), 'SELECI ONAREG', 'selecionareg'), 'selecciona')) and not(contains(translate(text(), 'SELECI ONAREG', 'selecionareg'), 'seleccione'))] | //*[ @role='option' and not(contains(translate(text(), 'SELECI ONAREG', 'selecionareg'), 'selecciona'))]")
                        valid_custom_opts = [o for o in options if o.is_displayed() and (o.text or "").strip()]
                        if valid_custom_opts:
                            chosen_opt = random.choice(valid_custom_opts)
                            chosen_opt.click()
                            filled_dropdown = True
                            break
                    except Exception:
                        pass
            except Exception as e:
                print(f"Error custom dropdown fallback: {e}")
        return (filled_text and filled_dropdown)

    def check_single_url(self, landing_url, expected_form_url, row_num, tomar_capturas=True):
        """
        Navega a la URL y ejecuta la lógica de verificación (HTTP Status, Iframe Match, Llenado parcial)
        """
        if getattr(self, "stop_event", None) and self.stop_event.is_set():
            raise Exception("Cancelado por el usuario")
        landing_url = self._sanitize_url(landing_url)
        expected_form_url = self._sanitize_url(expected_form_url)

        self.expected_form_url = expected_form_url
        self.ss_counter = row_num
        self._url_form_encontrado = ""
        self._landing_issue = ""
        self._errores_ss_taken = False
        
        if self.screenshot_manager:
            self.screenshot_manager.url_form_esperado = expected_form_url
            self.screenshot_manager.url_form_encontrado = ""
            self.screenshot_manager.form_slug = self._slug_for(landing_url, expected_form_url)
            self.screenshot_manager.current_frame = None

        # 1. Obtener HTTP Status Code y Redirecciones del landing page
        landing_status_str = ""
        final_url_found = landing_url
        try:
            import urllib3
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            }
            r = requests.get(landing_url, headers=headers, timeout=8, allow_redirects=True, verify=False)
            final_url_found = r.url
            if r.history:
                # Ocurrió un redirect
                code = r.history[0].status_code
                landing_status_str = f"{code} - 🔀 FAIL Link redirects. Code: {code}, URL Original: {landing_url} redirects to: {r.url}"
            else:
                if r.status_code == 200:
                    landing_status_str = f"200 - ✅ PASS Valid link. HTTP Code: 200"
                else:
                    landing_status_str = f"{r.status_code} - ❌ FAIL HTTP Code: {r.status_code}"
        except Exception as e:
            landing_status_str = f"Error - ❌ FAIL HTTP Code: {type(e).__name__}"

        # 1b. Obtener HTTP Status del expected secure URL (Secure_Excel)
        secure_status_str = ""
        if expected_form_url:
            try:
                import urllib3
                urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
                headers = {
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
                }
                r_sec = requests.get(expected_form_url, headers=headers, timeout=8, allow_redirects=True, verify=False)
                if r_sec.status_code == 200:
                    secure_status_str = "200 - ✅ PASS Valid link. HTTP Code: 200"
                else:
                    secure_status_str = f"{r_sec.status_code} - ❌ FAIL HTTP Code: {r_sec.status_code}"
            except Exception as e:
                secure_status_str = f"Error - ❌ FAIL HTTP Code: {type(e).__name__}"
        else:
            secure_status_str = "N/A"

        if getattr(self, "stop_event", None) and self.stop_event.is_set():
            raise Exception("Cancelado por el usuario")
        # 2. Navegar con Selenium
        ss_landing_path = ""
        try:
            ss_landing_path = self.process_landing_page(landing_url, row_num, take_screenshot=tomar_capturas)
        except Exception as e:
            return {
                "ok": False,
                "error": f"Error navegación: {str(e)}",
                "current_url": final_url_found,
                "landing_status": landing_status_str,
                "iframe_correct": "NO",
                "form_url_found": "",
                "form_coincide": "❌ FAIL",
                "displayed": "❌ FAIL",
                "secure_status": secure_status_str,
                "forms_count": 0,
                "ss_landing": ss_landing_path or "",
                "ss_completado": ""
            }

        # Validar si base_form_filler detectó 404 en el título/encabezados
        if self._landing_issue:
            return {
                "ok": False,
                "error": self._landing_issue,
                "current_url": self.driver.current_url,
                "landing_status": f"404 - ❌ FAIL {self._landing_issue}",
                "iframe_correct": "NO",
                "form_url_found": "",
                "form_coincide": "❌ FAIL",
                "displayed": "❌ FAIL",
                "secure_status": secure_status_str,
                "forms_count": 0,
                "ss_landing": ss_landing_path,
                "ss_completado": ""
            }

        # 3. Detectar iframe e inyectar URL de form
        use_iframe = bool(expected_form_url)
        form_url_mismatch = False
        target_iframe = None
        iframe_src = ""
        
        # Contar iframes disponibles (solo los que correspondan a formularios GM)
        forms_count = 0
        try:
            iframes_found = self.driver.find_elements(By.TAG_NAME, "iframe")
            for iframe in iframes_found:
                try:
                    src = str(iframe.get_attribute("src") or "").lower()
                    if any(x in src for x in ["gm_forms", "gm_formns", "gm_front", "gm_admin"]):
                        forms_count += 1
                except Exception:
                    pass
        except Exception:
            pass

        if use_iframe:
            try:
                target_iframe = self.find_and_position_to_form(expected_form_url)
                if not target_iframe:
                    # Intento de fallback
                    try:
                        _cand, _es_gm = self._pick_gm_iframe(iframes_found, expected_url=expected_form_url)
                    except Exception:
                        _cand, _es_gm = None, False
                        
                    if _cand is not None:
                        target_iframe = _cand
                        if not _es_gm:
                            form_url_mismatch = True
                    else:
                        form_url_mismatch = True
                
                if target_iframe:
                    iframe_src = target_iframe.get_attribute("src") or ""
                    self.driver.switch_to.frame(target_iframe)
                    if self.screenshot_manager:
                        self.screenshot_manager.current_frame = target_iframe
            except Exception as e:
                return {
                    "ok": False,
                    "error": f"Error con iframe: {str(e)}",
                    "current_url": self.driver.current_url,
                    "landing_status": landing_status_str,
                    "iframe_correct": "NO",
                    "form_url_found": "",
                    "form_coincide": "❌ FAIL",
                    "displayed": "❌ FAIL",
                    "secure_status": secure_status_str,
                    "forms_count": forms_count,
                    "ss_landing": ss_landing_path,
                    "ss_completado": ""
                }
        else:
            iframe_src = self.driver.current_url

        iframe_correct_str = "SÍ" if (target_iframe and not form_url_mismatch) else ("NO" if use_iframe else "N/A (Standalone)")
        
        # Validar coincidencia estricta de URL (Excel == Inserted)
        form_coincide_str = "✅ PASS"
        if use_iframe:
            if not target_iframe or form_url_mismatch:
                form_coincide_str = f"❌ FAIL Mismatch: expected {expected_form_url} but found {iframe_src}"
        
        if use_iframe and not target_iframe:
            return {
                "ok": False,
                "error": "Iframe del formulario no encontrado en la landing.",
                "current_url": self.driver.current_url,
                "landing_status": landing_status_str,
                "iframe_correct": "NO",
                "form_url_found": "",
                "form_coincide": form_coincide_str,
                "displayed": "❌ FAIL",
                "secure_status": secure_status_str,
                "forms_count": forms_count,
                "ss_landing": ss_landing_path,
                "ss_completado": ""
            }

        if getattr(self, "stop_event", None) and self.stop_event.is_set():
            raise Exception("Cancelado por el usuario")
        # 4. Rellenado parcial (Nombre/Apellido y Dropdown)
        ss_completado_path = ""
        displayed_str = "❌ FAIL"
        fill_success = False
        try:
            self.wait_for_form_ready_in_iframe()
            if getattr(self, "stop_event", None) and self.stop_event.is_set():
                raise Exception("Cancelado por el usuario")
            # Ejecutar rellenado mínimo personalizado (dos campos interactivos)
            fill_success = self.fill_minimum_fields_massive()
            if fill_success:
                displayed_str = "✅ PASS"
            else:
                displayed_str = "❌ FAIL"
            time.sleep(0.5)
            
            # Tomar captura después del rellenado si tomar_capturas está activo
            if tomar_capturas and self.screenshot_manager:
                ss_completado_path = self.screenshot_manager.take_form_screenshot(row_num, "completado", full_page=True)
        except Exception as e:
            displayed_str = "❌ FAIL"
            return {
                "ok": False,
                "error": f"Error completando campos: {str(e)}",
                "current_url": self.driver.current_url,
                "landing_status": landing_status_str,
                "iframe_correct": iframe_correct_str,
                "form_url_found": iframe_src,
                "form_coincide": form_coincide_str,
                "displayed": displayed_str,
                "secure_status": secure_status_str,
                "forms_count": forms_count,
                "ss_landing": ss_landing_path,
                "ss_completado": ""
            }

        return {
            "ok": True if fill_success else False,
            "error": "" if fill_success else "No se pudo rellenar los dos campos interactivos (nombre y dropdown)",
            "current_url": self.driver.current_url,
            "landing_status": landing_status_str,
            "iframe_correct": iframe_correct_str,
            "form_url_found": iframe_src,
            "form_coincide": form_coincide_str,
            "displayed": displayed_str,
            "secure_status": secure_status_str,
            "forms_count": forms_count,
            "ss_landing": ss_landing_path,
            "ss_completado": ss_completado_path
        }


def load_passed_results_from_wb(wb, custom_cols):
    passed_results = {}
    for s_name in wb.sheetnames:
        sheet = wb[s_name]
        header_row_idx = None
        for r_idx, row in enumerate(sheet.iter_rows(values_only=True), start=1):
            row_vals = [str(x).strip().upper() for x in row if x is not None]
            if (custom_cols["segmento"].upper() in row_vals and
                custom_cols["estado"].upper() in row_vals and
                custom_cols["url_live"].upper() in row_vals and
                custom_cols["url_secure"].upper() in row_vals):
                header_row_idx = r_idx
                break
        if not header_row_idx:
            continue
            
        header_row = [str(cell.value).strip().upper() if cell.value is not None else "" for cell in sheet[header_row_idx]]
        try:
            col_seg = header_row.index(custom_cols["segmento"].upper()) + 1
            col_est = header_row.index(custom_cols["estado"].upper()) + 1
            col_url_live = header_row.index(custom_cols["url_live"].upper()) + 1
            col_url_secure = header_row.index(custom_cols["url_secure"].upper()) + 1
            
            col_current = header_row.index("CURRENTURL") + 1
            col_landing = header_row.index("LANDINGSTATUS") + 1
            col_secure_ins = header_row.index("SECURE_INSERTED") + 1
            col_excel_ins = header_row.index("EXCEL==INSERTED?") + 1
            col_disp = header_row.index("DISPLAYED?") + 1
            col_status_ins = header_row.index("STATUSINSERTED") + 1
            col_forms = header_row.index("FORMSCOUNT") + 1
            
            col_comments = None
            for col_idx in range(1, len(header_row) + 1):
                val = header_row[col_idx - 1]
                if val == "COMENTARIOS" or val == "COMENTARIO":
                    col_comments = col_idx
                    break
        except ValueError:
            continue
            
        for r_idx in range(header_row_idx + 1, sheet.max_row + 1):
            url_live = str(sheet.cell(row=r_idx, column=col_url_live).value or "").strip()
            url_secure = str(sheet.cell(row=r_idx, column=col_url_secure).value or "").strip()
            if not url_live and not url_secure:
                continue
                
            res_val = str(sheet.cell(row=r_idx, column=col_excel_ins).value or "").strip()
            disp_val = str(sheet.cell(row=r_idx, column=col_disp).value or "").strip()
            land_val = str(sheet.cell(row=r_idx, column=col_landing).value or "").strip()
            
            is_ok = ("✅ PASS" in res_val) and ("✅ PASS" in disp_val) and ("✅ PASS" in land_val)
            if is_ok:
                key = (url_live.lower(), url_secure.lower())
                passed_results[key] = {
                    "currentUrl": sheet.cell(row=r_idx, column=col_current).value,
                    "LandingStatus": sheet.cell(row=r_idx, column=col_landing).value,
                    "Secure_Inserted": sheet.cell(row=r_idx, column=col_secure_ins).value,
                    "Excel==Inserted?": sheet.cell(row=r_idx, column=col_excel_ins).value,
                    "Displayed?": sheet.cell(row=r_idx, column=col_disp).value,
                    "statusInserted": sheet.cell(row=r_idx, column=col_status_ins).value,
                    "formsCount": sheet.cell(row=r_idx, column=col_forms).value,
                    "Comentarios": sheet.cell(row=r_idx, column=col_comments).value if col_comments else ""
                }
    return passed_results


def run_massive_check(excel_path, custom_cols, selected_markets, borrar_comentarios=False, tomar_capturas=True, solo_fails=False, browser="chrome", headless=True, progress_callback=None, stop_event=None):
    """
    Función controladora del chequeo masivo.
    Crea un nuevo archivo Excel consolidado y carpetas separadas por cada mercado.
    """
    import time
    start_time = time.time()
    # 1. Crear carpeta destino dentro de resultados/resultado_urlsinsertas
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    dest_dir = os.path.join(base_dir, "resultados", "resultado_urlsinsertas")
    os.makedirs(dest_dir, exist_ok=True)
    
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_excel_name = f".temp_consolidado_{timestamp}.xlsx"
    out_excel_path = os.path.join(dest_dir, out_excel_name)

    if progress_callback:
        progress_callback("GENERAL", "Cargando archivo Excel...", 0)

    # Cargar Excel con openpyxl
    try:
        wb = openpyxl.load_workbook(excel_path)
    except Exception as e:
        return False, f"Error cargando Excel: {e}"

    # Crear libro consolidado de salida limpio
    wb_out = openpyxl.Workbook()
    if "Sheet" in wb_out.sheetnames:
        del wb_out["Sheet"]

    # Cargar nombres de hojas
    sheet_names = wb.sheetnames
    
    # Mapeo de hojas a países
    sheet_to_country = {
        "COLOMBIA": "Colombia",
        "ECUADOR": "Ecuador",
        "ARGENTINA": "Argentina",
        "CHILE": "Chile",
        "PERÚ": "Peru",
        "PERU": "Peru",
        "PARAGUAY": "Paraguay",
        "URUGUAY": "Uruguay",
        "BOLIVIA": "Bolivia",
        "BRASIL": "Brasil",
        "BRAZIL": "Brasil"
    }

    # Definir estilos premium con colores verde/rojo
    fill_pass = PatternFill(fill_type="solid", fgColor="C6EFCE") # verde claro
    font_pass = Font(color="006100", bold=True)
    
    fill_fail = PatternFill(fill_type="solid", fgColor="FFC7CE") # rojo claro
    font_fail = Font(color="9C0006", bold=True)
    
    fill_skip = PatternFill(fill_type="solid", fgColor="EAECEE") # gris claro
    font_bold = Font(bold=True)

    # Helper para colorear celdas basado en contenido
    def color_cell_by_content(cell):
        val_str = str(cell.value or "")
        if "✅ PASS" in val_str or "PASS" in val_str or "SÍ" == val_str:
            cell.fill = fill_pass
            cell.font = font_pass
        elif "❌ FAIL" in val_str or "FAIL" in val_str or "NO" == val_str or "Mismatch" in val_str:
            cell.fill = fill_fail
            cell.font = font_fail

    # Contadores generales
    total_processed = 0
    total_skipped = 0
    total_failed = 0
    total_passed = 0
    wb_country = None   # se crea por hoja procesada; queda None si no matcheó ningún mercado

    # Normalizar mercados seleccionados a mayúsculas
    selected_markets_upper = [m.upper() for m in selected_markets]

    # Iterar por cada hoja
    for s_idx, s_name in enumerate(sheet_names):
        if stop_event and stop_event.is_set():
            break

        # El nombre de la hoja no siempre coincide con el del mercado en la UI: la hoja de
        # Perú se llama "PERÚ" (con tilde) y la UI manda "Peru". Se compara contra el país
        # normalizado por sheet_to_country además del nombre crudo, si no Perú (y Brazil)
        # quedaban afuera de la revisión sin ningún aviso.
        country_name = sheet_to_country.get(s_name.upper())
        if (s_name.upper() not in selected_markets_upper
                and (country_name or "").upper() not in selected_markets_upper):
            # Omitir mercados no seleccionados
            continue

        if not country_name:
            continue

        sheet = wb[s_name]

        # Buscar la fila cabecera
        header_row_idx = None
        for r_idx, row in enumerate(sheet.iter_rows(values_only=True), start=1):
            row_vals = [str(x).strip().upper() for x in row if x is not None]
            if (custom_cols["segmento"].upper() in row_vals and
                custom_cols["estado"].upper() in row_vals and
                custom_cols["url_live"].upper() in row_vals and
                custom_cols["url_secure"].upper() in row_vals):
                header_row_idx = r_idx
                break

        if not header_row_idx:
            print(f"No se encontró la fila cabecera en hoja {s_name}")
            continue

        # Obtener los índices de columna correspondientes (1-indexed para openpyxl)
        header_row = [str(cell.value).strip().upper() if cell.value is not None else "" for cell in sheet[header_row_idx]]
        
        col_indices = {}
        try:
            col_indices["segmento"] = header_row.index(custom_cols["segmento"].upper()) + 1
            col_indices["estado"] = header_row.index(custom_cols["estado"].upper()) + 1
            col_indices["url_live"] = header_row.index(custom_cols["url_live"].upper()) + 1
            col_indices["url_secure"] = header_row.index(custom_cols["url_secure"].upper()) + 1
        except ValueError as e:
            print(f"Error indexando columnas en la hoja {s_name}: {e}")
            continue

        # Comprobar si ya existe la columna de Comentarios en el original
        comentarios_col_idx = None
        for col_idx in range(1, len(header_row) + 1):
            val = header_row[col_idx - 1]
            if val == "COMENTARIOS" or val == "COMENTARIO":
                comentarios_col_idx = col_idx
                break

        # Cargar resultados anteriores de forma semántica si solo_fails está activo
        passed_results = {}
        if solo_fails:
            try:
                passed_results.update(load_passed_results_from_wb(wb, custom_cols))
            except Exception:
                pass
            try:
                files = [os.path.join(dest_dir, f) for f in os.listdir(dest_dir) if f.startswith("Resultados_Revision_Masiva_") and f.endswith(".xlsx")]
                if not files:
                    # check temp or subdirectories too
                    files = []
                    for root, dirs, filenames in os.walk(dest_dir):
                        for f in filenames:
                            if f.startswith("Resultados_Revision_Masiva_") and f.endswith(".xlsx"):
                                files.append(os.path.join(root, f))
                if files:
                    files.sort(key=os.path.getmtime, reverse=True)
                    wb_latest = openpyxl.load_workbook(files[0], data_only=True)
                    passed_results.update(load_passed_results_from_wb(wb_latest, custom_cols))
            except Exception:
                pass

        # 1a. Crear carpeta de salida específica del mercado
        country_folder_name = f"resultadoMasivo_{country_name}"
        country_dir = os.path.join(dest_dir, country_folder_name)
        os.makedirs(country_dir, exist_ok=True)
        
        # Carpeta de capturas del mercado
        country_capturas_dir = os.path.join(country_dir, "Capturas") if tomar_capturas else None
        if country_capturas_dir:
            os.makedirs(country_capturas_dir, exist_ok=True)

        # 1b. Crear hojas en consolidado y libro individual para el mercado
        out_sheet = wb_out.create_sheet(title=s_name)
        
        wb_country = openpyxl.Workbook()
        if "Sheet" in wb_country.sheetnames:
            del wb_country["Sheet"]
        country_sheet = wb_country.create_sheet(title=s_name)

        new_headers = [
            custom_cols["segmento"],
            custom_cols["estado"],
            custom_cols["url_live"],
            custom_cols["url_secure"],
            "currentUrl",
            "LandingStatus",
            "Secure_Inserted",
            "Excel==Inserted?",
            "Displayed?",
            "statusInserted",
            "formsCount",
            "Comentarios"
        ]
        
        for c_idx, title in enumerate(new_headers, start=1):
            # Consolidado
            cell_cons = out_sheet.cell(row=1, column=c_idx)
            cell_cons.value = title
            cell_cons.font = font_bold
            # Individual
            cell_ind = country_sheet.cell(row=1, column=c_idx)
            cell_ind.value = title
            cell_ind.font = font_bold

        new_col_indices = {
            "currentUrl": 5,
            "LandingStatus": 6,
            "Secure_Inserted": 7,
            "Excel==Inserted?": 8,
            "Displayed?": 9,
            "statusInserted": 10,
            "formsCount": 11,
            "Comentarios": 12
        }

        # Inicializar el runner para el país actual
        filler = None
        try:
            # Iterar por las filas de datos
            max_row = sheet.max_row
            active_rows = []
            
            # Pre-escanear para saber cuántas filas ON hay que procesar para el progreso local
            for r_idx in range(header_row_idx + 1, max_row + 1):
                url_live = str(sheet.cell(row=r_idx, column=col_indices["url_live"]).value or "").strip()
                url_secure = str(sheet.cell(row=r_idx, column=col_indices["url_secure"]).value or "").strip()
                if not url_live and not url_secure:
                    continue
                active_rows.append(r_idx)

            total_active_sheet = len(active_rows)
            processed_sheet = 0
            out_row_idx = 2

            for i_idx, r_idx in enumerate(active_rows):
                if stop_event and stop_event.is_set():
                    break

                segmento = str(sheet.cell(row=r_idx, column=col_indices["segmento"]).value or "").strip()
                estado = str(sheet.cell(row=r_idx, column=col_indices["estado"]).value or "").strip()
                url_live = str(sheet.cell(row=r_idx, column=col_indices["url_live"]).value or "").strip()
                url_secure = str(sheet.cell(row=r_idx, column=col_indices["url_secure"]).value or "").strip()

                # Conservar comentarios si corresponde
                comentario_anterior = ""
                if comentarios_col_idx:
                    comentario_anterior = sheet.cell(row=r_idx, column=comentarios_col_idx).value or ""

                # Determinar si la fila está OFF
                is_off = "OFF" in estado.upper() or "❌" in estado

                # Si está activado solo_fails y esta URL ya pasó previamente, copiar resultado
                key = (url_live.lower(), url_secure.lower())
                if solo_fails and key in passed_results and not is_off:
                    total_processed += 1
                    total_passed += 1
                    processed_sheet += 1
                    
                    prev_res = passed_results[key]
                    for col_key, col_pos in new_col_indices.items():
                        if col_key == "Comentarios":
                            val = "" if borrar_comentarios else (comentario_anterior or prev_res.get("Comentarios") or "")
                        else:
                            val = prev_res.get(col_key)
                            
                        for s in (out_sheet, country_sheet):
                            s.cell(row=out_row_idx, column=col_pos).value = val
                            
                    for s in (out_sheet, country_sheet):
                        for col_key, col_pos in new_col_indices.items():
                            if col_key != "Comentarios":
                                color_cell_by_content(s.cell(row=out_row_idx, column=col_pos))

                    if progress_callback:
                        pct = int((processed_sheet / total_active_sheet) * 100)
                        progress_callback(country_name, f"Revisando row {r_idx}: {segmento} (Prev PASS)", pct)

                    out_row_idx += 1
                    continue

                # Escribir las columnas de entrada en las hojas de salida
                for s in (out_sheet, country_sheet):
                    s.cell(row=out_row_idx, column=1).value = segmento
                    s.cell(row=out_row_idx, column=2).value = estado
                    s.cell(row=out_row_idx, column=3).value = url_live
                    s.cell(row=out_row_idx, column=4).value = url_secure

                if is_off:
                    total_skipped += 1
                    processed_sheet += 1
                    
                    # Registrar como omitido
                    for s in (out_sheet, country_sheet):
                        s.cell(row=out_row_idx, column=5).value = "N/A"
                        s.cell(row=out_row_idx, column=6).value = "OFF - Omitido"
                        s.cell(row=out_row_idx, column=7).value = "N/A"
                        s.cell(row=out_row_idx, column=8).value = "SKIPPED"
                        s.cell(row=out_row_idx, column=9).value = "SKIPPED"
                        s.cell(row=out_row_idx, column=10).value = "N/A"
                        s.cell(row=out_row_idx, column=11).value = 0
                        
                        if borrar_comentarios:
                            s.cell(row=out_row_idx, column=12).value = ""
                        else:
                            s.cell(row=out_row_idx, column=12).value = comentario_anterior

                        # Colorear en gris
                        for col_key, col_pos in new_col_indices.items():
                            if col_key != "Comentarios" or borrar_comentarios:
                                s.cell(row=out_row_idx, column=col_pos).fill = fill_skip
                    
                    # Reportar progreso intermedio
                    if progress_callback:
                        pct = int((processed_sheet / total_active_sheet) * 100)
                        progress_callback(country_name, f"{processed_sheet}/{total_active_sheet} (OFF)", pct)
                    
                    out_row_idx += 1
                    continue

                # Si está ON o prioridad, realizar verificación
                if not filler:
                    if progress_callback:
                        progress_callback(country_name, "Abriendo navegador...", 0)
                    filler = MassiveFormFiller(country_name, browser=browser, headless=headless, background=True)
                    filler.stop_event = stop_event
                    filler.setup_directories_and_files(custom_dir=country_capturas_dir)
                    filler.initialize_browser()
                    if not tomar_capturas:
                        filler.screenshot_manager = None

                processed_sheet += 1
                if progress_callback:
                    pct = int((processed_sheet / total_active_sheet) * 100)
                    progress_callback(country_name, f"Revisando row {r_idx}: {segmento}", pct)

                total_processed += 1
                
                # Ejecutar verificación de URL
                res = filler.check_single_url(url_live, url_secure, r_idx, tomar_capturas=tomar_capturas)

                # Guardar valores
                for s in (out_sheet, country_sheet):
                    s.cell(row=out_row_idx, column=5).value = res["current_url"]
                    s.cell(row=out_row_idx, column=6).value = res["landing_status"]
                    s.cell(row=out_row_idx, column=7).value = res["form_url_found"]
                    s.cell(row=out_row_idx, column=8).value = res["form_coincide"]
                    s.cell(row=out_row_idx, column=9).value = res["displayed"]
                    s.cell(row=out_row_idx, column=10).value = res["secure_status"]
                    s.cell(row=out_row_idx, column=11).value = res["forms_count"]
                    
                    if borrar_comentarios:
                        s.cell(row=out_row_idx, column=12).value = ""
                    else:
                        s.cell(row=out_row_idx, column=12).value = comentario_anterior

                    # Colorear cada celda según el contenido
                    for col_key, col_pos in new_col_indices.items():
                        if col_key != "Comentarios":
                            color_cell_by_content(s.cell(row=out_row_idx, column=col_pos))

                if res["ok"]:
                    total_passed += 1
                else:
                    total_failed += 1

                out_row_idx += 1

        except Exception as e:
            print(f"Error procesando la hoja {s_name}: {e}")
        finally:
            if filler:
                try:
                    filler.driver.quit()
                except Exception:
                    pass
                filler = None

    # Calcular tiempo transcurrido antes de guardar para incluirlo en los Excels
    elapsed = time.time() - start_time
    hours = int(elapsed // 3600)
    minutes = int((elapsed % 3600) // 60)
    seconds = int(elapsed % 60)
    if hours > 0:
        time_str = f"{hours}h {minutes}m {seconds}s"
    elif minutes > 0:
        time_str = f"{minutes}m {seconds}s"
    else:
        time_str = f"{seconds}s"

    # Escribir fila de resumen con el tiempo en cada hoja de cada libro
    summary_font = Font(bold=True, color="7D4E9F")
    summary_fill = PatternFill(start_color="F2E6FF", end_color="F2E6FF", fill_type="solid")
    # wb_country sólo existe si se procesó al menos una hoja. Si ningún mercado
    # seleccionado matcheó una hoja del Excel, antes reventaba acá con UnboundLocalError
    # en vez de avisar que no había nada para revisar.
    if wb_country is None:
        return False, ("No se encontró ninguna hoja para los mercados seleccionados "
                       f"({', '.join(selected_markets)}). Hojas del Excel: {', '.join(sheet_names)}.")

    for wb in [wb_out, wb_country]:
        for ws_name in wb.sheetnames:
            ws = wb[ws_name]
            summary_row = ws.max_row + 2
            cell_label = ws.cell(row=summary_row, column=1)
            cell_label.value = "Duración del test:"
            cell_label.font = summary_font
            cell_label.fill = summary_fill
            cell_value = ws.cell(row=summary_row, column=2)
            cell_value.value = time_str
            cell_value.font = summary_font
            cell_value.fill = summary_fill
            # Stats
            stats_row = summary_row + 1
            ws.cell(row=stats_row, column=1).value = "Resumen:"
            ws.cell(row=stats_row, column=1).font = summary_font
            ws.cell(row=stats_row, column=1).fill = summary_fill
            stats_text = f"Procesados: {total_processed} | PASS: {total_passed} | FAIL: {total_failed} | Omitidos: {total_skipped}"
            ws.cell(row=stats_row, column=2).value = stats_text
            ws.cell(row=stats_row, column=2).font = summary_font
            ws.cell(row=stats_row, column=2).fill = summary_fill

    # Guardar Excel individual del mercado en su carpeta (re-save con tiempo incluido)
    try:
        country_excel_name = f"Resultados_Revision_Masiva_{country_name}_{timestamp}.xlsx"
        country_excel_path = os.path.join(country_dir, country_excel_name)
        wb_country.save(country_excel_path)
    except Exception as e:
        print(f"Error guardando libro de resultados individual ({country_name}): {e}")

    # Guardar Excel consolidado de salida
    try:
        wb_out.save(out_excel_path)
    except Exception as e:
        print(f"Error guardando libro de resultados consolidado: {e}")
        return False, f"Error al guardar reporte: {e}"

    msg = f"Revisión finalizada en {time_str}. Procesados: {total_processed} (PASS: {total_passed}, FAIL: {total_failed}), Omitidos: {total_skipped}"
    return True, {
        "msg": msg,
        "excel_path": out_excel_path,
        "total_processed": total_processed,
        "total_passed": total_passed,
        "total_failed": total_failed,
        "total_skipped": total_skipped,
        "elapsed_time": time_str
    }



def sincronizar_reporte_con_matriz(master_path, results_path, custom_cols):
    """
    Sincroniza y compara el reporte de resultados (results_path) con la matriz origen (master_path).
    Soporta la adición de nuevas filas, actualización de estados (ON/OFF), actualización de URLs (live/secure),
    y detección de eliminados. Utiliza matching exacto de URLs y fallback por Segmento.
    """
    try:
        wb_master = openpyxl.load_workbook(master_path, data_only=True)
        wb_results = openpyxl.load_workbook(results_path)
    except Exception as e:
        return False, f"Error cargando los archivos Excel: {e}"

    stats = {
        "added": 0,
        "updated_state": 0,
        "deleted": 0
    }
    
    font_bold = Font(bold=True)
    timestamp = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    changes_log = []

    # Mapear nombres de columnas de salida y sus posiciones en el reporte
    new_headers = [
        custom_cols["segmento"],
        custom_cols["estado"],
        custom_cols["url_live"],
        custom_cols["url_secure"],
        "currentUrl",
        "LandingStatus",
        "Secure_Inserted",
        "Excel==Inserted?",
        "Displayed?",
        "statusInserted",
        "formsCount",
        "Comentarios"
    ]

    # Iterar por cada hoja en la matriz maestra
    for s_name in wb_master.sheetnames:
        sheet_master = wb_master[s_name]
        
        # 1. Buscar cabecera en maestro
        header_row_idx_master = None
        for r_idx, row in enumerate(sheet_master.iter_rows(values_only=True), start=1):
            row_vals = [str(x).strip().upper() for x in row if x is not None]
            if (custom_cols["segmento"].upper() in row_vals and
                custom_cols["estado"].upper() in row_vals and
                custom_cols["url_live"].upper() in row_vals and
                custom_cols["url_secure"].upper() in row_vals):
                header_row_idx_master = r_idx
                break
        
        if not header_row_idx_master:
            continue

        # Obtener índices de columna de maestro
        header_row_master = [str(cell.value).strip().upper() if cell.value is not None else "" for cell in sheet_master[header_row_idx_master]]
        
        try:
            col_seg_m = header_row_master.index(custom_cols["segmento"].upper()) + 1
            col_est_m = header_row_master.index(custom_cols["estado"].upper()) + 1
            col_url_live_m = header_row_master.index(custom_cols["url_live"].upper()) + 1
            col_url_secure_m = header_row_master.index(custom_cols["url_secure"].upper()) + 1
        except ValueError:
            continue

        comentarios_col_idx_m = None
        for col_idx in range(1, len(header_row_master) + 1):
            val = header_row_master[col_idx - 1]
            if val == "COMENTARIOS" or val == "COMENTARIO":
                comentarios_col_idx_m = col_idx
                break

        # 2. Obtener o crear la hoja correspondiente en resultados
        if s_name not in wb_results.sheetnames:
            sheet_results = wb_results.create_sheet(title=s_name)
            for c_idx, title in enumerate(new_headers, start=1):
                cell = sheet_results.cell(row=1, column=c_idx)
                cell.value = title
                cell.font = font_bold
            header_row_idx_res = 1
        else:
            sheet_results = wb_results[s_name]
            header_row_idx_res = None
            for r_idx, row in enumerate(sheet_results.iter_rows(values_only=True), start=1):
                row_vals = [str(x).strip().upper() for x in row if x is not None]
                if (custom_cols["segmento"].upper() in row_vals and
                    custom_cols["estado"].upper() in row_vals and
                    custom_cols["url_live"].upper() in row_vals and
                    custom_cols["url_secure"].upper() in row_vals):
                    header_row_idx_res = r_idx
                    break
            
            if not header_row_idx_res:
                # Si la hoja existe pero está rota o vacía, la recreamos
                sheet_results.delete_rows(1, sheet_results.max_row + 1)
                for c_idx, title in enumerate(new_headers, start=1):
                    cell = sheet_results.cell(row=1, column=c_idx)
                    cell.value = title
                    cell.font = font_bold
                header_row_idx_res = 1

        # 3. Leer registros de maestro
        master_rows = []
        for r_idx in range(header_row_idx_master + 1, sheet_master.max_row + 1):
            url_live = str(sheet_master.cell(row=r_idx, column=col_url_live_m).value or "").strip()
            url_secure = str(sheet_master.cell(row=r_idx, column=col_url_secure_m).value or "").strip()
            if not url_live and not url_secure:
                continue
                
            segmento = str(sheet_master.cell(row=r_idx, column=col_seg_m).value or "").strip()
            estado = str(sheet_master.cell(row=r_idx, column=col_est_m).value or "").strip()
            comentario = ""
            if comentarios_col_idx_m:
                comentario = str(sheet_master.cell(row=r_idx, column=comentarios_col_idx_m).value or "").strip()

            master_rows.append({
                "segmento": segmento,
                "estado": estado,
                "url_live": url_live,
                "url_secure": url_secure,
                "comentario": comentario,
                "matched": False
            })

        # 4. Leer registros de resultados
        results_rows = []
        header_row_res = [str(cell.value).strip().upper() if cell.value is not None else "" for cell in sheet_results[header_row_idx_res]]
        
        try:
            col_seg_r = header_row_res.index(custom_cols["segmento"].upper()) + 1
            col_est_r = header_row_res.index(custom_cols["estado"].upper()) + 1
            col_url_live_r = header_row_res.index(custom_cols["url_live"].upper()) + 1
            col_url_secure_r = header_row_res.index(custom_cols["url_secure"].upper()) + 1
        except ValueError:
            # Si no coincide exactamente con el formato esperado, no podemos mapear
            continue

        col_comentarios_r = 12 # Columna Comentarios estándar
        
        for r_idx in range(header_row_idx_res + 1, sheet_results.max_row + 1):
            url_live = str(sheet_results.cell(row=r_idx, column=col_url_live_r).value or "").strip()
            url_secure = str(sheet_results.cell(row=r_idx, column=col_url_secure_r).value or "").strip()
            if not url_live and not url_secure:
                continue
                
            segmento = str(sheet_results.cell(row=r_idx, column=col_seg_r).value or "").strip()
            estado = str(sheet_results.cell(row=r_idx, column=col_est_r).value or "").strip()
            comentario = str(sheet_results.cell(row=r_idx, column=col_comentarios_r).value or "").strip()

            results_rows.append({
                "row_idx": r_idx,
                "segmento": segmento,
                "estado": estado,
                "url_live": url_live,
                "url_secure": url_secure,
                "comentario": comentario,
                "matched": False
            })

        # 5. Sincronizar - Paso A: Matching exacto por (url_live, url_secure)
        for m_row in master_rows:
            for r_row in results_rows:
                if r_row["matched"]:
                    continue
                if (m_row["url_live"].lower() == r_row["url_live"].lower() and 
                    m_row["url_secure"].lower() == r_row["url_secure"].lower()):
                    
                    # Match exacto encontrado
                    m_row["matched"] = True
                    r_row["matched"] = True
                    r_idx = r_row["row_idx"]
                    
                    # Sincronizar Segmento si cambió
                    if r_row["segmento"] != m_row["segmento"]:
                        sheet_results.cell(row=r_idx, column=col_seg_r).value = m_row["segmento"]

                    # Sincronizar Estado si cambió
                    if r_row["estado"].upper() != m_row["estado"].upper():
                        sheet_results.cell(row=r_idx, column=col_est_r).value = m_row["estado"]
                        
                        # Si cambió a ON, limpiar celdas de prueba para obligar al rerun
                        is_new_on = "OFF" not in m_row["estado"].upper() and "❌" not in m_row["estado"]
                        if is_new_on:
                            for col_pos in range(5, 12):
                                cell = sheet_results.cell(row=r_idx, column=col_pos)
                                cell.value = None
                                cell.fill = PatternFill(fill_type=None)
                                cell.font = Font(color="000000", bold=False)

                        # Registrar comentario
                        prev_comm = str(sheet_results.cell(row=r_idx, column=col_comentarios_r).value or "").strip()
                        new_comm = f"[Actualización {timestamp}]: Estado cambiado de {r_row['estado']} a {m_row['estado']}."
                        if prev_comm:
                            sheet_results.cell(row=r_idx, column=col_comentarios_r).value = f"{prev_comm} | {new_comm}"
                        else:
                            sheet_results.cell(row=r_idx, column=col_comentarios_r).value = new_comm
                        
                        stats["updated_state"] += 1
                        changes_log.append(f"[ESTADO] {r_row['segmento']}: {r_row['estado']} → {m_row['estado']}")
                    break

        # 6. Sincronizar - Paso B: Matching por Segmento (para detectar cambios de URL en la misma fila/segmento)
        for m_row in master_rows:
            if m_row["matched"]:
                continue
            for r_row in results_rows:
                if r_row["matched"]:
                    continue
                if m_row["segmento"].lower() == r_row["segmento"].lower():
                    # Match por segmento encontrado (las URLs cambiaron)
                    m_row["matched"] = True
                    r_row["matched"] = True
                    r_idx = r_row["row_idx"]
                    
                    # Actualizar URLs
                    sheet_results.cell(row=r_idx, column=col_url_live_r).value = m_row["url_live"]
                    sheet_results.cell(row=r_idx, column=col_url_secure_r).value = m_row["url_secure"]
                    
                    # Sincronizar Estado si cambió
                    sheet_results.cell(row=r_idx, column=col_est_r).value = m_row["estado"]
                    
                    # Como cambiaron las URLs, limpiar SIEMPRE las columnas de resultados anteriores
                    for col_pos in range(5, 12):
                        cell = sheet_results.cell(row=r_idx, column=col_pos)
                        cell.value = None
                        cell.fill = PatternFill(fill_type=None)
                        cell.font = Font(color="000000", bold=False)

                    # Registrar comentario de actualización
                    prev_comm = str(sheet_results.cell(row=r_idx, column=col_comentarios_r).value or "").strip()
                    new_comm = f"[Actualización {timestamp}]: URLs y datos actualizados (Live/Secure modificadas)."
                    if prev_comm:
                        sheet_results.cell(row=r_idx, column=col_comentarios_r).value = f"{prev_comm} | {new_comm}"
                    else:
                        sheet_results.cell(row=r_idx, column=col_comentarios_r).value = new_comm
                    
                    stats["updated_state"] += 1
                    changes_log.append(f"[URLS] {m_row['segmento']}: URLs Live/Secure actualizadas")
                    break

        # 7. Procesar eliminados (registros en resultados que no se emparejaron)
        for r_row in results_rows:
            if not r_row["matched"]:
                r_idx = r_row["row_idx"]
                old_est = r_row["estado"]
                if "ELIMINADO" not in old_est.upper() and "OFF" not in old_est.upper():
                    sheet_results.cell(row=r_idx, column=col_est_r).value = "❌ OFF (Eliminado de origen)"
                    
                    # Registrar comentario
                    prev_comm = str(sheet_results.cell(row=r_idx, column=col_comentarios_r).value or "").strip()
                    new_comm = f"[Actualización {timestamp}]: Eliminado de la matriz origen."
                    if prev_comm:
                        sheet_results.cell(row=r_idx, column=col_comentarios_r).value = f"{prev_comm} | {new_comm}"
                    else:
                        sheet_results.cell(row=r_idx, column=col_comentarios_r).value = new_comm
                    
                    stats["deleted"] += 1
                    changes_log.append(f"[ELIMINADO] {r_row['segmento']}: Eliminado de la matriz origen")

        # 8. Agregar nuevas filas (registros en maestro que no se emparejaron)
        for m_row in master_rows:
            if not m_row["matched"]:
                next_row = sheet_results.max_row + 1
                
                # Escribir columnas de entrada
                sheet_results.cell(row=next_row, column=1).value = m_row["segmento"]
                sheet_results.cell(row=next_row, column=2).value = m_row["estado"]
                sheet_results.cell(row=next_row, column=3).value = m_row["url_live"]
                sheet_results.cell(row=next_row, column=4).value = m_row["url_secure"]
                
                # Inicializar columnas de test
                for col_pos in range(5, 12):
                    sheet_results.cell(row=next_row, column=col_pos).value = ""
                    
                # Comentario
                new_comm = f"[Actualización {timestamp}]: Agregado por sincronización."
                if m_row["comentario"]:
                    sheet_results.cell(row=next_row, column=12).value = f"{m_row['comentario']} | {new_comm}"
                else:
                    sheet_results.cell(row=next_row, column=12).value = new_comm
                
                stats["added"] += 1
                changes_log.append(f"[NUEVO] {m_row['segmento']}: {m_row['url_live']}")

    # Guardar el reporte actualizado
    try:
        wb_results.save(results_path)
    except Exception as e:
        return False, f"Error guardando los resultados actualizados: {e}"

    # Generar Excel de changelog en la misma carpeta del reporte
    changelog_path = ""
    if changes_log:
        try:
            ts_file = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            results_dir = os.path.dirname(results_path)
            changelog_path = os.path.join(results_dir, f"Changelog_Sincronizacion_{ts_file}.xlsx")
            
            wb_log = openpyxl.Workbook()
            ws_log = wb_log.active
            ws_log.title = "Changelog"
            
            # Headers
            log_headers = ["Tipo", "Hoja", "Segmento", "Detalle", "Fecha/Hora"]
            header_font = Font(bold=True, color="FFFFFF")
            header_fill = PatternFill(start_color="7D4E9F", end_color="7D4E9F", fill_type="solid")
            for c_idx, h in enumerate(log_headers, 1):
                cell = ws_log.cell(row=1, column=c_idx)
                cell.value = h
                cell.font = header_font
                cell.fill = header_fill
            
            # Filas de cambios
            add_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            upd_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
            del_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            
            for log_idx, entry in enumerate(changes_log, 2):
                if entry.startswith("[NUEVO]"):
                    tipo = "Agregado"
                    fill = add_fill
                elif entry.startswith("[ESTADO]"):
                    tipo = "Estado Actualizado"
                    fill = upd_fill
                elif entry.startswith("[URLS]"):
                    tipo = "URLs Actualizadas"
                    fill = upd_fill
                elif entry.startswith("[ELIMINADO]"):
                    tipo = "Eliminado de Origen"
                    fill = del_fill
                else:
                    tipo = "Cambio"
                    fill = upd_fill
                
                # Parse: [TYPE] segmento: detalle
                parts = entry.split("] ", 1)
                detail = parts[1] if len(parts) > 1 else entry
                seg_parts = detail.split(": ", 1)
                segmento = seg_parts[0] if seg_parts else ""
                detalle = seg_parts[1] if len(seg_parts) > 1 else detail
                
                ws_log.cell(row=log_idx, column=1).value = tipo
                ws_log.cell(row=log_idx, column=1).fill = fill
                ws_log.cell(row=log_idx, column=2).value = ""  # Hoja (se puede expandir)
                ws_log.cell(row=log_idx, column=3).value = segmento
                ws_log.cell(row=log_idx, column=4).value = detalle
                ws_log.cell(row=log_idx, column=5).value = timestamp
            
            # Ajustar anchos
            ws_log.column_dimensions["A"].width = 20
            ws_log.column_dimensions["B"].width = 15
            ws_log.column_dimensions["C"].width = 30
            ws_log.column_dimensions["D"].width = 60
            ws_log.column_dimensions["E"].width = 22
            
            wb_log.save(changelog_path)
        except Exception as e:
            print(f"Error generando changelog: {e}")
            changelog_path = ""

    stats["changelog_path"] = changelog_path
    stats["changes_log"] = changes_log
    return True, stats
