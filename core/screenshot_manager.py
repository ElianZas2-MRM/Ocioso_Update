"""
screenshot_manager.py — Captura y combina screenshots del formulario.
Chrome y Edge hacen scroll + captura + merge para página completa.
Firefox usa su propia función nativa de screenshot de página entera.
"""
import os
import time
from PIL import Image
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait

class ScreenshotManager:
    """Gestiona la toma de screenshots completos de páginas"""
    
    def __init__(self, driver, screenshot_dir):
        self.driver = driver
        self.screenshot_dir = screenshot_dir
        self.browser_name = driver.name.lower()
        self.url_landing        = ""
        self.url_form_esperado  = ""
        self.url_form_encontrado = ""
        # Firefox no puede tomar screenshots en contexto iframe; guardamos el iframe activo
        # para salir a default_content antes del screenshot y volver después.
        self.current_frame = None

    def _add_url_banner(self, path: str):
        """Agrega un banner oscuro en la parte superior con las URLs de la sesión."""
        try:
            from PIL import ImageDraw, ImageFont
            img = Image.open(path).convert("RGB")
            w, h = img.size
            AMARILLO = (220, 220, 100)
            VERDE    = (80, 200, 80)
            ROJO     = (220, 80, 80)
            lineas = []  # (texto, color)
            if self.url_landing:
                lineas.append((f"Landing:         {self.url_landing}", AMARILLO))
            if self.url_form_esperado:
                lineas.append((f"Form esperado:   {self.url_form_esperado}", AMARILLO))
            if self.url_form_encontrado:
                lineas.append((f"Form encontrado: {self.url_form_encontrado}", AMARILLO))
                if self.url_form_esperado:
                    if self.url_form_encontrado == self.url_form_esperado:
                        lineas.append(("  ✓  Coincide con el esperado", VERDE))
                    else:
                        lineas.append(("  ✗  NO coincide con el esperado", ROJO))
            if not lineas:
                return
            font_size = max(14, w // 80)
            line_h    = font_size + 6
            banner_h  = line_h * len(lineas) + 10
            banner    = Image.new("RGB", (w, banner_h), (30, 30, 30))
            draw      = ImageDraw.Draw(banner)
            try:
                font = ImageFont.truetype("arial.ttf", font_size)
            except Exception:
                font = ImageFont.load_default()
            for i, (line, color) in enumerate(lineas):
                draw.text((8, 5 + i * line_h), line, fill=color, font=font)
            combined = Image.new("RGB", (w, banner_h + h))
            combined.paste(banner, (0, 0))
            combined.paste(img, (0, banner_h))
            self._save_compressed(combined, path)
        except Exception as e:
            print(f"Error banner URLs: {e}")

    def _save_compressed(self, img, path):
        """Guarda el PNG pesando lo menos posible sin perder legibilidad: las capturas
        largas (landings kilométricas) se cuantizan a 256 colores."""
        try:
            if img.height > 2500:
                img = img.quantize(colors=256, method=Image.MEDIANCUT).convert("P")
            img.save(path, "PNG", optimize=True)
        except Exception:
            img.save(path)
        
    def take_full_page_screenshot(self, filename):
        """Toma screenshot completa de toda la página uniendo múltiples capturas"""
        screenshot_path = os.path.join(self.screenshot_dir, filename)
        in_iframe = (self.current_frame is not None)
        
        # Firefox: no puede capturar estando dentro de un iframe; salir a default_content primero
        # y volver al iframe después (si no, el llenado posterior no encuentra los campos).
        if self.driver.name.lower() == 'firefox':
            if in_iframe:
                try:
                    self.driver.switch_to.default_content()
                except Exception:
                    pass
            try:
                try:
                    self.driver.get_full_page_screenshot_as_file(screenshot_path)
                    print(f"Captura completa Firefox (nativa) guardada: {filename}")
                    return True
                except Exception as e:
                    print(f"Error en captura nativa Firefox: {e}")
                    try:
                        self.driver.save_screenshot(screenshot_path)
                        print(f"Captura de respaldo Firefox guardada: {filename}")
                        return True
                    except Exception as e2:
                        print(f"Error crítico captura Firefox: {e2}")
                        return False
            finally:
                if in_iframe:
                    try:
                        self.driver.switch_to.frame(self.current_frame)
                    except Exception:
                        pass
        
        # Chrome, Edge y otros navegadores: método original con scroll y merge
        scroll_y = 0
        try:
            if in_iframe:
                try:
                    self.driver.switch_to.default_content()
                except Exception:
                    pass
            # Guardar el scroll del documento principal para devolverlo al terminar
            # (si no, la página queda arriba de todo tras cada captura full-page).
            try:
                scroll_y = self.driver.execute_script("return window.pageYOffset;") or 0
            except Exception:
                scroll_y = 0

            total_width = self.driver.execute_script("return document.body.scrollWidth")
            total_height = self.driver.execute_script("return document.body.parentNode.scrollHeight")
            viewport_width = self.driver.execute_script("return window.innerWidth")
            viewport_height = self.driver.execute_script("return window.innerHeight")
            
            print(f"Dimensiones página: {total_width}x{total_height}px")
            
            self.driver.execute_script("window.scrollTo(0, 0);")
            time.sleep(0.5)
            
            if total_height <= viewport_height:
                self.driver.save_screenshot(screenshot_path)
                print(f"Captura simple guardada: {filename}")
                return True
            
            screenshots = []
            current_position = 0
            section = 0
            
            while current_position < total_height:
                self.driver.execute_script(f"window.scrollTo(0, {current_position});")
                time.sleep(0.3)
                
                temp_filename = f"temp_section_{section}.png"
                temp_path = os.path.join(self.screenshot_dir, temp_filename)
                self.driver.save_screenshot(temp_path)
                screenshots.append({
                    'path': temp_path,
                    'position': current_position,
                    'height': min(viewport_height, total_height - current_position)
                })
                
                print(f"Sección {section} capturada (posición: {current_position}px)")
                
                current_position += viewport_height
                section += 1
                
                if section > 50:
                    print("Demasiadas secciones, cortando captura")
                    break
            
            self.driver.execute_script("window.scrollTo(0, 0);")
            
            if len(screenshots) > 1:
                print(f"Uniendo {len(screenshots)} secciones...")
                final_image = self._merge_screenshots(screenshots, total_width, total_height)
                if final_image:
                    final_image.save(screenshot_path, 'PNG', optimize=True)
                    print(f"Captura completa guardada: {filename}")
                else:
                    Image.open(screenshots[0]['path']).save(screenshot_path)
            else:
                os.rename(screenshots[0]['path'], screenshot_path)
            
            for screenshot in screenshots:
                try:
                    if os.path.exists(screenshot['path']):
                        os.remove(screenshot['path'])
                except:
                    pass
            
            return True
            
        except Exception as e:
            print(f"Error en captura completa: {e}")
            try:
                self.driver.save_screenshot(screenshot_path)
                print(f"Captura de respaldo guardada: {filename}")
                return True
            except:
                print(f"Error crítico al tomar screenshot")
                return False
        finally:
            try:
                self.driver.execute_script(f"window.scrollTo(0, {int(scroll_y)});")
            except Exception:
                pass
            if in_iframe:
                try:
                    self.driver.switch_to.frame(self.current_frame)
                except Exception:
                    pass

    def _merge_screenshots(self, screenshots, total_width, total_height):
        """Une múltiples screenshots en una sola imagen"""
        try:
            final_image = Image.new('RGB', (total_width, total_height))
            y_offset = 0
            
            for i, screenshot_info in enumerate(screenshots):
                try:
                    section_image = Image.open(screenshot_info['path'])
                    section_height = screenshot_info['height']
                    final_image.paste(section_image, (0, y_offset))
                    y_offset += section_height
                    print(f"Sección {i} unida (altura: {section_height}px)")
                except Exception as e:
                    print(f" Error procesando sección {i}: {e}")
                    continue
            
            return final_image
            
        except Exception as e:
            print(f"Error al unir screenshots: {e}")
            return None
    
    def take_landing_screenshot(self, ss_number, stage):
        """Toma screenshot de la landing page completa"""
        filename = f"landing_{stage}_{ss_number}_{self.browser_name}.png"
        result = self.take_full_page_screenshot(filename)
        self._add_url_banner(os.path.join(self.screenshot_dir, filename))
        return result

    def _find_form_region_element(self):
        """Elemento del documento principal que contiene el formulario: el iframe si el
        form está embebido, o el <form> más grande de la página si no hay iframe."""
        if self.current_frame is not None:
            return self.current_frame
        try:
            self.driver.switch_to.default_content()
        except Exception:
            pass
        best, best_area = None, 0
        try:
            for f in self.driver.find_elements(By.TAG_NAME, "form"):
                try:
                    if not f.is_displayed():
                        continue
                    area = f.size.get("width", 0) * f.size.get("height", 0)
                    if area > best_area:
                        best, best_area = f, area
                except Exception:
                    continue
        except Exception:
            pass
        return best

    def take_form_area_screenshot(self, filename):
        """Captura SOLO el área del formulario (no toda la landing, que puede ser larguísima).
        Si el form no entra en el viewport, se toma por partes y se unen."""
        screenshot_path = os.path.join(self.screenshot_dir, filename)
        in_iframe = (self.current_frame is not None)
        scroll_y = 0
        temp_files = []
        try:
            el = self._find_form_region_element()
            if el is None:
                return False
            try:
                self.driver.switch_to.default_content()
            except Exception:
                pass

            scroll_y = self.driver.execute_script("return window.pageYOffset;") or 0
            rect = self.driver.execute_script(
                "var r = arguments[0].getBoundingClientRect();"
                "return {top: r.top + window.pageYOffset, left: r.left + window.pageXOffset,"
                " width: r.width, height: r.height};", el)
            vh = self.driver.execute_script("return window.innerHeight;")
            vw = self.driver.execute_script("return window.innerWidth;")

            top = max(0, int(rect["top"]))
            height = int(rect["height"])
            if height <= 0:
                return False

            sections = []
            pos = top
            idx = 0
            while pos < top + height and idx <= 30:
                self.driver.execute_script(f"window.scrollTo(0, {pos});")
                time.sleep(0.35)
                real_y = self.driver.execute_script("return window.pageYOffset;") or 0
                tmp = os.path.join(self.screenshot_dir, f"temp_form_{idx}.png")
                self.driver.save_screenshot(tmp)
                temp_files.append(tmp)
                img = Image.open(tmp)
                # Escala por devicePixelRatio / zoom del navegador
                scale = img.width / float(vw) if vw else 1.0
                crop_top = max(0, int((pos - real_y) * scale))
                remaining = (top + height) - pos
                crop_h = int(min(vh - (pos - real_y), remaining) * scale)
                if crop_h <= 0:
                    break
                sections.append(img.crop((0, crop_top, img.width, min(img.height, crop_top + crop_h))))
                pos += int(vh - (pos - real_y))
                idx += 1

            if not sections:
                return False

            total_h = sum(s.height for s in sections)
            final = Image.new("RGB", (sections[0].width, total_h), "white")
            y = 0
            for s in sections:
                final.paste(s, (0, y))
                y += s.height
            final.save(screenshot_path, "PNG", optimize=True)
            print(f"Captura del formulario ({len(sections)} parte/s) guardada: {filename}")
            return True
        except Exception as e:
            print(f"Error en captura del área del formulario: {e}")
            return False
        finally:
            for t in temp_files:
                try:
                    os.remove(t)
                except Exception:
                    pass
            try:
                self.driver.execute_script(f"window.scrollTo(0, {int(scroll_y)});")
            except Exception:
                pass
            if in_iframe:
                try:
                    self.driver.switch_to.frame(self.current_frame)
                except Exception:
                    pass

    def take_form_screenshot(self, ss_number, stage, full_page=False):
        """Toma screenshot del formulario dentro del iframe.
        full_page=True → captura del área completa del formulario (por partes si es largo),
        sin arrastrar toda la landing cuando ésta es kilométrica."""
        filename = f"form_{stage}_{ss_number}_{self.browser_name}.png"
        if full_page:
            try:
                if self.take_form_area_screenshot(filename):
                    self._add_url_banner(os.path.join(self.screenshot_dir, filename))
                    return True
                result = self.take_full_page_screenshot(filename)
                self._add_url_banner(os.path.join(self.screenshot_dir, filename))
                print(f"Captura de formulario (página completa) guardada: {filename}")
                return result
            except Exception as e:
                print(f"Error capturando formulario completo, usando viewport: {e}")
        try:
            screenshot_path = os.path.join(self.screenshot_dir, filename)
            if self.browser_name == 'firefox' and self.current_frame is not None:
                # Firefox no puede capturar en contexto iframe: salir, foto, volver
                try:
                    self.driver.switch_to.default_content()
                except Exception:
                    pass
                self.driver.save_screenshot(screenshot_path)
                try:
                    self.driver.switch_to.frame(self.current_frame)
                except Exception:
                    pass
            else:
                self.driver.save_screenshot(screenshot_path)
            self._add_url_banner(screenshot_path)
            print(f"Captura de formulario guardada: {filename}")
            return True
        except Exception as e:
            print(f"Error capturando formulario: {e}")
            return False