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
        # Slug derivado de la landing/form para que el nombre de cada captura diga a qué form
        # pertenece (ej. form_flotas-pesados_paso1_completado_1_chrome.png). Lo setea el runner.
        self.form_slug = ""
        # Líneas extra para el banner: lista de (texto, (r,g,b)) — ej. el Comparador
        # Dealers agrega "Revisado: ..." y "Estado: PASS/FAIL".
        self.extra_lines = []
        # Firefox no puede tomar screenshots en contexto iframe; guardamos el iframe activo
        # para salir a default_content antes del screenshot y volver después.
        self.current_frame = None

    def fname(self, kind, stage, ss_number):
        """Nombre estándar de una captura, con el slug del form si está seteado:
        p.ej. form_flotas-pesados_paso1_completado_3_chrome.png. `kind` = 'form' | 'landing'."""
        slug = (self.form_slug or "").strip()
        slug_part = f"{slug}_" if slug else ""
        return f"{kind}_{slug_part}{stage}_{ss_number}_{self.browser_name}.png"

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
            for extra in (self.extra_lines or []):
                if isinstance(extra, (list, tuple)) and len(extra) == 2:
                    lineas.append((str(extra[0]), tuple(extra[1])))
                else:
                    lineas.append((str(extra), AMARILLO))
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
        """Guarda el PNG pesando lo menos posible sin perder legibilidad. Las capturas son
        UI (formularios/landings: pocos colores planos), así que cuantizar a 256 colores con
        MEDIANCUT reduce el peso ~70-85% sin diferencia visible y mantiene el total del run
        muy por debajo de 50 MB."""
        try:
            if img.mode not in ("RGB", "P"):
                img = img.convert("RGB")
            pal = img.quantize(colors=256, method=Image.MEDIANCUT)
            pal.save(path, "PNG", optimize=True)
        except Exception:
            try:
                img.save(path, "PNG", optimize=True)
            except Exception:
                img.save(path)
        
    def _neutralize_fixed_elements(self):
        """Neutraliza position:fixed/sticky (nav, chat, cookie bars) durante una captura por
        secciones: si no, esos elementos quedan pegados al viewport y aparecen REPETIDOS /
        'arrastrados' en la imagen final. Se restauran con _restore_fixed_elements()."""
        try:
            self.driver.execute_script("""
                (function(){
                  var changed = [];
                  
                  // 1. Ocultar nav y header
                  var navs = document.querySelectorAll('nav, header, [class*="nav-bar"], [class*="navbar"], [class*="header"]');
                  for (var j=0; j<navs.length; j++){
                    try {
                      var curDisplay = navs[j].style.display;
                      var priority = navs[j].style.getPropertyPriority('display');
                      changed.push({
                        element: navs[j],
                        property: 'display',
                        value: curDisplay,
                        priority: priority
                      });
                      navs[j].style.setProperty('display', 'none', 'important');
                    } catch(e){}
                  }
                  
                  // 2. Neutralizar otros elementos fixed/sticky
                  var all = document.querySelectorAll('body *');
                  for (var i=0; i<all.length; i++){
                    try {
                      var cs = window.getComputedStyle(all[i]);
                      if (cs && (cs.position === 'fixed' || cs.position === 'sticky' || cs.position === '-webkit-sticky')) {
                        // Skip if already hidden as nav/header
                        var alreadyHidden = false;
                        for (var k=0; k<changed.length; k++){
                          if (changed[k].element === all[i] && changed[k].property === 'display') {
                            alreadyHidden = true;
                            break;
                          }
                        }
                        if (alreadyHidden) continue;
                        
                        var curPos = all[i].style.position;
                        var priority = all[i].style.getPropertyPriority('position');
                        changed.push({
                          element: all[i],
                          property: 'position',
                          value: curPos,
                          priority: priority
                        });
                        all[i].style.setProperty('position', 'static', 'important');
                      }
                    } catch(e){}
                  }
                  
                  window.__osocio_style_changes = changed;
                  return changed.length;
                })();
            """)
        except Exception as e:
            print(f"Error neutralizando elementos: {e}")

    def _restore_fixed_elements(self):
        """Restaura el estilo original de los elementos neutralizados."""
        try:
            self.driver.execute_script("""
                (function(){
                  var changes = window.__osocio_style_changes || [];
                  for (var i=0; i<changes.length; i++){
                    try {
                      var item = changes[i];
                      if (item.value) {
                        item.element.style.setProperty(item.property, item.value, item.priority || '');
                      } else {
                        item.element.style.removeProperty(item.property);
                      }
                    } catch(e){}
                  }
                  window.__osocio_style_changes = null;
                })();
            """)
        except Exception as e:
            print(f"Error restaurando elementos: {e}")

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

            # Neutralizar nav/chat/cookie fijos ANTES de medir, para que no se arrastren.
            self._neutralize_fixed_elements()

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
            self._restore_fixed_elements()
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
        filename = self.fname("landing", stage, ss_number)
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

    def _medir_scroll_interno_iframe(self):
        """
        ¿El form vive en un iframe de altura fija con scroll propio adentro?

        Si es así, el rectángulo del iframe en la página padre es solo la "ventanita" visible:
        capturar eso deja el formulario cortado (se ven los últimos campos y el botón Enviar,
        pero no los de arriba). Devuelve {'content', 'client'} del documento del iframe, o None.
        Deja el scroll interno del iframe en 0.
        """
        if self.current_frame is None:
            return None
        try:
            # OJO: hay que volver al documento principal ANTES de entrar al iframe. Si el
            # driver ya estaba adentro, switch_to.frame() con un elemento del padre falla.
            self.driver.switch_to.default_content()
            self.driver.switch_to.frame(self.current_frame)
            info = self.driver.execute_script("""
                window.scrollTo(0, 0);
                var d = document.documentElement, b = document.body;
                return {content: Math.max(d.scrollHeight || 0, b ? b.scrollHeight : 0),
                        client:  d.clientHeight || 0};
            """)
            return info
        except Exception:
            return None
        finally:
            try:
                self.driver.switch_to.default_content()
            except Exception:
                pass

    def _scroll_dentro_iframe(self, y):
        """Scrollea el documento DE ADENTRO del iframe."""
        try:
            self.driver.switch_to.default_content()
            self.driver.switch_to.frame(self.current_frame)
            self.driver.execute_script(f"window.scrollTo(0, {int(y)});")
            time.sleep(0.25)
            return self.driver.execute_script("return window.pageYOffset;") or 0
        except Exception:
            return 0
        finally:
            try:
                self.driver.switch_to.default_content()
            except Exception:
                pass

    def _capturar_iframe_scrolleando_adentro(self, screenshot_path, el, contenido, visible):
        """
        Captura un formulario que vive en un iframe con scroll interno: va scrolleando DENTRO
        del iframe y pega las partes, para que queden TODOS los campos y no solo los visibles.
        """
        temp_files = []
        try:
            vw = self.driver.execute_script("return window.innerWidth;")
            partes = []
            capturado = 0     # px del contenido del iframe que ya están en la imagen final
            idx = 0
            while capturado < contenido and idx <= 30:
                real = self._scroll_dentro_iframe(capturado)

                # El iframe tiene que estar en pantalla para poder fotografiarlo
                self.driver.execute_script(
                    "arguments[0].scrollIntoView({block:'center', behavior:'instant'});", el)
                time.sleep(0.25)

                rect = self.driver.execute_script(
                    "var r = arguments[0].getBoundingClientRect();"
                    "return {top: r.top, left: r.left, width: r.width, height: r.height};", el)

                tmp = os.path.join(self.screenshot_dir, f"temp_iframe_{idx}.png")
                self.driver.save_screenshot(tmp)
                temp_files.append(tmp)
                img = Image.open(tmp)
                scale = img.width / float(vw) if vw else 1.0

                # Al llegar al fondo, el iframe ya no puede scrollear lo que le pedimos:
                # queda solapado con lo anterior. Sin descontarlo, el form sale DUPLICADO.
                solapa = max(0, capturado - real)
                nuevo = min(visible - solapa, contenido - capturado)
                if nuevo <= 0:
                    break

                top = int((max(0, rect["top"]) + solapa) * scale)
                alto = int(nuevo * scale)
                partes.append(img.crop((0, top, img.width, min(img.height, top + alto))))

                capturado += nuevo
                idx += 1

            if not partes:
                return False

            total_h = sum(p.height for p in partes)
            final = Image.new("RGB", (partes[0].width, total_h), "white")
            y = 0
            for p in partes:
                final.paste(p, (0, y))
                y += p.height
            final.save(screenshot_path, "PNG", optimize=True)
            print(f"Captura del formulario ({len(partes)} parte/s, scroll dentro del iframe): "
                  f"{os.path.basename(screenshot_path)}")
            return True
        except Exception as e:
            print(f"Error capturando el iframe por partes: {e}")
            return False
        finally:
            for t in temp_files:
                try:
                    os.remove(t)
                except Exception:
                    pass
            self._scroll_dentro_iframe(0)

    def take_form_area_screenshot(self, filename):
        """Captura SOLO el área del formulario (no toda la landing, que puede ser larguísima).
        Si el form no entra en el viewport, se toma por partes y se unen."""
        screenshot_path = os.path.join(self.screenshot_dir, filename)
        in_iframe = (self.current_frame is not None)
        scroll_y = 0
        temp_files = []
        try:
            # Neutralizar nav/chat fijos del documento padre para que no tapen el form al
            # scrollear por secciones.
            try:
                self.driver.switch_to.default_content()
            except Exception:
                pass
            self._neutralize_fixed_elements()

            el = self._find_form_region_element()
            if el is None:
                return False

            # Caso iframe de altura fija con scroll propio: hay que scrollear ADENTRO, si no
            # la foto sale cortada (solo los últimos campos + el botón Enviar).
            interno = self._medir_scroll_interno_iframe()
            if interno and interno.get("client") and \
                    interno["content"] > interno["client"] + 5:
                try:
                    self.driver.switch_to.default_content()
                except Exception:
                    pass
                if self._capturar_iframe_scrolleando_adentro(
                        screenshot_path, el, interno["content"], interno["client"]):
                    if in_iframe:
                        try:
                            self.driver.switch_to.frame(self.current_frame)
                        except Exception:
                            pass
                    return True

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
            # Restaurar los fijos (el neutralizado se hizo sobre el documento padre).
            try:
                self.driver.switch_to.default_content()
            except Exception:
                pass
            self._restore_fixed_elements()
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
        filename = self.fname("form", stage, ss_number)
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

def enforce_screenshot_budget(screenshot_dir, max_mb=48, log=print):
    """Garantiza que TODAS las capturas de la carpeta pesen en total menos de `max_mb` MB,
    perdiendo la MENOR calidad posible: solo actúa si se pasa del presupuesto, y ahí recomprime
    (y como último recurso reescala) las imágenes MÁS pesadas primero. Las capturas chicas
    (área del form) casi nunca se tocan; el peso está en las landings completas kilométricas."""
    max_bytes = int(max_mb * 1024 * 1024)
    try:
        files = [os.path.join(screenshot_dir, f) for f in os.listdir(screenshot_dir)
                 if f.lower().endswith(".png")]
    except Exception:
        return

    def _total():
        return sum(os.path.getsize(f) for f in files if os.path.exists(f))

    total0 = _total()
    if total0 <= max_bytes:
        return  # ya entra: no se toca nada (calidad intacta)

    log(f"Capturas: {total0/1e6:.1f} MB > {max_mb} MB — recomprimiendo las más pesadas...")

    # Hasta N pasadas: cada una ataca el archivo más pesado. Primero recomprime (cuantiza a
    # menos colores, casi sin pérdida visible en UI); si ya está muy cuantizado, lo reescala.
    for _ in range(60):
        if _total() <= max_bytes:
            break
        try:
            f = max((p for p in files if os.path.exists(p)), key=os.path.getsize)
        except ValueError:
            break
        try:
            img = Image.open(f)
            w, h = img.size
            mode = img.mode
            # 1) Si todavía es RGB (no cuantizada) o muy alta → cuantizar a 128 colores.
            if mode != "P":
                out = img.convert("RGB").quantize(colors=128, method=Image.MEDIANCUT)
            elif h > 4000 or w > 1100:
                # 2) Ya cuantizada y grande → reescalar 82% (sigue legible) y recuantizar.
                nw, nh = max(1, int(w * 0.82)), max(1, int(h * 0.82))
                out = img.convert("RGB").resize((nw, nh), Image.LANCZOS).quantize(colors=128, method=Image.MEDIANCUT)
            else:
                # 3) Chica y ya cuantizada → bajar a 64 colores.
                out = img.convert("RGB").quantize(colors=64, method=Image.MEDIANCUT)
            out.save(f, "PNG", optimize=True)
        except Exception:
            # Si una imagen falla, sacarla de la lista para no ciclar en ella.
            files = [p for p in files if p != f]

    log(f"Capturas: total final {_total()/1e6:.1f} MB")
