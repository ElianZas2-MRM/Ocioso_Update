"""
lt_screenshots.py
=================
Sistema de capturas dedicado para LambdaTest Mac.

REALIDAD TÉCNICA
----------------
driver.save_screenshot() en LambdaTest = solo el viewport interno del browser.
NO incluye la barra de Chrome ni el frame del OS Mac. Eso es una limitación
de WebDriver — no hay forma de evitarlo durante la sesión activa.

LO QUE SÍ FUNCIONA
-------------------
lambda-screenshot=true + API de LT AL CERRAR LA SESIÓN.

Entonces la estrategia es:

  DURANTE la sesión:
    - En cada momento de captura, ejecutar lambda-screenshot=true
      para que LT registre esa captura con el frame Mac en su servidor.
    - También guardar con save_screenshot() para tener la imagen
      local inmediata (sin frame, pero disponible al instante).

  AL CERRAR el driver (driver.quit()):
    - Descargar TODOS los lambda-screenshots via API de LT.
    - Renombrarlos con nombres descriptivos en orden:
        01_landing_inicial.png
        02_form_errores.png
        03_paso_01.png
        04_paso_02.png   (si hay)
        05_form_completado.png
        06_ty_page.png
    - Guardarlos en screenshots_dir/con_frame_mac/

RESULTADO FINAL
---------------
screenshots_dir/
  ├── landing_inserto_1_chrome.png     ← local inmediata (sin frame)
  ├── form_errores_1_chrome.png        ← local inmediata (sin frame)
  ├── form_paso_01_1_chrome.png        ← local inmediata (sin frame)
  ├── form_completado_1_chrome.png     ← local inmediata (sin frame)
  ├── landing_typage_1_chrome.png      ← local inmediata (sin frame)
  └── con_frame_mac/
        ├── 01_landing_inicial.png     ← LT, CON frame Mac
        ├── 02_form_errores.png        ← LT, CON frame Mac
        ├── 03_paso_01.png             ← LT, CON frame Mac
        ├── 04_paso_02.png             ← LT, CON frame Mac (si hay 2 pasos)
        ├── 05_form_completado.png     ← LT, CON frame Mac
        └── 06_ty_page.png             ← LT, CON frame Mac
"""

import os
import time
from typing import Callable, Optional

try:
    import requests as _req
    _REQUESTS_OK = True
except ImportError:
    _REQUESTS_OK = False

_LT_API_BASE = "https://api.lambdatest.com/automation/api/v1/"


# ══════════════════════════════════════════════════════════════════════════════
# CLASE PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════

class LTScreenshotManager:
    """
    Gestiona todas las capturas de una sesión de LambdaTest.

    Uso:
        sm = LTScreenshotManager(driver, screenshots_dir, lead_num,
                                  session_id, username, access_key,
                                  iframe_src, log)

        sm.captura_landing_inicial()
        sm.captura_form_errores()
        sm.captura_paso(1)
        sm.captura_paso(2)       # solo si hay más pasos
        sm.captura_form_completado()
        sm.captura_ty_page()

        # Al final, después de driver.quit():
        sm.descargar_con_frame_mac()
    """

    def __init__(self, driver, screenshots_dir: str, lead_num: int,
                 session_id: str, username: str, access_key: str,
                 iframe_src: str = "", log: Callable = print,
                 landing_url: str = "", actual_iframe_src: str = ""):
        self.driver            = driver
        self.screenshots_dir   = screenshots_dir
        self.lead_num          = lead_num
        self.session_id        = session_id
        self.username          = username
        self.access_key        = access_key
        self.iframe_src        = iframe_src        # URL esperada del form (Excel col B)
        self.landing_url       = landing_url       # URL de la landing
        self.actual_iframe_src = actual_iframe_src # URL real encontrada en el iframe
        self.log               = log

        # Registro ordenado de capturas LT para renombrar al final
        # Cada entrada: (orden, nombre_descriptivo)
        self._lt_sequence = []
        self._lt_count    = 0  # contador de lambda-screenshots disparados

        os.makedirs(screenshots_dir, exist_ok=True)

    # ── Capturas locales (save_screenshot) ────────────────────────────────────

    def _add_url_banner(self, path: str):
        """Agrega un banner oscuro en la parte superior de la imagen con las URLs."""
        try:
            from PIL import Image, ImageDraw, ImageFont
            img = Image.open(path).convert("RGB")
            w, h = img.size

            AMARILLO = (220, 220, 100)
            VERDE    = (80, 200, 80)
            ROJO     = (220, 80, 80)
            lineas = []  # (texto, color)
            if self.landing_url:
                lineas.append((f"Landing:         {self.landing_url}", AMARILLO))
            if self.iframe_src:
                lineas.append((f"Form esperado:   {self.iframe_src}", AMARILLO))
            if self.actual_iframe_src:
                lineas.append((f"Form encontrado: {self.actual_iframe_src}", AMARILLO))
                if self.iframe_src:
                    if self.actual_iframe_src == self.iframe_src:
                        lineas.append(("  ✓  Coincide con el esperado", VERDE))
                    else:
                        lineas.append(("  ✗  NO coincide con el esperado", ROJO))
            if not lineas:
                return

            font_size  = max(14, w // 80)
            line_h     = font_size + 6
            banner_h   = line_h * len(lineas) + 10
            banner     = Image.new("RGB", (w, banner_h), (30, 30, 30))
            draw       = ImageDraw.Draw(banner)
            try:
                font = ImageFont.truetype("arial.ttf", font_size)
            except Exception:
                font = ImageFont.load_default()
            for i, (line, color) in enumerate(lineas):
                draw.text((8, 5 + i * line_h), line, fill=color, font=font)

            combined = Image.new("RGB", (w, banner_h + h))
            combined.paste(banner, (0, 0))
            combined.paste(img, (0, banner_h))
            combined.save(path)
        except Exception as e:
            self.log(f"  ⚠ Error banner URLs: {e}")

    def _local(self, stage: str) -> str:
        """
        Toma captura local con save_screenshot.
        Sale al default_content, centra el iframe, captura.
        Retorna el nombre del archivo.
        """
        filename = f"{stage}_{self.lead_num}_chrome.png"
        path     = os.path.join(self.screenshots_dir, filename)

        try:
            # Salir al contexto principal
            try:
                self.driver.switch_to.default_content()
            except Exception:
                pass
            time.sleep(0.3)

            # Centrar el iframe en pantalla
            self._center_iframe()

            # Captura local
            self.driver.save_screenshot(path)
            self._add_url_banner(path)
            self.log(f"  📸 Local: {filename}")
        except Exception as e:
            self.log(f"  ⚠ Error captura local {stage}: {e}")

        return filename

    def _center_iframe(self):
        """Centra el iframe del formulario en el viewport."""
        try:
            from selenium.webdriver.common.by import By
            iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
            iframe_el = None

            # Buscar por URL exacta primero
            for ifr in iframes:
                src = ifr.get_attribute("src") or ""
                if self.iframe_src and self.iframe_src.strip() in src:
                    iframe_el = ifr
                    break

            # Fallback: primer gm_forms/gm_admin
            if not iframe_el:
                for ifr in iframes:
                    src = ifr.get_attribute("src") or ""
                    if "gm_forms" in src or "gm_admin" in src:
                        iframe_el = ifr
                        break

            if iframe_el:
                self.driver.execute_script("""
                    const el = arguments[0];
                    const rect = el.getBoundingClientRect();
                    const center = rect.top + window.scrollY + (rect.height / 2);
                    const vp = window.innerHeight / 2;
                    window.scrollTo({ top: Math.max(0, center - vp), behavior: 'instant' });
                """, iframe_el)
                time.sleep(0.4)
            else:
                self.driver.execute_script("window.scrollTo(0, 0);")
                time.sleep(0.2)
        except Exception:
            pass

    def _volver_iframe(self):
        """Vuelve al iframe del formulario después de una captura."""
        try:
            from selenium.webdriver.common.by import By
            iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
            for ifr in iframes:
                src = ifr.get_attribute("src") or ""
                if (self.iframe_src and self.iframe_src.strip() in src) or \
                   "gm_forms" in src or "gm_admin" in src:
                    self.driver.switch_to.frame(ifr)
                    return True
        except Exception:
            pass
        return False

    # ── Trigger lambda-screenshot ──────────────────────────────────────────────

    def _trigger(self, nombre_descriptivo: str):
        """
        Ejecuta lambda-screenshot=true y registra el orden para el renombrado final.
        """
        self._lt_count += 1
        orden = self._lt_count
        self._lt_sequence.append((orden, nombre_descriptivo))
        try:
            self.driver.execute_script("lambda-screenshot=true")
            self.log(f"  🎯 LT screenshot #{orden}: {nombre_descriptivo}")
        except Exception as e:
            self.log(f"  ⚠ Error lambda-screenshot: {e}")

    # ── API PÚBLICA — una función por momento de captura ──────────────────────

    def captura_landing_inicial(self) -> str:
        """
        Captura la landing page con el iframe visible.
        Contexto: default_content (ya estamos fuera del iframe).
        """
        self._trigger("01_landing_inicial")
        return self._local("landing_inserto")

    def captura_form_errores(self) -> str:
        """
        Captura el formulario con los mensajes de error (envío vacío).
        Contexto: dentro del iframe. Sale, captura, VUELVE al iframe.
        """
        self._trigger("02_form_errores")
        filename = self._local("form_errores")
        self._volver_iframe()
        return filename

    def captura_paso(self, num_paso: int) -> str:
        """
        Captura el formulario con los campos llenados del paso N.
        Contexto: dentro del iframe. Sale, captura, VUELVE al iframe.
        """
        nombre_lt = f"{num_paso + 2:02d}_paso_{num_paso:02d}"
        self._trigger(nombre_lt)
        filename = self._local(f"form_paso_{num_paso:02d}")
        self._volver_iframe()
        return filename

    def captura_form_completado(self) -> str:
        """
        Captura el formulario completamente llenado, antes de enviar.
        Contexto: dentro del iframe. Sale, captura, VUELVE al iframe.
        """
        # El número de orden depende de cuántos pasos hubo
        orden = self._lt_count + 1
        self._trigger(f"{orden:02d}_form_completado")
        filename = self._local("form_completado")
        self._volver_iframe()
        return filename

    def captura_ty_page(self) -> str:
        """
        Captura la página de thank you después del envío.
        Contexto: default_content.
        """
        orden = self._lt_count + 1
        self._trigger(f"{orden:02d}_ty_page")
        return self._local("landing_typage")

    # ── Descarga con frame Mac al final ───────────────────────────────────────

    def descargar_con_frame_mac(self) -> int:
        """
        Descarga todos los lambda-screenshots de la sesión terminada via API.
        Los renombra con los nombres descriptivos registrados en _lt_sequence.
        Los guarda en screenshots_dir/con_frame_mac/

        Llamar DESPUÉS de driver.quit().
        Retorna la cantidad de imágenes descargadas.
        """
        if not _REQUESTS_OK:
            self.log("  ⚠ requests no disponible — no se pueden descargar screenshots LT")
            return 0

        if not self.session_id:
            self.log("  ⚠ Sin session_id — no se pueden descargar screenshots LT")
            return 0

        mac_folder = os.path.join(self.screenshots_dir, "con_frame_mac")
        os.makedirs(mac_folder, exist_ok=True)

        url  = f"{_LT_API_BASE}sessions/{self.session_id}/screenshots"
        auth = (self.username, self.access_key)

        self.log(f"\n  📥 Descargando {self._lt_count} capturas con frame Mac de LT...")
        self.log(f"     (disponibles ~30s después de cerrar la sesión)")

        # LT necesita tiempo para procesar — reintentar hasta 6 veces con 10s
        screenshots = []
        for attempt in range(1, 7):
            try:
                resp = _req.get(url, auth=auth, timeout=30)
                if resp.status_code != 200:
                    self.log(f"  ⏳ API status={resp.status_code} (intento {attempt}/6)")
                    time.sleep(10)
                    continue

                try:
                    data = resp.json()
                except Exception:
                    self.log(f"  ⏳ Respuesta no es JSON (intento {attempt}/6)")
                    time.sleep(10)
                    continue

                screenshots = (
                    data.get("data") or
                    data.get("screenshots") or
                    data.get("result") or
                    []
                )

                if len(screenshots) >= self._lt_count:
                    self.log(f"  ✓ {len(screenshots)} screenshots disponibles en LT")
                    break
                else:
                    self.log(
                        f"  ⏳ {len(screenshots)}/{self._lt_count} disponibles "
                        f"(intento {attempt}/6)"
                    )
                    time.sleep(10)

            except Exception as e:
                self.log(f"  ⚠ Error API (intento {attempt}/6): {e}")
                time.sleep(10)

        if not screenshots:
            self.log("  ⚠ No se encontraron screenshots en LT para esta sesión")
            return 0

        # Descargar y renombrar según el orden registrado
        downloaded = 0
        for (orden, nombre), ss in zip(self._lt_sequence, screenshots):
            img_url = (
                ss.get("screenshot_url") or
                ss.get("url") or
                ss.get("image_url") or
                ""
            )
            if not img_url:
                continue
            try:
                img_resp = _req.get(img_url, timeout=30)
                if img_resp.status_code == 200 and img_resp.content:
                    # Nombre: 01_landing_inicial_lead1.png
                    fname = f"{nombre}_lead{self.lead_num}.png"
                    dest  = os.path.join(mac_folder, fname)
                    with open(dest, "wb") as f:
                        f.write(img_resp.content)
                    self.log(f"  ✓ {fname}")
                    downloaded += 1
                else:
                    self.log(f"  ⚠ Error descargando {nombre}: {img_resp.status_code}")
            except Exception as e:
                self.log(f"  ⚠ Error {nombre}: {e}")

        if downloaded:
            self.log(f"\n  ✅ {downloaded} capturas con frame Mac en: con_frame_mac/")
        return downloaded
