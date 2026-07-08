"""
ty_cta.py — Detección e investigación del CTA y de links "raros" en la Thank-You page.

Fuente única compartida por el motor desktop (base_form_filler) y los runners de
LambdaTest (Mac / Android). Aplica a forms gm_forms / gm_front (todo lo que no es
forms 2.0 / AEM, que tienen su propio flujo).

Produce info para DOS columnas del Excel:
- "TYP con CTA": si hay un <a> CTA real → "SÍ | texto -> href (target) | llegó a: URL
  | captura: <archivo>". Se verifica el destino con click real (pestaña aparte).
- "LINK ISSUE TYP": si hay un link "raro" (atributo href/src con HTML roto/inyectado,
  ej. termina en "</span" — el link que aparece sobre el texto y no es un <a> normal):
  describe dónde está (captura con el elemento señalado en ROJO) y a dónde lleva
  (captura con banner "LINKEO EXTRAÑO LLEVO A: ...").

En LambdaTest (take_screenshot=False) NO se guardan capturas — el click y la landing
quedan grabados en el video; igual se reporta el texto/URL destino.
"""
import os
import time

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait


def add_banner(path, text, log=print):
    """Banner oscuro con `text` arriba de la imagen — mismo estilo que el resto de
    las capturas (lt_screenshots.py._add_url_banner)."""
    if not text:
        return
    try:
        from PIL import Image, ImageDraw, ImageFont
        img = Image.open(path).convert("RGB")
        w, h = img.size
        font_size = max(14, w // 80)
        line_h = font_size + 6
        banner_h = line_h + 10
        banner = Image.new("RGB", (w, banner_h), (30, 30, 30))
        draw = ImageDraw.Draw(banner)
        try:
            font = ImageFont.truetype("arial.ttf", font_size)
        except Exception:
            font = ImageFont.load_default()
        draw.text((8, 5), text, fill=(220, 220, 100), font=font)
        combined = Image.new("RGB", (w, banner_h + h))
        combined.paste(banner, (0, 0))
        combined.paste(img, (0, banner_h))
        combined.save(path)
    except Exception as e:
        log(f"  ⚠ Error agregando banner a captura: {e}")


def _target_label(target):
    t = (target or "").strip().lower()
    if t == "_blank":
        return "Pestaña nueva (target=_blank)"
    if t == "_parent":
        return "Misma pestaña (target=_parent)"
    if t == "_self":
        return "Misma pestaña (target=_self)"
    if not t:
        return "Misma pestaña (sin target)"
    return f"Misma pestaña (target={t})"


def _open_and_capture(driver, url, banner_text, evidence_dir, take_screenshot, log, fname_prefix="cta"):
    """Abre `url` en una pestaña aparte, devuelve (landed_url, screenshot_path).
    Deja el foco de vuelta en la pestaña original. screenshot_path="" si no se captura."""
    landed_url, shot = "", ""
    original_handles = driver.window_handles
    try:
        driver.execute_script("window.open(arguments[0], '_blank');", url)
        WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > len(original_handles))
        new_handle = [h for h in driver.window_handles if h not in original_handles][0]
        driver.switch_to.window(new_handle)
        try:
            WebDriverWait(driver, 10).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
        except Exception:
            pass
        time.sleep(1.5)
        landed_url = driver.current_url or ""
        log(f"  ✓ {banner_text} {landed_url}")
        if take_screenshot:
            _edir = evidence_dir or os.getcwd()
            os.makedirs(_edir, exist_ok=True)
            shot = os.path.join(_edir, f"{fname_prefix}_{int(time.time()*1000)}.png")
            driver.save_screenshot(shot)
            add_banner(shot, f"{banner_text} {landed_url}", log)
    except Exception as e:
        log(f"  ⚠ Error abriendo/verificando link: {e}")
    finally:
        try:
            if driver.current_window_handle != original_handles[0]:
                driver.close()
            driver.switch_to.window(original_handles[0])
        except Exception:
            pass
    return landed_url, shot


def investigate_ty_cta(driver, log=print, evidence_dir=None, take_screenshot=True):
    """Devuelve dict con la info de CTA(s) y de link raro de la TY page. Puede haber
    más de un <a> en la TY: se reportan y verifican todos (dedup por href)."""
    info = {
        "has_cta": False, "ctas": [],
        "has_weird": False, "weird_desc": "", "weird_href": "",
        "weird_location_shot": "", "weird_landed_url": "", "weird_landed_shot": "",
    }
    try:
        data = driver.execute_script(r"""
            var out = {links: [], weird: []};
            var containers = document.querySelectorAll('div#thank-you, div.rp-wrapper');
            for (var i=0; i<containers.length; i++) {
                var c = containers[i];
                var as = c.querySelectorAll('a[href]');
                for (var j=0; j<as.length; j++) {
                    var a = as[j];
                    var href = a.getAttribute('href') || '';
                    if (!href || href.indexOf('javascript:') === 0 || href === '#') continue;
                    out.links.push({text:(a.textContent||'').trim(), href:a.href, target:a.getAttribute('target')||''});
                }
            }
            // Link "raro": atributo href/data-href/data-url/src cuyo valor tiene pinta de
            // HTML inyectado/mal escapado (contiene un tag literal tipo "</span"). Se marca
            // el primero con un data-attr para poder señalarlo/navegarlo luego.
            var attrs = ['href', 'data-href', 'data-url', 'src'];
            var candidates = document.querySelectorAll('[href], [data-href], [data-url]');
            for (var m = 0; m < candidates.length; m++) {
                var el = candidates[m];
                for (var n = 0; n < attrs.length; n++) {
                    var v = el.getAttribute(attrs[n]);
                    if (v && (v.indexOf('</') >= 0 || v.indexOf('&lt;') >= 0 || v.indexOf('<span') >= 0)) {
                        if (out.weird.length === 0) el.setAttribute('data-weird-cta-marker', '1');
                        var resolved = '';
                        try { resolved = el.href || ''; } catch(e) {}
                        out.weird.push({tag: el.tagName, attr: attrs[n], value: v.slice(0, 200),
                                        text: (el.textContent || '').trim().slice(0, 80), resolved: resolved});
                    }
                }
            }
            return out;
        """) or {"links": [], "weird": []}
    except Exception as e:
        log(f"  ⚠ Error buscando CTA/links en TY: {e}")
        return info

    links = data.get("links") or []
    weird = data.get("weird") or []

    # ── Link raro: captura de UBICACIÓN primero (in-page, sin cambiar de pestaña) ──
    # Debe ir ANTES de cualquier _open_and_capture: abrir/cerrar pestañas resetea el
    # contexto de frame, y el marcador del link raro vive en el iframe del form.
    if weird:
        info["has_weird"] = True
        info["weird_desc"] = "; ".join(
            f"{w.get('tag')}[{w.get('attr')}] '{w.get('text','')}' = {w.get('value','')}"
            for w in weird
        )[:500]
        info["weird_href"] = (weird[0].get("resolved") or "").strip()
        log(f"  ⚠ Link raro (href roto/inyectado) en la TY: {info['weird_desc']}")
        if take_screenshot:
            try:
                el = driver.find_element(By.CSS_SELECTOR, '[data-weird-cta-marker="1"]')
                driver.execute_script(
                    "arguments[0].scrollIntoView({block:'center'});"
                    "arguments[0].style.outline='4px solid red';"
                    "arguments[0].style.outlineOffset='2px';", el)
                time.sleep(0.3)
                _edir = evidence_dir or os.getcwd()
                os.makedirs(_edir, exist_ok=True)
                shot = os.path.join(_edir, f"linkissue_ubicacion_{int(time.time()*1000)}.png")
                driver.save_screenshot(shot)
                add_banner(shot, "LINKEO EXTRAÑO — señalado en rojo", log)
                info["weird_location_shot"] = shot
                log(f"  📸 Ubicación del link raro: {shot}")
            except Exception as e:
                log(f"  ⚠ Error señalando link raro: {e}")

    # En browsers (take_screenshot=True) navegamos los links para dejar evidencia;
    # antes salimos al top-frame porque window.open desde el iframe cross-origin del
    # form es frágil. En LambdaTest (take_screenshot=False) NO navegamos — abrir/cerrar
    # pestañas rompería el contexto del iframe que el runner necesita después; sólo se
    # reporta texto/href/target (lo demás del flujo queda grabado en el video).
    if take_screenshot:
        try:
            driver.switch_to.default_content()
        except Exception:
            pass

    # ── CTA(s) real(es) — puede haber más de un <a> ──────────────────────────
    _seen = set()
    for lk in links:
        href = lk.get("href") or ""
        if not href or href in _seen:
            continue
        _seen.add(href)
        cta = {
            "text": lk.get("text") or "(sin texto)",
            "href": href,
            "target": _target_label(lk.get("target")),
            "landed_url": "", "screenshot": "",
        }
        log(f"  🔗 CTA en TY: '{cta['text']}' -> {href} — segun target abre en: {cta['target']}")
        if take_screenshot:
            cta["landed_url"], cta["screenshot"] = _open_and_capture(
                driver, href, "CTA LLEVÓ A:", evidence_dir, take_screenshot, log
            )
        info["ctas"].append(cta)
    info["has_cta"] = bool(info["ctas"])

    # ── Link raro: navegación (a dónde lleva) — sólo browsers ──────────────────
    if take_screenshot and weird and info["weird_href"]:
        info["weird_landed_url"], info["weird_landed_shot"] = _open_and_capture(
            driver, info["weird_href"], "LINKEO EXTRAÑO LLEVÓ A:", evidence_dir,
            take_screenshot, log, fname_prefix="linkissue_destino"
        )

    return info


def _basename(p):
    return os.path.basename(p) if p else ""


def format_ty_cta(info):
    """Columna 'TYP con CTA'. Si hay varios <a>, se listan separados por ' || '."""
    if not info or not info.get("has_cta"):
        return "NO"
    ctas = info.get("ctas") or []
    chunks = []
    for cta in ctas:
        p = [f"{cta['text']} -> {cta['href']} ({cta['target']})"]
        if cta.get("landed_url"):
            p.append(f"llegó a: {cta['landed_url']}")
        if cta.get("screenshot"):
            p.append(f"captura: {_basename(cta['screenshot'])}")
        chunks.append(" | ".join(p))
    prefix = "SÍ" if len(chunks) == 1 else f"SÍ ({len(chunks)} links)"
    return f"{prefix} || " + " || ".join(chunks)


def format_link_issue(info):
    """Columna 'LINK ISSUE TYP'."""
    if not info or not info.get("has_weird"):
        return "-"
    parts = [f"SÍ | {info['weird_desc']}"]
    if info.get("weird_location_shot"):
        parts.append(f"ubicación: {_basename(info['weird_location_shot'])}")
    if info.get("weird_landed_url"):
        parts.append(f"llevó a: {info['weird_landed_url']}")
    if info.get("weird_landed_shot"):
        parts.append(f"captura: {_basename(info['weird_landed_shot'])}")
    return " | ".join(parts)
