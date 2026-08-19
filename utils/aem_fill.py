"""
aem_fill.py — Llenado de Adobe AEM Adaptive Forms ("formularios T3 2.0").

Fuente única compartida por el motor desktop (base_form_filler) y los runners de
LambdaTest (Mac / Android). Funciones puras basadas en `driver` (Selenium), sin estado.

En estos forms el ID del widget es genérico (guidetextbox_1763313) y trae un panel
volátil; el keyword semántico (cpf, email, celular...) vive en el <label for>. Por eso
se localiza cada campo por el texto del label, no por ID fijo.
"""
import random
import time
import unicodedata
from urllib.parse import parse_qs, unquote, urlsplit

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import StaleElementReferenceException


def _cpf_checksum_ok(digits: str) -> bool:
    """Valida el dígito verificador real de un CPF de 11 dígitos."""
    if len(digits) != 11 or len(set(digits)) == 1 or not digits.isdigit():
        return False
    d = [int(c) for c in digits]
    r1 = sum(v * (10 - i) for i, v in enumerate(d[:9])) % 11
    c1 = 0 if r1 < 2 else 11 - r1
    if c1 != d[9]:
        return False
    r2 = sum(v * (11 - i) for i, v in enumerate(d[:10])) % 11
    c2 = 0 if r2 < 2 else 11 - r2
    return c2 == d[10]


def _cnpj_checksum_ok(digits: str) -> bool:
    """Valida el dígito verificador real de un CNPJ de 14 dígitos."""
    if len(digits) != 14 or len(set(digits)) == 1 or not digits.isdigit():
        return False
    d = [int(c) for c in digits]
    w1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    r1 = sum(v * w for v, w in zip(d[:12], w1)) % 11
    c1 = 0 if r1 < 2 else 11 - r1
    if c1 != d[12]:
        return False
    w2 = [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    r2 = sum(v * w for v, w in zip(d[:12] + [c1], w2)) % 11
    c2 = 0 if r2 < 2 else 11 - r2
    return c2 == d[13]


def normalize_text(value):
    """Minúsculas sin tildes."""
    if not value:
        return ""
    nfkd = unicodedata.normalize("NFKD", str(value).lower())
    return "".join(c for c in nfkd if not unicodedata.combining(c)).strip()


_PLACEHOLDER_KEYWORDS = (
    "seleccione", "selecciona", "seleccionar", "selecione", "selecionar",
    "escolha", "escolher", "elija", "elegir", "opcao", "opcion",
    "select", "choose", "please", "favor",
)


def _is_placeholder(text):
    t = normalize_text(text)
    return (not t) or any(k in t for k in _PLACEHOLDER_KEYWORDS)


# El orden importa: entradas específicas primero para evitar colisiones de substring
# ("Nome da empresa" contiene "nome"; "Número de veículos" contiene "veículo").
# spec: ("data", data_key) | ("brasil", cpf|cnpj|cep) | ("gen", company|amount)
_FIELD_KEYWORDS = (
    (("sobrenome", "apellido", "lastname", "surname"),              ("data", "lastname")),
    (("empresa", "razaosocial", "razao social", "nome fantasia", "company"), ("gen", "company")),
    (("nome", "nombre", "firstname", "primeiro"),                   ("data", "firstname")),
    (("cnpj",),                                                     ("brasil", "cnpj")),
    (("cpf",),                                                      ("brasil", "cpf")),
    (("email", "e-mail", "correo"),                                 ("data", "email")),
    (("cep", "postal", "zipcode"),                                  ("brasil", "cep")),
    (("celular", "telefone", "telephone", "whatsapp", "phone", "movil"), ("data", "phone")),
    (("chassi", "chasis", "chassis", "vin"),                        ("data", "vin")),
    (("veiculo", "quantidade", "cantidad", "numero de", "qtd", "amount"), ("gen", "amount")),
    (("modelo", "model"),                                           ("data", "model")),
    (("comentario", "mensagem", "mensaje", "observ", "comment"),    ("data", "comment")),
    (("cidade", "ciudad"),                                          ("data", "city")),
    (("concessionaria", "concesionario", "dealer", "loja"),         ("data", "dealer")),
)

_COMMENT_FALLBACKS = (
    "Interesse em frota corporativa.",
    "Gostaria de mais informacoes sobre venda direta.",
    "Aguardo contato do consultor.",
)


def is_aem_adaptive_form(driver):
    """True si la página actual es un Adobe AEM Adaptive Form (Guide)."""
    try:
        return bool(driver.execute_script(
            "return !!(document.querySelector('.guideFieldWidget') "
            "|| document.querySelector('[id$=\"___widget\"]'));"
        ))
    except Exception:
        return False


def _resolve_value(form_data, spec, is_brasil, gen_doc):
    kind, arg = spec
    if kind == "gen":
        if arg == "amount":
            return str(random.randint(2, 30))
        if arg == "company":
            razon = random.choice(("Transportes", "Comercio", "Servicos", "Logistica", "Industria"))
            return f"{razon} {random.randint(100, 999)} LTDA"
        return ""
    if kind == "brasil":
        # Valor del Excel por tipo (columna CPF/CNPJ/CEP); fallback a "document" legado.
        raw = str(form_data.get(arg, "") or form_data.get("document", "") or "").strip()
        digits = "".join(c for c in raw if c.isdigit())
        need = 14 if arg == "cnpj" else (8 if arg == "cep" else 11)
        if digits and len(digits) == need - 1:
            digits = digits.zfill(need)
        # No basta con el largo correcto: si el valor del Excel no tiene un dígito
        # verificador válido (CPF/CNPJ), el form real lo va a rechazar igual —
        # se regenera por API en vez de usarlo tal cual.
        checksum_ok = True
        if arg == "cpf":
            checksum_ok = _cpf_checksum_ok(digits) if digits else False
        elif arg == "cnpj":
            checksum_ok = _cnpj_checksum_ok(digits) if digits else False
        if is_brasil and (not digits or len(digits) < need or not checksum_ok):
            return gen_doc(arg)
        return digits if digits else (gen_doc(arg) if is_brasil else "")
    # kind == "data"
    val = str(form_data.get(arg, "") or "").strip()
    if val.endswith(".0"):
        val = val[:-2]
    if not val and arg == "comment":
        return random.choice(_COMMENT_FALLBACKS)
    return val


def _fill_text(driver, el, value):
    """Setea texto libre por native setter con eventos (input/change/blur)."""
    driver.execute_script(
        "var el=arguments[0],v=arguments[1];el.focus();"
        "try{var p=el.tagName==='TEXTAREA'?window.HTMLTextAreaElement.prototype:window.HTMLInputElement.prototype;"
        "Object.getOwnPropertyDescriptor(p,'value').set.call(el,v);}catch(e){el.value=v;}"
        "el.dispatchEvent(new Event('input',{bubbles:true}));"
        "el.dispatchEvent(new Event('change',{bubbles:true}));"
        "el.dispatchEvent(new Event('blur',{bubbles:true}));",
        el, value,
    )


def _fill_text_sendkeys(driver, el, value):
    """Texto libre vía send_keys real. Usado en Android: igual que en los forms de
    React, el value+dispatchEvent puede no registrar en el estado interno del form."""
    try:
        el.click()
    except Exception:
        try:
            driver.execute_script("arguments[0].click();", el)
        except Exception:
            pass
    try:
        el.clear()
    except Exception:
        try:
            driver.execute_script("arguments[0].value='';", el)
        except Exception:
            pass
    for ch in str(value):
        try:
            el.send_keys(ch)
        except Exception:
            pass
        time.sleep(0.005)
    try:
        driver.execute_script(
            "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));"
            "arguments[0].dispatchEvent(new Event('blur',{bubbles:true}));",
            el,
        )
    except Exception:
        pass


def _fill_masked(driver, el, digits):
    """Campos con máscara (tel/CPF/CNPJ/CEP): setear DÍGITOS CRUDOS sin focus() previo
    (el focus mueve el caret y la máscara sólo formatea hasta ahí) y sin maxlength
    (cuenta el valor formateado). La máscara del campo formatea al disparar input."""
    driver.execute_script(
        "var el=arguments[0],v=arguments[1];"
        "el.removeAttribute('maxlength');"
        "var set=Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype,'value').set;"
        "set.call(el,'');el.dispatchEvent(new Event('input',{bubbles:true}));"
        "set.call(el,v);"
        "el.dispatchEvent(new Event('input',{bubbles:true}));"
        "el.dispatchEvent(new Event('change',{bubbles:true}));"
        "el.dispatchEvent(new Event('blur',{bubbles:true}));",
        el, digits,
    )


def _fill_masked_sendkeys(driver, el, digits):
    """
    Campos con máscara (tel/CPF/CNPJ/CEP) vía send_keys real, dígito a dígito.
    Usado en Android: la librería de máscara puede no reaccionar a un value+dispatchEvent
    (sin keystrokes reales), igual que ya vimos en los forms de React — send_keys pasa
    por el pipeline real de eventos de teclado y la máscara formatea correctamente.
    Verifica al final que entraron todos los dígitos (secuencias largas como CNPJ a
    veces pierden alguno bajo latencia de LambdaTest) y reintenta si no coinciden.
    """
    for _attempt in range(3):
        try:
            driver.execute_script("arguments[0].removeAttribute('maxlength');", el)
        except Exception:
            pass
        try:
            el.click()
        except Exception:
            try:
                driver.execute_script("arguments[0].click();", el)
            except Exception:
                pass
        try:
            el.clear()
        except Exception:
            try:
                driver.execute_script("arguments[0].value='';", el)
            except Exception:
                pass
        for ch in str(digits):
            try:
                el.send_keys(ch)
                # La máscara reposiciona el cursor al reformatear (ej. tras insertar el
                # ")" o el "-"); si el siguiente dígito entra donde quedó el cursor y no
                # al final, el reformateo posterior corta la cola. End real lo corrige.
                el.send_keys(Keys.END)
            except Exception:
                pass
            time.sleep(0.015)
        time.sleep(0.15)
        try:
            # Solo 'change' acá — el 'blur' sintético entra en conflicto con el blur
            # REAL que dispara Selenium al mover el foco al siguiente campo (doble
            # blur), y algunas máscaras truncan el valor en el segundo. El blur real
            # de la transición al próximo campo alcanza para que la máscara valide.
            driver.execute_script(
                "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
                el,
            )
        except Exception:
            pass
        try:
            current_digits = "".join(c for c in (el.get_attribute("value") or "") if c.isdigit())
        except Exception:
            current_digits = ""
        if current_digits == str(digits):
            return
        time.sleep(0.2)


def _scroll_into_view(driver, el):
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        time.sleep(0.1)
    except Exception:
        pass


def _select_option_js(driver, sel, chosen):
    """Setea el <select> por el setter nativo del prototipo, no por asignación directa
    de .value. AEM Adaptive Forms trackea el value por componente (React-like); una
    asignación directa no pasa por ese tracker y el widget puede revertir la UI al
    placeholder ("Selecione"/"Selecionar") en el próximo re-render, aunque el 'change'
    se haya disparado. Mismo patrón ya usado para inputs de texto (native setter +
    dispatchEvent) en el resto del código."""
    _scroll_into_view(driver, sel)
    driver.execute_script(
        "var s=arguments[0],v=arguments[1];"
        "try{var d=Object.getOwnPropertyDescriptor(window.HTMLSelectElement.prototype,'value');"
        "if(d&&d.set){d.set.call(s,v);}else{s.value=v;}}catch(e){s.value=v;}"
        "s.dispatchEvent(new Event('input',{bubbles:true}));"
        "s.dispatchEvent(new Event('change',{bubbles:true}));"
        "s.dispatchEvent(new Event('blur',{bubbles:true}));",
        sel, chosen,
    )


def _is_persona_select(valid_options):
    texts = [normalize_text(t) for _, t in valid_options]
    return any("empresa" in x for x in texts) and any("pessoa" in x for x in texts)


def _fill_persona_select(driver, form_data, log):
    """
    Solo el selector Pessoa/Empresa: elige según el doc del Excel (CNPJ→'Empresa',
    CPF→'Pessoa', ninguno→'Empresa'). Debe correr ANTES que el resto del llenado
    porque determina qué campos (CPF vs CNPJ/empresa) quedan visibles.
    """
    has_cnpj = bool(str(form_data.get("cnpj", "") or "").strip())
    has_cpf = bool(str(form_data.get("cpf", "") or "").strip())
    prefer = "empresa" if has_cnpj else ("pessoa" if has_cpf else "empresa")
    try:
        selects = driver.find_elements(By.CSS_SELECTOR, "select[id*='widget']")
    except Exception:
        return 0
    for sel in selects:
        try:
            if not sel.is_displayed() or not sel.is_enabled():
                continue
            options = [(o.get_attribute("value") or "", (o.text or "").strip())
                       for o in sel.find_elements(By.TAG_NAME, "option")]
            valid = [(v, t) for v, t in options if v and not _is_placeholder(t)]
            if not valid or not _is_persona_select(valid):
                continue
            chosen = next((v for v, t in valid if prefer in normalize_text(t)), valid[0][0])
            _select_option_js(driver, sel, chosen)
            return 1
        except StaleElementReferenceException:
            continue
        except Exception as e:
            log(f" AEM select persona: {e}")
    return 0


# Claves de __by_id que corresponden a <select> mapeados en el Excel (Modelo, Fecha
# estimada, Región, Ciudad, Concesionario). Se usan para respetar el valor del Excel.
_SELECT_EXCEL_KEYS = (
    "models", "model", "estimated-date-purchase", "estimated-day",
    "region", "city", "dealer",
)


# Claves de __by_id (id del campo mapeado) que sí representan al Modelo.
_MODEL_KEYS = ("models", "model")


def _desired_select_values(form_data):
    """Pares (key, value) de select tipeados en el Excel (desde __by_id), no vacíos."""
    by_id = form_data.get("__by_id", {}) if isinstance(form_data, dict) else {}
    if not isinstance(by_id, dict):
        return []
    out = []
    for k in _SELECT_EXCEL_KEYS:
        raw = str(by_id.get(k, "") or "").strip()
        if raw.endswith(".0"):
            raw = raw[:-2]
        if raw:
            out.append((k, raw))
    return out


def _match_option_to_excel(valid, desired):
    """Opción de este select que coincide con algún valor del Excel.

    Devuelve (option_value, excel_key) o (None, None). Prioridad: match exacto
    normalizado, luego 'contiene' (tolerante, como LambdaTest).
    """
    if not desired:
        return None, None
    norm_opts = [(v, normalize_text(t)) for v, t in valid]
    for key, d in desired:
        nd = normalize_text(d)
        if not nd:
            continue
        for v, nt in norm_opts:
            if nt == nd:
                return v, key
    for key, d in desired:
        nd = normalize_text(d)
        if not nd:
            continue
        for v, nt in norm_opts:
            if nd in nt or nt in nd:
                return v, key
    return None, None


def _model_from_url(driver):
    """Modelo parametrizado en la URL del form (?model=...). '' si no está."""
    try:
        url = driver.current_url or ""
    except Exception:
        return ""
    try:
        query = urlsplit(url).query
        raw = parse_qs(query).get("model", [""])[0]
        return unquote(raw).replace("+", " ").strip()
    except Exception:
        return ""


def _fill_other_selects(driver, form_data, log, record=None):
    """
    El resto de los <select> (Modelo, Fecha estimada, etc.). Respeta PRIMERO el valor
    del Excel (match por texto de opción, ya que el id del widget AEM es genérico) y solo
    usa la 1ª opción válida cuando el Excel no trae ese valor o no coincide. Se llama
    DESPUÉS del llenado de texto (Nome/Sobrenome primero) para respetar el orden visual.

    Registra (record) cada select mapeado que llena, para que el Excel de resultado
    muestre el valor elegido (Modelo, Fecha estimada, etc.).
    """
    try:
        selects = driver.find_elements(By.CSS_SELECTOR, "select[id*='widget']")
    except Exception:
        return 0
    desired = _desired_select_values(form_data)
    model_recorded = False
    done = 0
    for sel in selects:
        try:
            if not sel.is_displayed() or not sel.is_enabled():
                continue
            options = [(o.get_attribute("value") or "", (o.text or "").strip())
                       for o in sel.find_elements(By.TAG_NAME, "option")]
            valid = [(v, t) for v, t in options if v and not _is_placeholder(t)]
            if not valid or _is_persona_select(valid):
                continue  # ya resuelto por _fill_persona_select
            current = (sel.get_attribute("value") or "").strip()
            if current:
                continue  # ya tiene un valor (ej. re-selección tras reload)
            # Excel primero: si alguna opción coincide con un valor tipeado, usarla.
            chosen, matched_key = _match_option_to_excel(valid, desired)
            if chosen is None:
                chosen = valid[0][0]  # sin valor en Excel → 1ª opción válida
            chosen_text = next((t for v, t in valid if v == chosen), chosen)
            _select_option_js(driver, sel, chosen)
            done += 1
            # Registrar el campo mapeado que coincidió (Modelo/Fecha/Región/etc.).
            if matched_key and record:
                try:
                    record(matched_key, chosen_text)
                except Exception:
                    pass
                if matched_key in _MODEL_KEYS:
                    model_recorded = True
                    log(f"🚗 AEM Modelo (del Excel) = {chosen_text}")
        except StaleElementReferenceException:
            continue
        except Exception as e:
            log(f" AEM select: {e}")
    # Si el Excel no traía Modelo, reportar el modelo parametrizado en la URL (?model=).
    if not model_recorded:
        url_model = _model_from_url(driver)
        if url_model and record:
            by_id = form_data.get("__by_id", {}) if isinstance(form_data, dict) else {}
            model_key = next((k for k in _MODEL_KEYS if k in by_id), "model")
            try:
                record(model_key, url_model)
            except Exception:
                pass
            log(f"🚗 AEM Modelo (de la URL ?model=) = {url_model}")
    return done


def fill_aem_form(driver, form_data, is_brasil, gen_doc, log=print, record=None, is_android=False):
    """Llena un AEM Adaptive Form. form_data: dict con firstname/lastname/email/phone/
    document/comment y cpf/cnpj/cep. gen_doc(kind) genera doc brasileño ('cpf'/'cnpj'/'cep').
    is_android=True: campos enmascarados (tel/cpf/cnpj/cep) vía send_keys real —
    en Android la máscara puede no reaccionar a value+dispatchEvent sin keystrokes reales."""
    _fill_persona_select(driver, form_data, log)   # CNPJ→Empresa, CPF→Pessoa (define qué campos se ven)
    # Android es más lento para terminar de renderizar el form tras el selector de
    # persona — si el snapshot de widgets se toma antes de tiempo, Nome/Sobrenome
    # (arriba del todo) pueden no estar aún en el DOM y quedar directamente afuera
    # (no es un problema de orden de llenado, es que ni se detectaron). Se espera un
    # poco más ahí y después el llenado en sí va rápido igual.
    if is_android:
        for _ in range(20):  # hasta ~3s, corta apenas aparecen widgets
            try:
                if driver.execute_script(
                    "return document.querySelectorAll(\"input[id*='widget'], textarea[id*='widget']\").length;"
                ):
                    break
            except Exception:
                pass
            time.sleep(0.15)
        time.sleep(0.5)
    else:
        time.sleep(0.8)

    # Un solo round-trip para reunir id/tipo/visible/label/posición de TODOS los widgets,
    # en vez de 5-6 llamadas remotas por campo (crítico en Android/LambdaTest: cada
    # round-trip pesa varios cientos de ms). Se ordena por posición visual (top, left)
    # para que se llene de arriba hacia abajo — document.querySelectorAll no garantiza
    # ese orden, sólo el orden del DOM, que en layouts AEM no siempre coincide.
    try:
        candidates = driver.execute_script(r"""
            var out = [];
            var els = document.querySelectorAll("input[id*='widget'], textarea[id*='widget']");
            for (var i=0; i<els.length; i++) {
                var el = els[i];
                var id = el.id || "";
                if (id.toLowerCase().indexOf("widget") < 0) continue;
                var type = (el.getAttribute("type") || "text").toLowerCase();
                var cs = getComputedStyle(el);
                var displayed = cs.display !== "none" && cs.visibility !== "hidden"
                    && (el.offsetWidth > 0 || el.offsetHeight > 0 || el.getClientRects().length > 0);
                var lbl = document.querySelector('label[for="' + id + '"]');
                var r = el.getBoundingClientRect();
                out.push({
                    id: id, type: type, displayed: displayed, enabled: !el.disabled,
                    label: lbl ? lbl.textContent : "", top: r.top, left: r.left,
                });
            }
            out.sort(function(a, b) { return (a.top - b.top) || (a.left - b.left); });
            return out;
        """) or []
    except Exception as e:
        log(f" AEM: error buscando widgets — {e}")
        return 0

    skip_types = ("hidden", "submit", "button", "checkbox", "radio", "file", "reset")
    filled = 0
    for cand in candidates:
        try:
            wid = cand.get("id") or ""
            if not cand.get("displayed") or not cand.get("enabled"):
                continue
            if (cand.get("type") or "text") in skip_types:
                continue

            lbl = cand.get("label") or ""
            hint = normalize_text(lbl) + " " + wid.lower()

            entry = next((e for e in _FIELD_KEYWORDS if any(k in hint for k in e[0])), None)
            if not entry:
                continue

            value = _resolve_value(form_data, entry[1], is_brasil, gen_doc)
            if not value:
                continue

            try:
                el = driver.find_element(By.ID, wid)
            except Exception:
                continue

            _scroll_into_view(driver, el)
            kind, arg = entry[1]
            logged_value = value
            if kind == "brasil" or arg == "phone":
                digits = "".join(c for c in value if c.isdigit())
                logged_value = digits  # lo que realmente se tipeó, sin puntuación
                if is_android:
                    _fill_masked_sendkeys(driver, el, digits)
                else:
                    _fill_masked(driver, el, digits)
            else:
                if is_android:
                    _fill_text_sendkeys(driver, el, value)
                else:
                    _fill_text(driver, el, value)
            log(f"✅ AEM '{(lbl or wid).strip()}' = {logged_value}")
            if record:
                try:
                    record(wid, logged_value)
                except Exception:
                    pass
            filled += 1
        except StaleElementReferenceException:
            continue
        except Exception as e:
            log(f" AEM: error llenando widget — {e}")

    log(f" AEM: {filled} input(s) llenado(s) por keyword")

    # Resto de selects (Modelo, etc.) al final — respeta el orden visual: texto
    # (Nome/Sobrenome/CPF/Celular/Email arriba) primero, Modelo después.
    _fill_other_selects(driver, form_data, log, record)

    return filled


def mark_aem_terms(driver, log=print):
    """Marca los checkboxes de términos AEM (guidecheckbox / required)."""
    marked = 0
    try:
        cbs = driver.find_elements(By.CSS_SELECTOR, "input[type='checkbox'][id*='guidecheckbox'], input[type='checkbox'][id*='widget']")
    except Exception:
        cbs = []
    for cb in cbs:
        try:
            if not cb.is_displayed():
                continue
            if cb.is_selected():
                continue
            ok = driver.execute_script(
                "var inp=arguments[0];var item=inp.closest('.guideCheckBoxItem');"
                "if(item){item.click();}else{inp.click();}"
                "return inp.checked===true||inp.getAttribute('aria-checked')==='true';", cb)
            if ok:
                marked += 1
        except StaleElementReferenceException:
            continue
        except Exception as e:
            log(f" AEM checkbox: {e}")
    if marked:
        log(f" AEM: {marked} checkbox(es) de términos marcado(s)")
    return marked
