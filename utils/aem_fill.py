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

from selenium.webdriver.common.by import By
from selenium.common.exceptions import StaleElementReferenceException


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
        if is_brasil and (not digits or len(digits) < need):
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


def _scroll_into_view(driver, el):
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        time.sleep(0.1)
    except Exception:
        pass


def _fill_selects(driver, form_data, log):
    """Selecciona los <select> AEM. En el selector persona/empresa elige según el doc del
    Excel: si viene CNPJ → 'Empresa'; si viene CPF → 'Pessoa'; si ninguno → 'Empresa' por
    defecto. El resto de selects: primera opción válida."""
    has_cnpj = bool(str(form_data.get("cnpj", "") or "").strip())
    has_cpf = bool(str(form_data.get("cpf", "") or "").strip())
    prefer = "empresa" if has_cnpj else ("pessoa" if has_cpf else "empresa")
    try:
        selects = driver.find_elements(By.CSS_SELECTOR, "select[id*='widget']")
    except Exception:
        return 0
    done = 0
    for sel in selects:
        try:
            if not sel.is_displayed() or not sel.is_enabled():
                continue
            options = [(o.get_attribute("value") or "", (o.text or "").strip())
                       for o in sel.find_elements(By.TAG_NAME, "option")]
            valid = [(v, t) for v, t in options if v and not _is_placeholder(t)]
            if not valid:
                continue
            texts = [normalize_text(t) for _, t in valid]
            is_persona = any("empresa" in x for x in texts) and any("pessoa" in x for x in texts)
            if is_persona:
                chosen = next((v for v, t in valid if prefer in normalize_text(t)), valid[0][0])
            else:
                chosen = valid[0][0]
            _scroll_into_view(driver, sel)
            driver.execute_script(
                "var s=arguments[0];s.value=arguments[1];"
                "s.dispatchEvent(new Event('input',{bubbles:true}));"
                "s.dispatchEvent(new Event('change',{bubbles:true}));"
                "s.dispatchEvent(new Event('blur',{bubbles:true}));",
                sel, chosen,
            )
            done += 1
        except StaleElementReferenceException:
            continue
        except Exception as e:
            log(f" AEM select: {e}")
    return done


def fill_aem_form(driver, form_data, is_brasil, gen_doc, log=print, record=None):
    """Llena un AEM Adaptive Form. form_data: dict con firstname/lastname/email/phone/
    document/comment y cpf/cnpj/cep. gen_doc(kind) genera doc brasileño ('cpf'/'cnpj'/'cep')."""
    _fill_selects(driver, form_data, log)   # persona según doc del Excel (CNPJ→Empresa, CPF→Pessoa)
    time.sleep(0.8)

    try:
        widgets = driver.find_elements(By.CSS_SELECTOR, "input[id*='widget'], textarea[id*='widget']")
    except Exception as e:
        log(f" AEM: error buscando widgets — {e}")
        return 0

    skip_types = ("hidden", "submit", "button", "checkbox", "radio", "file", "reset")
    filled = 0
    for el in widgets:
        try:
            if not el.is_displayed() or not el.is_enabled():
                continue
            wid = el.get_attribute("id") or ""
            if "widget" not in wid.lower():
                continue
            if (el.get_attribute("type") or "text").lower() in skip_types:
                continue

            lbl = ""
            try:
                lbl = driver.execute_script(
                    "var l=document.querySelector('label[for=\"'+arguments[0]+'\"]');"
                    "return l ? l.textContent : '';", wid) or ""
            except Exception:
                pass
            hint = normalize_text(lbl) + " " + wid.lower()

            entry = next((e for e in _FIELD_KEYWORDS if any(k in hint for k in e[0])), None)
            if not entry:
                continue

            value = _resolve_value(form_data, entry[1], is_brasil, gen_doc)
            if not value:
                continue

            _scroll_into_view(driver, el)
            kind, arg = entry[1]
            if kind == "brasil" or arg == "phone":
                _fill_masked(driver, el, "".join(c for c in value if c.isdigit()))
            else:
                _fill_text(driver, el, value)
            log(f"✅ AEM '{(lbl or wid).strip()}' = {value}")
            if record:
                try:
                    record(wid, value)
                except Exception:
                    pass
            filled += 1
        except StaleElementReferenceException:
            continue
        except Exception as e:
            log(f" AEM: error llenando widget — {e}")

    log(f" AEM: {filled} input(s) llenado(s) por keyword")
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
