"""BR Montana accesorios: "Saber mais" lleva a una categoría y ahí "Tenho interesse" al form.
Algunas categorías (ej. /pacotes) no tienen "Tenho interesse": se vuelve a la landing y se
prueba por otro "Saber mais" hasta encontrar un camino que llegue al form."""
from osocio.core.base_form_filler import BaseFormFiller

BASE = "https://www.chevrolet.com.br/picapes/chevrolet-montana/acessorios"


class _Elem:
    def __init__(self, href, title):
        self.href, self.title, self.text = href, title, title

    def is_displayed(self):
        return True

    def get_attribute(self, nombre):
        return {"href": self.href, "title": self.title}.get(nombre)


class _Driver:
    """Cada página tiene sus CTAs; hacer click en uno navega a su href."""

    def __init__(self, paginas):
        self.paginas, self.current_url, self.visitas = paginas, BASE, [BASE]

    def get(self, url):
        self.current_url = url
        self.visitas.append(url)

    def find_elements(self, by, selector):
        return [e for e in self.paginas.get(self.current_url, []) if f"title='{e.title}'" in selector]

    def execute_script(self, script, *args):
        if "readyState" in script:
            return "complete"
        if args and isinstance(args[0], _Elem):
            self.get(args[0].href)


def _form(paginas, monkeypatch):
    import osocio.core.base_form_filler as m
    monkeypatch.setattr(m.time, "sleep", lambda *_: None)
    monkeypatch.setattr(m.random, "choice", lambda xs: xs[0])   # siempre el primero disponible
    f = object.__new__(BaseFormFiller)
    f.driver = _Driver(paginas)
    f._landing_issue = ""
    f.handle_cookie_popups = lambda: None
    f.pre_scroll_for_dynamic_content = lambda: None
    return f


def _saber_mais(cat):
    return _Elem(f"{BASE}/{cat}", "Saber mais")


TENHO = _Elem(f"{BASE}/formulario", "Tenho interesse")


def test_si_la_categoria_no_tiene_tenho_interesse_prueba_otra(monkeypatch):
    f = _form({
        BASE: [_saber_mais("pacotes"), _saber_mais("aparencia")],
        f"{BASE}/pacotes": [],
        f"{BASE}/aparencia": [TENHO],
    }, monkeypatch)
    f._entrar_por_cta_de_landing(BASE)
    assert f._landing_issue == ""
    assert f.driver.current_url == f"{BASE}/formulario"
    assert f.driver.visitas.count(BASE) == 2   # volvió a la landing una vez


def test_si_ningun_camino_llega_lo_dice_con_los_que_probo(monkeypatch):
    f = _form({
        BASE: [_saber_mais("pacotes"), _saber_mais("rodas")],
        f"{BASE}/pacotes": [],
        f"{BASE}/rodas": [],
    }, monkeypatch)
    f._entrar_por_cta_de_landing(BASE)
    assert "Se probaron 2 caminos (pacotes, rodas)" in f._landing_issue
    assert "No se encontró ningún CTA" in f._landing_issue


def test_si_el_primer_camino_anda_no_prueba_otros(monkeypatch):
    f = _form({
        BASE: [_saber_mais("aparencia"), _saber_mais("pacotes")],
        f"{BASE}/aparencia": [TENHO],
    }, monkeypatch)
    f._entrar_por_cta_de_landing(BASE)
    assert f._landing_issue == ""
    assert f.driver.visitas == [BASE, f"{BASE}/aparencia", f"{BASE}/formulario"]


def test_landing_sin_ningun_saber_mais_no_reintenta(monkeypatch):
    f = _form({BASE: []}, monkeypatch)
    f._entrar_por_cta_de_landing(BASE)
    assert f._landing_issue.startswith("No se encontró ningún CTA")
    assert f.driver.visitas == [BASE]
