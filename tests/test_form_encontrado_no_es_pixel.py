"""BR eletrico/solicitar-contato en Firefox: "Form URL encontrada" quedó con un pixel
Floodlight de doubleclick. El pixel lleva la URL del form como parámetro (ref=…gm_forms…)
y se lo tomaba por el form GM; además, tras el error de event_id se buscaba el form con el
driver todavía adentro del iframe del form, donde está ese pixel."""
from osocio.core.base_form_filler import BaseFormFiller

PIXEL = ("https://5257814.fls.doubleclick.net/activityi;src=5257814;type=reser0;cat=dcm_g00;"
         "u31=ch%3Asa%3Abr%3Apt%3At1%3Aelectric%3Arequest-contact;frm=2;epver=2;"
         "ref=https%3A%2F%2Fsecure-developments.com%2Fcommonwealth%2Fbrasil%2Fgm_forms%2Fraq-eletricos;"
         "_dc_test=1;~oref=https%3A%2F%2Fsecure-developments.com%2Fcommonwealth%2Fbrasil%2Fgm_forms"
         "%2Fraq-eletricos?")
FORM = "https://secure-developments.com/commonwealth/brasil/gm_forms/raq-eletricos"


def _f():
    return object.__new__(BaseFormFiller)


def test_pixel_de_doubleclick_no_es_form_gm():
    assert not _f()._is_gm_form_src(PIXEL)


def test_urls_gm_de_siempre_siguen_siendo_form():
    f = _f()
    for u in (FORM,
              "https://secure-developments.com/commonwealth/peru/gm_front/form/raq?model=onix",
              "https://secure-developments.com/commonwealth/chile/gm_admin/form/x",
              "https://secure-developments.com/gm_frontend/chevrolet/t3/argentina/form/raq",
              "https://secure-developments.com/commonwealth/brasil/gm_formns/raq"):
        assert f._is_gm_form_src(u), u
    assert not f._is_gm_form_src("")
    assert not f._is_gm_form_src(None)


def test_pick_elige_el_form_aunque_el_pixel_venga_primero():
    f = _f()
    pixel, form = object(), object()
    f._src_y_visibilidad_de_iframes = lambda fr: [(PIXEL, True), (FORM, True)]
    elegido, es_gm = f._pick_gm_iframe([pixel, form])
    assert elegido is form and es_gm


def test_buscar_el_form_arranca_desde_la_landing():
    llamadas = []

    class _SwitchTo:
        def default_content(self):
            llamadas.append("default_content")

    class _Driver:
        switch_to = _SwitchTo()

        def find_elements(self, *a):
            llamadas.append("find_elements")
            return []

        def execute_script(self, *a):
            return None

    f = _f()
    f.driver = _Driver()
    f.pre_scroll_for_dynamic_content = lambda: None
    import osocio.core.base_form_filler as m
    orig_wait = m.WebDriverWait

    class _SinEspera:
        def __init__(self, d, t):
            self.d = d

        def until(self, fn):
            fn(self.d)
            raise m.TimeoutException()

    m.WebDriverWait = _SinEspera
    orig_sleep = m.time.sleep
    m.time.sleep = lambda *_: None
    try:
        assert f.find_and_position_to_form(FORM) is None
    finally:
        m.WebDriverWait = orig_wait
        m.time.sleep = orig_sleep
    assert llamadas[0] == "default_content"
