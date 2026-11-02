"""El generador de Excels elige el mercado por las URLs pegadas y frena si se mezclan."""
import pytest

from osocio.utils.mercado_por_url import (
    aviso_pares_inconsistentes, aviso_varios_mercados, detectar_mercado, mercados_de_url,
    pares_inconsistentes,
)


@pytest.mark.parametrize("url,mercado", [
    # Por dominio
    ("https://www.chevrolet.com.br/solicitar-contato", "Brasil"),
    ("https://www.chevrolet.cl/camiones", "Chile"),
    ("https://www.chevrolet.com.co/vans/n400-cargo", "Colombia"),
    ("https://www.chevrolet.com.pe/colision", "Peru"),
    ("https://www.chevrolet.com.py/suvs/sonic", "Paraguay"),
    ("https://www.chevrolet.com.ar/x", "Argentina"),
    ("https://www.chevrolet.com.uy/x", "Uruguay"),
    ("https://www.chevrolet.com.ec/x", "Ecuador"),
    ("https://www.chevrolet.com.bo/x", "Bolivia"),
    ("www.chevrolet.cl/autos", "Chile"),                     # sin https
    # Por ruta (forms sueltos)
    ("https://secure-developments.com/commonwealth/chile/gm_forms/vehicle-shopper", "Chile"),
    ("https://stage.secure-developments.com/commonwealth/gm_frontend/cadillac/t1/brasil/form/lyriq",
     "Brasil"),
    # Dominio de preview de AEM (con guiones)
    ("https://www-cadillacbrasil-com-br-preview.prod.acp.gm.com/content/cadillac/sa/br/pt/lyriq.html",
     "Brasil"),
])
def test_mercado_de_cada_url(url, mercado):
    assert mercados_de_url(url) == {mercado}


@pytest.mark.parametrize("url", [
    "https://oss.gm.com/oss/o/b/CO/CH/?origin=CHEVY_CO",      # no dice el mercado
    "https://www.ejemplo.com/landing-de-prueba-1",
    "https://www.chevrolet.com/peruano",                      # no es el segmento "peru"
    "https://www.clima.com/x",                                 # .cl solo al final del dominio
    "",
])
def test_sin_mercado_no_se_adivina(url):
    assert mercados_de_url(url) == set()


def test_un_solo_mercado_se_elige_solo():
    mercado, por = detectar_mercado([
        "https://www.chevrolet.cl/camiones",
        "https://secure-developments.com/commonwealth/chile/gm_forms/camiones-raq",
        "https://www.ejemplo.com/sin-pais",                   # no cuenta para ninguno
    ])
    assert mercado == "Chile" and list(por) == ["Chile"] and len(por["Chile"]) == 2


def test_varios_mercados_no_elige_y_explica_cuales():
    mercado, por = detectar_mercado([
        "https://www.chevrolet.cl/camiones",
        "https://www.chevrolet.com.br/solicitar-contato",
        "https://www.chevrolet.com.br/form-interesse",
    ])
    assert mercado is None and sorted(por) == ["Brasil", "Chile"]
    aviso = aviso_varios_mercados(por)
    assert "2 mercados distintos" in aviso
    assert "Brasil (2):" in aviso and "Chile (1):" in aviso
    assert "https://www.chevrolet.cl/camiones" in aviso


def test_el_aviso_recorta_las_listas_largas():
    por = {"Brasil": [f"https://www.chevrolet.com.br/{i}" for i in range(5)],
           "Chile": ["https://www.chevrolet.cl/x"]}
    assert "… y 2 más" in aviso_varios_mercados(por)


def test_sin_urls_con_mercado():
    assert detectar_mercado(["https://www.ejemplo.com/a"]) == (None, {})
    assert detectar_mercado([]) == (None, {})


def test_landing_de_un_mercado_con_form_de_otro():
    preview_br = "https://www-cadillacbrasil-com-br-preview.prod.acp.gm.com/content/cadillac/sa/br/pt/lyriq.html"
    form_cl = "https://secure-developments.com/commonwealth/chile/gm_forms/vehicle-shopper"
    form_br = "https://stage.secure-developments.com/commonwealth/gm_frontend/cadillac/t1/brasil/form/lyriq"
    pares = [(preview_br, form_br), (preview_br, form_cl),
             ("https://www.ejemplo.com/preview", form_cl)]    # landing sin país: no se compara
    malos = pares_inconsistentes(pares)
    assert malos == [(2, preview_br, "Brasil", form_cl, "Chile")]
    aviso = aviso_pares_inconsistentes(malos)
    assert "Par 2:" in aviso and "Landing (Brasil)" in aviso and "Form (Chile)" in aviso
