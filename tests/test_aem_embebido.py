"""Forms AEM que el Excel pide por su URL de iframe (…/formcontainer.iframe.br) pero que la
landing muestra embebidos, sin iframe (BR 248 solicitar-proposta): se validan por la ruta."""
from osocio.core.base_form_filler import BaseFormFiller as B
from osocio.core.form_filler.aem import FormulariosAEMMixin, ruta_aem_de_url, rutas_aem_coinciden

ESPERADA = ("https://www.chevrolet.com.br/request-proposal/jcr:content/root/container/"
            "adv_grid_builder/form/formcontainer.iframe.br?dataRef=&wcmmode=DISABLED")
EN_PAGINA = ("/content/chevrolet/sa/br/pt/index/request-proposal/jcr:content/root/container/"
             "adv_grid_builder/form/formcontainer")


def test_ruta_sin_dominio_parametros_ni_selectores():
    assert ruta_aem_de_url(ESPERADA) == (
        "/request-proposal/jcr:content/root/container/adv_grid_builder/form/formcontainer")
    assert ruta_aem_de_url(EN_PAGINA) == EN_PAGINA
    assert ruta_aem_de_url("https://www.chevrolet.com.br/gm_forms/raq") == ""
    assert ruta_aem_de_url("") == ""


def test_coincide_por_el_final_de_la_ruta():
    assert rutas_aem_coinciden(EN_PAGINA, ESPERADA)
    assert rutas_aem_coinciden(ESPERADA, ESPERADA)
    # El mismo componente en otra página no es el form pedido.
    otra = EN_PAGINA.replace("request-proposal", "test-drive")
    assert not rutas_aem_coinciden(otra, ESPERADA)
    # Ni un nombre de página que sólo termina igual ("…-request-proposal").
    assert not rutas_aem_coinciden(EN_PAGINA.replace("/request-proposal", "/old-request-proposal"),
                                   ESPERADA)
    assert not rutas_aem_coinciden("", ESPERADA)


def test_form_coincide_acepta_la_ruta_encontrada():
    assert B._diferencia_url_form(EN_PAGINA, ESPERADA) == (True, [])
    assert B._mismas_urls_de_form(EN_PAGINA, ESPERADA)


class _Pagina(FormulariosAEMMixin):
    def __init__(self, rutas):
        self.driver = self
        self._rutas = rutas

    def execute_script(self, js, *args):
        return self._rutas


def test_busca_el_componente_en_la_pagina():
    otra = EN_PAGINA.replace("request-proposal", "test-drive")
    assert _Pagina([otra, EN_PAGINA])._aem_form_embebido(ESPERADA) == EN_PAGINA
    assert _Pagina([otra])._aem_form_embebido(ESPERADA) == ""
    # Una URL que no es de AEM no se busca por ruta.
    assert _Pagina([EN_PAGINA])._aem_form_embebido("https://x.com/gm_forms/raq") == ""
