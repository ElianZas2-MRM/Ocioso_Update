"""El pre-scroll tiene que seguir a la pagina mientras crece.

Muchos formularios montan sus campos recien cuando entran en pantalla
(IntersectionObserver, lazy-loading). Por eso todo lo que abre un navegador baja por la
pagina antes de mirarla: si no, la mitad del form todavia no existe en el DOM.

El bug que estos tests fijan: la altura se media **una sola vez, antes de empezar a
bajar**. Pero al bajar, una pagina con carga diferida CRECE. Con la altura inicial, el
bucle cortaba temprano y se salteaba todo el medio - justo el caso para el que la funcion
existe. Un salto final al fondo lo disimulaba: llegaba abajo, pero sin haber pasado por
las secciones intermedias, que quedaban sin cargar.

Estaba escrito cuatro veces, una por cada cosa que abre un navegador, y las cuatro copias
tenian el mismo bug. Ahora hay una sola implementacion y estos tests la prueban por los
cuatro accesos, asi que no puede volver a arreglarse en una y quedar roto en las otras.
"""
import re

import pytest


@pytest.fixture(autouse=True)
def sin_esperas(monkeypatch):
    """El pre-scroll duerme entre pasos para darle tiempo al navegador.

    Con hasta 60 pasos eso son decenas de segundos por test, y una suite lenta se corre
    menos seguido. Aca no hay navegador de verdad esperando nada, asi que las esperas se
    anulan: lo que se prueba es el RECORRIDO, no los tiempos.
    """
    import time

    monkeypatch.setattr(time, "sleep", lambda _s: None)


# Valor con el que se pide "hasta el fondo", sin saber cuanto mide.
FONDO = 999999


def _pedido(script, args, alto_actual):
    """Devuelve (posicion, es_salto_al_fondo) de un window.scrollTo.

    Las cuatro copias pedian el scroll distinto: unas pasan la posicion como
    arguments[0] y otras la incrustan en el texto del script; el salto al fondo a veces
    es 999999 y a veces document.body.parentNode.scrollHeight. El doble entiende las dos
    formas para poder probar las cuatro con los mismos tests.

    Distinguir el salto al fondo importa: es justamente lo que disimulaba el bug, asi que
    si contara como una posicion intermedia los tests no verian nada.
    """
    if args:
        pedida = args[0]
    else:
        objetivo = script.split("scrollTo(0,", 1)[1]
        if "scrollHeight" in objetivo:
            return alto_actual, True
        pedida = float(re.search(r"-?[\d.]+", objetivo).group())
    return (alto_actual, True) if pedida >= FONDO else (pedida, False)


class DriverQueCrece:
    """Un doble de Selenium que simula lazy-loading: cada scroll estira la pagina.

    Es lo que hace un formulario largo de React: arranca con lo que entra en pantalla y va
    montando secciones nuevas a medida que se baja.
    """

    def __init__(self, alto_inicial=2000, viewport=800, crece=1500, alto_maximo=12000):
        self.alto = alto_inicial
        self.viewport = viewport
        self.crece = crece
        self.alto_maximo = alto_maximo
        self.posiciones = []    # todo lo pedido, con el fondo normalizado a FONDO
        self.intermedias = []   # solo las paradas del recorrido, sin el fondo ni el tope
        self.eventos = 0        # cuantas veces se le aviso a la pagina que hubo scroll

    def execute_script(self, script, *args):
        if "scrollTo" in script:
            if "dispatchEvent" in script:
                self.eventos += 1
            posicion, es_fondo = _pedido(script, args, self.alto)
            self.posiciones.append(FONDO if es_fondo else posicion)
            if not es_fondo and posicion > 0:
                self.intermedias.append(posicion)
            # Bajar hace que cargue contenido nuevo, hasta un tope.
            if posicion > 0 and self.alto < self.alto_maximo:
                self.alto = min(self.alto + self.crece, self.alto_maximo)
            return None
        if "innerHeight" in script:
            return self.viewport
        if "scrollHeight" in script:
            return self.alto
        return None


class DriverEstatico(DriverQueCrece):
    """Una pagina normal, que no crece: el pre-scroll tiene que terminar igual."""

    def __init__(self, alto=3000, viewport=800):
        super().__init__(alto_inicial=alto, viewport=viewport, crece=0, alto_maximo=alto)


def _motor(driver, headless=False):
    """Un BaseFormFiller sin __init__, con lo justo para correr el pre-scroll."""
    from osocio.core.base_form_filler import BaseFormFiller

    m = object.__new__(BaseFormFiller)
    m.driver = driver
    m.config = {"headless": headless}
    return m


def _implementaciones():
    """Los cuatro accesos al pre-scroll, uno por cada cosa que abre un navegador."""
    from osocio.core.dealer_comparator_runner import _pre_scroll_for_dynamic_content as dealers
    from osocio.providers.lambdatest_mac.lt_runner import _pre_scroll_for_dynamic_content as lt
    from osocio.validation.selenium_validation_runner import _pre_scroll_for_dynamic_content as val

    return {
        "envio_de_leads": lambda d: _motor(d).pre_scroll_for_dynamic_content(),
        "comparar_dealers": dealers,
        "validacion_de_campos": val,
        "lambdatest": lt,
    }


CUALES = ["comparar_dealers", "envio_de_leads", "lambdatest", "validacion_de_campos"]


def _correr(cual, driver):
    implementaciones = _implementaciones()
    assert set(implementaciones) == set(CUALES), (
        "aparecio o desaparecio un acceso al pre-scroll: hay que actualizar CUALES"
    )
    implementaciones[cual](driver)
    return driver


# --- lo que motiva todo esto -------------------------------------------------------

@pytest.mark.parametrize("cual", CUALES)
def test_sigue_a_la_pagina_mientras_crece(cual):
    """El caso del formulario de Cadillac: largo y con carga diferida."""
    driver = _correr(cual, DriverQueCrece(alto_inicial=2000, alto_maximo=12000))

    assert driver.intermedias, f"{cual}: no scrolleo a ninguna posicion intermedia"
    assert max(driver.intermedias) >= 9000, (
        f"{cual}: solo bajo hasta {max(driver.intermedias)} de {driver.alto_maximo}: se "
        "salteo el medio, que es justo donde estaba el contenido sin cargar"
    )


@pytest.mark.parametrize("cual", CUALES)
def test_no_se_queda_con_la_altura_inicial(cual):
    """La regresion concreta: con la altura vieja cortaba a los ~2000px."""
    driver = _correr(cual, DriverQueCrece(alto_inicial=2000, alto_maximo=12000))

    assert max(driver.intermedias) > 2000, (
        f"{cual}: corto en la altura que tenia la pagina antes de empezar a bajar"
    )


@pytest.mark.parametrize("cual", CUALES)
def test_le_avisa_a_la_pagina_que_hubo_scroll(cual):
    """Sin el evento, un IntersectionObserver no se entera y no monta nada.

    Mover window.scrollTo no alcanza en todos los frameworks. Dos de las cuatro copias
    solo movian la ventana, sin disparar el evento que espera el codigo de la pagina.
    """
    driver = _correr(cual, DriverQueCrece())

    assert driver.eventos > 0, f"{cual}: movio la ventana sin avisarle a la pagina"


# --- lo que no se tiene que romper al arreglar lo de arriba -------------------------

@pytest.mark.parametrize("cual", CUALES)
def test_termina_en_una_pagina_que_no_crece(cual):
    driver = _correr(cual, DriverEstatico(alto=3000))

    assert driver.intermedias, f"{cual}: no scrolleo nada"
    assert max(driver.intermedias) < 3000 + 800, f"{cual}: se fue mas alla del fondo real"


@pytest.mark.parametrize("cual", CUALES)
def test_no_se_cuelga_con_scroll_infinito(cual):
    """Una pagina que crece siempre no puede dejar la corrida colgada para siempre."""
    from osocio.utils.scroll_dinamico import MAX_PASOS

    driver = _correr(cual, DriverQueCrece(alto_inicial=2000, crece=5000, alto_maximo=10 ** 9))

    assert len(driver.intermedias) <= MAX_PASOS, (
        f"{cual}: paso el techo de pasadas; con scroll infinito no terminaria nunca"
    )


@pytest.mark.parametrize("cual", CUALES)
def test_vuelve_al_tope_al_terminar(cual):
    """El llenado arranca desde arriba: si queda en el fondo, el primer campo no se ve."""
    driver = _correr(cual, DriverQueCrece())

    assert driver.posiciones[-1] == 0, f"{cual}: tiene que dejar la pagina arriba"


@pytest.mark.parametrize("cual", CUALES)
def test_pasa_por_el_fondo_antes_de_volver(cual):
    """El salto al fondo dispara lo que quede colgando al final del form."""
    driver = _correr(cual, DriverQueCrece())

    assert FONDO in driver.posiciones, f"{cual}: nunca fue hasta el fondo"
    assert driver.posiciones.index(FONDO) < len(driver.posiciones) - 1


# --- el motor, que es el unico con modo headless -----------------------------------

@pytest.mark.parametrize("headless", [True, False])
def test_el_motor_funciona_igual_con_y_sin_ventana(headless):
    """En headless se espera mas entre pasos, pero el recorrido es el mismo."""
    driver = DriverQueCrece(alto_inicial=2000, alto_maximo=8000)
    _motor(driver, headless=headless).pre_scroll_for_dynamic_content()

    assert max(driver.intermedias) >= 6000
