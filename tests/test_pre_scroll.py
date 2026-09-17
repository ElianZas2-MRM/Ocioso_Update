"""El pre-scroll tiene que seguir a la página mientras crece.

Muchos formularios montan sus campos recién cuando entran en pantalla
(IntersectionObserver, lazy-loading). Por eso el motor baja por la página antes de llenar:
si no, la mitad del form todavía no existe en el DOM.

El bug que estos tests fijan: la altura se medía **una sola vez, antes de empezar a
bajar**. Pero al bajar, una página con carga diferida CRECE. Con la altura inicial, el
bucle cortaba temprano y se salteaba todo el medio — justo el caso para el que la función
existe. Un salto final al fondo lo disimulaba: llegaba abajo, pero sin haber pasado por
las secciones intermedias, que quedaban sin cargar.
"""
import pytest


@pytest.fixture(autouse=True)
def sin_esperas(monkeypatch):
    """El pre-scroll duerme entre pasos para darle tiempo al navegador.

    Con hasta 60 pasos eso son decenas de segundos por test, y una suite lenta se corre
    menos seguido. Acá no hay navegador de verdad esperando nada, asi que las esperas se
    anulan: lo que se prueba es el RECORRIDO, no los tiempos.
    """
    import time

    monkeypatch.setattr(time, "sleep", lambda _s: None)


class DriverQueCrece:
    """Un doble de Selenium que simula lazy-loading: cada scroll estira la página.

    Es lo que hace un formulario largo de React: arranca con lo que entra en pantalla y va
    montando secciones nuevas a medida que se baja.
    """

    def __init__(self, alto_inicial=2000, viewport=800, crece=1500, alto_maximo=12000):
        self.alto = alto_inicial
        self.viewport = viewport
        self.crece = crece
        self.alto_maximo = alto_maximo
        self.posiciones = []

    def execute_script(self, script, *args):
        if "innerHeight" in script:
            return self.viewport
        if "scrollHeight" in script:
            return self.alto
        if "scrollTo" in script:
            posicion = args[0] if args else 0
            self.posiciones.append(posicion)
            # Bajar hace que cargue contenido nuevo, hasta un tope.
            if posicion > 0 and self.alto < self.alto_maximo:
                self.alto = min(self.alto + self.crece, self.alto_maximo)
            return None
        return None


class DriverEstatico:
    """Una página normal, que no crece: el pre-scroll tiene que terminar igual."""

    def __init__(self, alto=3000, viewport=800):
        self.alto = alto
        self.viewport = viewport
        self.posiciones = []

    def execute_script(self, script, *args):
        if "innerHeight" in script:
            return self.viewport
        if "scrollHeight" in script:
            return self.alto
        if "scrollTo" in script:
            self.posiciones.append(args[0] if args else 0)
        return None


def _motor(driver):
    """Un BaseFormFiller sin __init__, con lo justo para correr el pre-scroll."""
    from osocio.core.base_form_filler import BaseFormFiller

    m = object.__new__(BaseFormFiller)
    m.driver = driver
    m.config = {"headless": False}
    return m


def test_recorre_toda_la_pagina_aunque_crezca_mientras_baja():
    """El caso del formulario de Cadillac: largo y con carga diferida."""
    driver = DriverQueCrece(alto_inicial=2000, alto_maximo=12000)
    _motor(driver).pre_scroll_for_dynamic_content()

    # Sin contar el salto final al fondo ni el regreso al tope.
    intermedias = [p for p in driver.posiciones if 0 < p < 999999]
    assert intermedias, "no scrolleó a ninguna posición intermedia"
    assert max(intermedias) >= 9000, (
        f"solo bajó hasta {max(intermedias)} de {driver.alto_maximo}: se salteó el medio, "
        "que es justo donde estaba el contenido sin cargar"
    )


def test_no_se_queda_con_la_altura_inicial():
    """La regresión concreta: con la altura vieja cortaba a los ~2000px."""
    driver = DriverQueCrece(alto_inicial=2000, alto_maximo=12000)
    _motor(driver).pre_scroll_for_dynamic_content()

    intermedias = [p for p in driver.posiciones if 0 < p < 999999]
    assert max(intermedias) > 2000, (
        "cortó en la altura que tenía la página antes de empezar a bajar"
    )


def test_termina_en_una_pagina_que_no_crece():
    driver = DriverEstatico(alto=3000)
    _motor(driver).pre_scroll_for_dynamic_content()

    intermedias = [p for p in driver.posiciones if 0 < p < 999999]
    assert intermedias
    assert max(intermedias) < 3000 + 800, "se fue mucho más allá del fondo real"


def test_no_se_cuelga_con_scroll_infinito():
    """Una página que crece siempre no puede dejar la corrida colgada para siempre."""
    from osocio.core.base_form_filler import BaseFormFiller

    driver = DriverQueCrece(alto_inicial=2000, crece=5000, alto_maximo=10 ** 9)
    _motor(driver).pre_scroll_for_dynamic_content()

    scrolls = [p for p in driver.posiciones if 0 < p < 999999]
    assert len(scrolls) <= BaseFormFiller._MAX_PASOS_PRE_SCROLL, (
        "el bucle pasó del techo de pasadas: con scroll infinito no terminaría nunca"
    )


def test_vuelve_al_tope_al_terminar():
    """El llenado arranca desde arriba: si queda en el fondo, el primer campo no se ve."""
    driver = DriverQueCrece()
    _motor(driver).pre_scroll_for_dynamic_content()

    assert driver.posiciones[-1] == 0, "tiene que dejar la página arriba"


def test_pasa_por_el_fondo_antes_de_volver():
    """El salto al fondo dispara lo que quede colgando al final del form."""
    driver = DriverQueCrece()
    _motor(driver).pre_scroll_for_dynamic_content()

    assert 999999 in driver.posiciones


@pytest.mark.parametrize("headless", [True, False])
def test_funciona_igual_con_y_sin_ventana(headless):
    """En headless se espera más entre pasos, pero el recorrido es el mismo."""
    driver = DriverQueCrece(alto_inicial=2000, alto_maximo=8000)
    motor = _motor(driver)
    motor.config = {"headless": headless}
    motor.pre_scroll_for_dynamic_content()

    intermedias = [p for p in driver.posiciones if 0 < p < 999999]
    assert max(intermedias) >= 6000
