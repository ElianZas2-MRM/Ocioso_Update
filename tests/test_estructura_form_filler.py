"""BaseFormFiller se está partiendo en mixins: esta es la red que lo vigila.

La clase tenía 156 métodos en 7.021 líneas. Se van moviendo de a grupos a
`osocio/core/form_filler/`, y cada grupo es un mixin del que BaseFormFiller hereda.

El riesgo de esa operación no es que falle ruidosamente — es que un método se pierda en
el camino y nadie se entere hasta que alguien lo llama en una corrida real. Estos tests
son justamente para eso: verifican que la superficie de la clase siga completa después de
cada extracción.
"""
import inspect

import pytest

from osocio.core.base_form_filler import BaseFormFiller


# BaseFormFiller tenia 156 metodos definidos en la clase antes de empezar a partirla,
# contando __init__. Como abajo se filtran los dunder, el piso es 155. Si una extraccion
# pierde metodos por el camino, este numero baja y el test canta.
METODOS_ESPERADOS_MINIMO = 155


def _metodos_de_la_clase():
    return [n for n, v in inspect.getmembers(BaseFormFiller, callable)
            if not n.startswith("__")]


def test_no_se_perdio_ningun_metodo_al_partir_la_clase():
    metodos = _metodos_de_la_clase()
    assert len(metodos) >= METODOS_ESPERADOS_MINIMO, (
        f"BaseFormFiller expone {len(metodos)} métodos y antes de partirla tenía al menos "
        f"{METODOS_ESPERADOS_MINIMO}. Alguna extracción se comió métodos."
    )


@pytest.mark.parametrize("metodo", [
    # El camino principal de una corrida
    "run",
    "initialize_browser",
    "process_landing_page",
    "submit_and_verify_form",
    "setup_directories_and_files",
    # Llenado y detección
    "fill_form_fields_auto_step",
    "extract_form_data",
    "safe_select_option_if_visible",
    # Navegación y DOM (primer grupo extraído a form_filler/navegacion.py)
    "handle_cookie_popups",
    "handle_gm_cookie_popup",
    "reposition_to_form",
    "pre_scroll_for_dynamic_content",
])
def test_los_metodos_del_camino_principal_siguen_accesibles(metodo):
    """Da igual en qué archivo viva cada uno: se tiene que poder llamar desde la clase."""
    assert hasattr(BaseFormFiller, metodo), f"BaseFormFiller perdió {metodo}()"
    assert callable(getattr(BaseFormFiller, metodo))


def test_el_import_de_siempre_sigue_funcionando():
    """Nada de afuera puede enterarse de que la clase se partió."""
    from osocio.core.base_form_filler import BaseFormFiller as Importada
    assert Importada is BaseFormFiller


def test_los_mixins_estan_en_la_cadena_de_herencia():
    nombres = [c.__name__ for c in BaseFormFiller.__mro__]
    assert "BaseFormFiller" == nombres[0]
    assert any(n.endswith("Mixin") for n in nombres), (
        "no hay ningún mixin en el MRO: la extracción no se aplicó"
    )


def test_ningun_mixin_pisa_metodos_de_otro():
    """Si dos mixins definen el mismo método, el MRO elige uno y el otro desaparece
    en silencio. Es el modo de falla clásico de la herencia múltiple."""
    vistos = {}
    duplicados = []
    for clase in BaseFormFiller.__mro__:
        if clase is object or clase is BaseFormFiller:
            continue
        for nombre, valor in vars(clase).items():
            if nombre.startswith("__") or not callable(valor):
                continue
            if nombre in vistos:
                duplicados.append(f"{nombre}: {vistos[nombre]} y {clase.__name__}")
            vistos[nombre] = clase.__name__

    assert not duplicados, "métodos definidos en más de un mixin:\n  " + "\n  ".join(duplicados)
