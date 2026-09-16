"""`lt_runner` se está partiendo en módulos: esta es la red que lo vigila.

El archivo tenía 4.962 líneas con 94 funciones de módulo. A diferencia de `BaseFormFiller`
—que es una clase, y al partirla en mixins la herencia mantiene todo junto sola— acá son
funciones sueltas: moverlas de archivo sí puede romper quien las importa.

Y hay dos módulos que importan nombres puntuales de acá:

    lt_android_runner.py  -> 11 nombres
    lt_controller.py      ->  3 nombres

Por eso `lt_runner` queda como fachada: reexporta todo, y nadie de afuera se entera de que
el código se mudó. Estos tests verifican exactamente eso, porque el riesgo no es que falle
ruidosamente sino que un nombre desaparezca y recién se note al correr LambdaTest.
"""
import inspect

import pytest

from osocio.providers.lambdatest_mac import lt_runner


# Cantidad de funciones publicas+privadas que el modulo exponia antes de partirlo.
FUNCIONES_ESPERADAS_MINIMO = 94


def _funciones_del_modulo():
    return [n for n, v in vars(lt_runner).items()
            if callable(v) and not n.startswith("__")]


def test_no_se_perdio_ninguna_funcion_al_partir_el_modulo():
    funciones = _funciones_del_modulo()
    assert len(funciones) >= FUNCIONES_ESPERADAS_MINIMO, (
        f"lt_runner expone {len(funciones)} invocables y antes de partirlo tenía al menos "
        f"{FUNCIONES_ESPERADAS_MINIMO}. Alguna extracción se comió funciones."
    )


# --- los nombres que OTROS modulos importan de aca --------------------------------

NOMBRES_QUE_USA_ANDROID = [
    "load_credentials", "_safe_log", "mark_lt_status", "LT_HUB", "_fetch_lt_video_url",
    "_load_field_dependencies", "_load_ids_dinamicos", "_get_field_mapping_for_pais",
    "fill_form_fields", "_run_single_lead", "_write_row_result",
]

NOMBRES_QUE_USA_EL_CONTROLLER = ["run_lt_batch", "LTRunOptions", "load_credentials"]


@pytest.mark.parametrize("nombre", NOMBRES_QUE_USA_ANDROID)
def test_lt_android_runner_puede_seguir_importando(nombre):
    """Android reutiliza casi toda la lógica de Mac. Si un nombre se va, Android muere."""
    assert hasattr(lt_runner, nombre), f"lt_runner ya no expone {nombre}"


@pytest.mark.parametrize("nombre", NOMBRES_QUE_USA_EL_CONTROLLER)
def test_el_controller_puede_seguir_importando(nombre):
    assert hasattr(lt_runner, nombre), f"lt_runner ya no expone {nombre}"


def test_los_importadores_reales_siguen_funcionando():
    """La prueba de fuego: que los módulos que dependen de esto importen de verdad."""
    import importlib

    importlib.import_module("osocio.providers.lambdatest_android.lt_android_runner")
    importlib.import_module("osocio.providers.lambdatest_mac.lt_controller")


# --- las rutas de resultados, que ya se tocaron una vez ----------------------------

def test_las_carpetas_de_resultados_salen_de_paths():
    from osocio import paths
    import os

    assert os.path.normpath(lt_runner._RESULTADOS_DIR) == os.path.normpath(
        paths.RESULTS_LT_MAC_DIR
    )


def test_el_punto_de_entrada_tiene_la_firma_esperada():
    """`run_lt_batch` es por donde entra todo: su firma es contrato con el controller."""
    params = inspect.signature(lt_runner.run_lt_batch).parameters
    assert "opts" in params, "run_lt_batch dejó de recibir opts"
