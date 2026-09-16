"""La raíz del proyecto se resuelve en UN solo lugar.

Cuando el código se movió a osocio/, cada módulo que calculaba la raíz por su cuenta con
`os.path.dirname(os.path.dirname(__file__))` quedó corto en un nivel: esa cuenta pasó a dar
`osocio/` en vez de la raíz. La app siguió abriendo sin quejarse y escribió en
`osocio/json/` y `osocio/temporales/` durante días.

Ningún test lo detectó porque todos parchean las rutas con monkeypatch para aislarse, así que
nunca ejercitaban la resolución real. Estos tests cubren justamente eso: que la resolución
verdadera sea correcta, y que nadie vuelva a calcularla por su cuenta.
"""
import ast
import os

import pytest

from osocio import paths


RAIZ_PAQUETE = os.path.dirname(os.path.abspath(paths.__file__ + "/../.."))


# --- la raíz es la que tiene run.py ------------------------------------------------

def test_base_dir_es_la_carpeta_que_contiene_run_py():
    """`run.py` es el punto de entrada y vive en la raíz: si BASE_DIR no lo ve, apunta mal."""
    assert os.path.isfile(os.path.join(paths.BASE_DIR, "run.py")), (
        f"BASE_DIR apunta a {paths.BASE_DIR}, donde no está run.py"
    )


def test_base_dir_no_es_la_carpeta_del_paquete():
    """El error concreto que se cometió: quedarse en osocio/ en vez de subir a la raíz."""
    assert os.path.basename(paths.BASE_DIR) != "osocio"


@pytest.mark.parametrize("nombre", ["JSON_DIR", "DATA_DIR", "RESULTS_DIR", "TEMPORALES_DIR", "FORMS_DIR"])
def test_las_carpetas_derivadas_cuelgan_de_base_dir(nombre):
    ruta = getattr(paths, nombre)
    assert ruta.startswith(paths.BASE_DIR), f"{nombre} = {ruta} no cuelga de BASE_DIR"


@pytest.mark.parametrize("nombre", ["JSON_DIR", "DATA_DIR", "RESULTS_DIR", "TEMPORALES_DIR"])
def test_las_carpetas_de_datos_no_caen_dentro_del_paquete(nombre):
    """json/, data/, resultados/ y temporales/ son del proyecto, no del paquete de código."""
    ruta = os.path.normpath(getattr(paths, nombre))
    dentro_del_paquete = os.path.normpath(os.path.join(paths.BASE_DIR, "osocio"))
    assert not ruta.startswith(dentro_del_paquete + os.sep), (
        f"{nombre} = {ruta} quedó adentro del paquete"
    )


# --- nadie más calcula la raíz -----------------------------------------------------

def _modulos_del_paquete():
    base = os.path.join(paths.BASE_DIR, "osocio")
    for carpeta, dirs, archivos in os.walk(base):
        dirs[:] = [d for d in dirs if d != "__pycache__"]
        for a in archivos:
            if a.endswith(".py"):
                yield os.path.join(carpeta, a)


def _cadenas_dirname_sobre_file(ruta):
    """Devuelve [(linea, cuántos dirname anidados)] de cada `dirname(...abspath(__file__))`."""
    with open(ruta, "r", encoding="utf-8") as fh:
        arbol = ast.parse(fh.read())

    encontrados = []
    for nodo in ast.walk(arbol):
        if not isinstance(nodo, ast.Call):
            continue
        if not (isinstance(nodo.func, ast.Attribute) and nodo.func.attr == "dirname"):
            continue
        # Contar cuántos dirname hay encadenados y si abajo de todo está __file__
        profundidad, actual = 0, nodo
        while (isinstance(actual, ast.Call)
               and isinstance(actual.func, ast.Attribute)
               and actual.func.attr in ("dirname", "abspath")):
            if actual.func.attr == "dirname":
                profundidad += 1
            if not actual.args:
                break
            actual = actual.args[0]
        usa_file = isinstance(actual, ast.Name) and actual.id == "__file__"
        if usa_file:
            encontrados.append((nodo.lineno, profundidad))
    return encontrados


def test_solo_paths_py_puede_calcular_la_raiz():
    """Una sola fuente de verdad.

    Un `dirname(abspath(__file__))` suelto está bien: es "mi propia carpeta". Encadenar dos o
    más es intentar llegar a la raíz, y eso lo hace `osocio/paths.py` y nadie más. Es
    exactamente la regla que, de haber existido, habría evitado el bug.
    """
    infractores = []
    for ruta in _modulos_del_paquete():
        if os.path.basename(ruta) == "paths.py":
            continue
        for linea, profundidad in _cadenas_dirname_sobre_file(ruta):
            if profundidad >= 2:
                rel = os.path.relpath(ruta, paths.BASE_DIR).replace(os.sep, "/")
                infractores.append(f"{rel}:{linea} ({profundidad} dirname encadenados)")

    assert not infractores, (
        "Estos módulos calculan la raíz por su cuenta en vez de importarla de "
        "osocio.paths:\n  " + "\n  ".join(infractores)
    )


# --- los módulos que se rompieron, uno por uno -------------------------------------

def _es_la_raiz(ruta):
    return os.path.normpath(ruta) == os.path.normpath(paths.BASE_DIR)


def test_scheduling_apunta_a_la_raiz():
    from osocio.utils import scheduling
    assert _es_la_raiz(scheduling.BASE_DIR)
    assert os.path.normpath(scheduling.JSON_DIR) == os.path.normpath(paths.JSON_DIR)


def test_autonomous_runner_apunta_a_la_raiz():
    from osocio import autonomous_runner
    assert _es_la_raiz(autonomous_runner.PROJECT_ROOT)
    assert os.path.normpath(autonomous_runner.JSON_DIR) == os.path.normpath(paths.JSON_DIR)
    assert os.path.normpath(autonomous_runner.RESULTS_DIR) == os.path.normpath(paths.RESULTS_DIR)


def test_popup_logger_escribe_el_runtime_log_en_la_raiz():
    from osocio.utils import popup_logger
    assert _es_la_raiz(popup_logger._get_base_dir())


def test_win_task_scheduler_apunta_a_la_raiz():
    from osocio.utils import win_task_scheduler
    assert _es_la_raiz(win_task_scheduler._project_root())


def test_validation_exporter_apunta_a_la_raiz():
    from osocio.validation import validation_exporter
    assert _es_la_raiz(validation_exporter._get_base_dir())


def test_field_validation_ui_apunta_a_la_raiz():
    from osocio.interface import field_validation_ui
    assert _es_la_raiz(field_validation_ui._get_base_dir())


def test_dealer_comparator_ui_apunta_a_la_raiz():
    from osocio.interface import dealer_comparator_ui
    assert _es_la_raiz(dealer_comparator_ui._get_base_dir())


def test_autovalores_usa_el_json_de_la_raiz_por_defecto():
    from osocio.utils.autovalores_campos_detectados import AutovaloresCamposDetectados
    autov = AutovaloresCamposDetectados("Peru")
    assert os.path.normpath(autov.json_dir) == os.path.normpath(paths.JSON_DIR)
