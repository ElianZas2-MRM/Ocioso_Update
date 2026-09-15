"""Todo lo que la app produce vive adentro de `resultados/`.

Antes había cuatro carpetas hermanas colgando de la raíz — `resultados/`,
`resultados_lambdatestmac/`, `resultados_lambdatest_android/` y
`Dealerscheck_resultados/` — así que después de una corrida había que buscar la evidencia
en cuatro lugares distintos, y ninguno decía a qué corrida pertenecía.

Ahora es una sola carpeta con subcarpetas por tipo de ejecución. Estos tests existen para
que no vuelva a dispersarse: es el tipo de cosa que se degrada sola, un `os.path.join`
a la vez.
"""
import os

import pytest

from osocio.utils import paths


SUBCARPETAS = [
    "RESULTS_LT_MAC_DIR",
    "RESULTS_LT_ANDROID_DIR",
    "RESULTS_DEALERS_DIR",
    "RESULTS_MASIVA_DIR",
]


@pytest.mark.parametrize("nombre", SUBCARPETAS)
def test_cada_tipo_de_resultado_cuelga_de_resultados(nombre):
    ruta = os.path.normpath(getattr(paths, nombre))
    raiz = os.path.normpath(paths.RESULTS_DIR)
    assert ruta.startswith(raiz + os.sep), (
        f"{nombre} = {ruta} quedó fuera de {raiz}"
    )


@pytest.mark.parametrize("nombre", SUBCARPETAS)
def test_son_subcarpetas_directas_y_no_nietas(nombre):
    """Un nivel alcanza: resultados/<tipo>. Más profundidad esconde la evidencia."""
    ruta = os.path.normpath(getattr(paths, nombre))
    relativa = os.path.relpath(ruta, paths.RESULTS_DIR)
    assert len(relativa.split(os.sep)) == 1, f"{nombre} está a más de un nivel: {relativa}"


def test_no_hay_dos_tipos_apuntando_a_la_misma_carpeta():
    rutas = {n: os.path.normpath(getattr(paths, n)) for n in SUBCARPETAS}
    assert len(set(rutas.values())) == len(rutas), (
        f"hay subcarpetas repetidas: {rutas}"
    )


def test_los_modulos_usan_las_rutas_centralizadas():
    """Cada runner tiene que tomar su carpeta de paths.py, no armarla por su cuenta."""
    from osocio.providers.lambdatest_mac import lt_runner
    from osocio.providers.lambdatest_android import lt_android_runner

    assert os.path.normpath(lt_runner._RESULTADOS_DIR) == os.path.normpath(paths.RESULTS_LT_MAC_DIR)
    assert os.path.normpath(lt_android_runner._ANDROID_RESULTADOS_DIR) == os.path.normpath(
        paths.RESULTS_LT_ANDROID_DIR
    )


def test_el_comparador_de_dealers_escribe_adentro_de_resultados():
    from osocio.interface.dealer_comparator_ui import _get_results_dir

    assert os.path.normpath(_get_results_dir()) == os.path.normpath(paths.RESULTS_DEALERS_DIR)


def test_el_ejecutor_autonomo_busca_donde_se_escribe():
    """Si el autónomo mirara la carpeta vieja, los mails saldrían sin adjuntos."""
    from osocio.autonomous_runner import _get_lt_results_dir

    assert os.path.normpath(_get_lt_results_dir("mac")) == os.path.normpath(paths.RESULTS_LT_MAC_DIR)
    assert os.path.normpath(_get_lt_results_dir("android")) == os.path.normpath(
        paths.RESULTS_LT_ANDROID_DIR
    )


def test_no_quedaron_carpetas_de_resultados_sueltas_en_la_raiz():
    """La raíz solo puede tener `resultados/`. Las otras tres se migraron adentro."""
    viejas = [
        "resultados_lambdatestmac",
        "resultados_lambdatest_android",
        "Dealerscheck_resultados",
        "resultados_dealers",
        "resultados_mac",
        "resultados_android",
    ]
    encontradas = [n for n in viejas if os.path.isdir(os.path.join(paths.BASE_DIR, n))]
    assert not encontradas, (
        "volvieron a aparecer carpetas de resultados sueltas en la raíz: " + ", ".join(encontradas)
    )


def test_ningun_modulo_arma_carpetas_de_resultados_a_mano():
    """El chequeo que impide que esto se desarme de a poco.

    Si alguien escribe os.path.join(BASE_DIR, "resultados_loquesea"), la evidencia vuelve
    a dispersarse y nadie se entera hasta que busca un reporte y no está.
    """
    import ast

    # Ojo con la diferencia: 'resultados_' tambien es el PREFIJO de los archivos que la app
    # genera (resultados_Argentina_Android1.xlsx), y eso es legitimo. Lo que no queremos son
    # nombres de CARPETA hermanos de resultados/, que es lo que estaba disperso.
    CARPETAS_PROHIBIDAS = {
        "resultados_lambdatestmac",
        "resultados_lambdatest_android",
        "Dealerscheck_resultados",
        "resultados_dealers",
        "resultados_mac",
        "resultados_android",
    }

    sospechosas = []
    base = os.path.join(paths.BASE_DIR, "osocio")
    for carpeta, dirs, archivos in os.walk(base):
        dirs[:] = [d for d in dirs if d != "__pycache__"]
        for a in archivos:
            if not a.endswith(".py"):
                continue
            ruta = os.path.join(carpeta, a)
            if os.path.basename(ruta) == "paths.py":
                continue
            with open(ruta, "r", encoding="utf-8") as fh:
                try:
                    arbol = ast.parse(fh.read())
                except SyntaxError:
                    continue
            for nodo in ast.walk(arbol):
                if not isinstance(nodo, ast.Constant) or not isinstance(nodo.value, str):
                    continue
                if nodo.value.strip("/\\") in CARPETAS_PROHIBIDAS:
                    rel = os.path.relpath(ruta, paths.BASE_DIR).replace(os.sep, "/")
                    sospechosas.append(f"{rel}:{nodo.lineno} -> {nodo.value!r}")

    assert not sospechosas, (
        "estos módulos nombran carpetas de resultados a mano en vez de usar osocio.utils.paths:"
        "\n  " + "\n  ".join(sospechosas)
    )
