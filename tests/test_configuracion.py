"""La configuración de json/ es la memoria de la app: sin ella no sabe llenar formularios.

Estos tests son de humo sobre esa carpeta. El caso que los motiva: cuando la resolución de
rutas se rompió, la app empezó a escribir un `ids_dinamicos.json` nuevo y truncado en
`osocio/json/`, mientras el bueno quedaba congelado en `json/`. El motor pasó a leer la
mitad sin dependencias y las cadenas region -> city -> dealer dejaron de aplicarse, en
silencio y sin que ningún test se enterara.
"""
import json
import os

import pytest

from osocio import paths


def _jsons_de_configuracion():
    if not os.path.isdir(paths.JSON_DIR):
        return []
    return sorted(
        os.path.join(paths.JSON_DIR, n)
        for n in os.listdir(paths.JSON_DIR)
        if n.endswith(".json")
    )


def _ids(rutas):
    return [os.path.basename(r) for r in rutas]


ARCHIVOS = _jsons_de_configuracion()


# --- la carpeta existe y es legible ------------------------------------------------

def test_la_carpeta_json_existe():
    assert os.path.isdir(paths.JSON_DIR), f"No existe {paths.JSON_DIR}"


def test_hay_configuracion_cargada():
    assert ARCHIVOS, "json/ está vacío: la app abriría sin saber llenar ningún formulario"


@pytest.mark.parametrize("ruta", ARCHIVOS, ids=_ids(ARCHIVOS))
def test_cada_json_parsea(ruta):
    """Un JSON corrupto rompe la corrida entera, y el error aparece recién en runtime."""
    with open(ruta, "r", encoding="utf-8") as fh:
        json.load(fh)


# --- ids_dinamicos.json: el que se partió en dos -----------------------------------

def _ids_dinamicos():
    ruta = os.path.join(paths.JSON_DIR, "ids_dinamicos.json")
    if not os.path.isfile(ruta):
        pytest.skip("no hay ids_dinamicos.json en este entorno")
    with open(ruta, "r", encoding="utf-8") as fh:
        return json.load(fh)


def test_ids_dinamicos_conserva_el_bloque_dependencies():
    """El archivo truncado que se escribió en el lugar equivocado no tenía esta clave.

    Es la señal barata de que se está leyendo el archivo bueno y no una copia a medias.
    """
    datos = _ids_dinamicos()
    assert "dependencies" in datos, (
        "ids_dinamicos.json no tiene 'dependencies': se está leyendo un archivo truncado"
    )
    assert datos["dependencies"], "'dependencies' está vacío"


def test_cada_dependencia_tiene_padre_e_hijo():
    for dep in _ids_dinamicos().get("dependencies", []):
        assert dep.get("padre"), f"dependencia sin padre: {dep}"
        assert dep.get("hijo"), f"dependencia sin hijo: {dep}"


# --- la prueba funcional: el motor tiene que VER esas dependencias -----------------

def test_el_motor_lee_las_dependencias_del_json_y_no_solo_las_hardcodeadas():
    """Esta es la que importa de verdad.

    `get_field_dependencies` arranca de una tabla hardcodeada corta (models, model, region,
    city) y la completa con lo que haya en ids_dinamicos.json. La dependencia
    document-type -> ci de Perú vive SOLO en el JSON, así que si el motor la ve es porque
    está leyendo el archivo correcto; si no la ve, está leyendo el truncado.
    """
    from osocio.core.field_dependencies import get_field_dependencies

    deps = get_field_dependencies("Peru")
    assert deps.get("document-type") == "ci", (
        "El motor no ve la dependencia document-type -> ci de Peru, que está en "
        "json/ids_dinamicos.json. Está leyendo otro archivo."
    )


def test_las_dependencias_hardcodeadas_siguen_estando():
    """Red de seguridad del test de arriba: si el JSON no se lee, estas igual aparecen,
    así que su presencia sola no prueba nada — pero su ausencia sí sería una regresión."""
    from osocio.core.field_dependencies import get_field_dependencies

    deps = get_field_dependencies()
    assert deps.get("region") == "city"
    assert deps.get("city") == "dealer"
