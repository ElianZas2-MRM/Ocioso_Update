"""Humo del paquete entero: que todo importe y que las piezas base respondan.

Un smoke test no prueba que algo esté bien, prueba que no está roto de entrada. Este
archivo cubre lo más barato y lo más rentable: que los 53 módulos se puedan importar.

Suena trivial, pero es exactamente lo que faltaba. Durante el refactor a `osocio/` hubo
un momento en que `import lt_controller` quedó convertido en
`import osocio.providers...lt_controller`, que liga el nombre `osocio` y no
`lt_controller` — el módulo importaba bien y el código explotaba recién al usarlo.
"""
import importlib
import pkgutil

import pytest

import osocio


def _todos_los_modulos():
    return sorted(m.name for m in pkgutil.walk_packages(osocio.__path__, "osocio."))


MODULOS = _todos_los_modulos()


def test_el_paquete_tiene_modulos():
    assert len(MODULOS) > 40, f"solo se encontraron {len(MODULOS)} módulos, algo falta"


@pytest.mark.parametrize("nombre", MODULOS, ids=lambda n: n.replace("osocio.", ""))
def test_cada_modulo_importa(nombre):
    """Si un módulo no importa, la app revienta al abrir esa pantalla y no antes."""
    importlib.import_module(nombre)


def test_el_entrypoint_importa():
    """run.py es el único arranque: si su cadena de imports se corta, no abre nada."""
    importlib.import_module("run")


# --- los nombres que el codigo usa de verdad ---------------------------------------

@pytest.mark.parametrize("modulo, atributo", [
    ("osocio.core.base_form_filler", "BaseFormFiller"),
    ("osocio.core.generic_country_base", "GenericCountryBase"),
    ("osocio.core.browser_manager", "BrowserManager"),
    ("osocio.forms._runner_common", "get_runner"),
    ("osocio.providers.lambdatest_mac.lt_controller", "run"),
    ("osocio.providers.lambdatest_android.lt_android_controller", "run"),
    ("osocio.autonomous_runner", "run_once"),
    ("osocio.utils.scheduling", "guardar_programacion"),
    ("osocio.utils.driver_updater", "ensure_drivers_ready"),
])
def test_el_simbolo_publico_existe(modulo, atributo):
    """Importar el módulo no alcanza: el nombre que se usa tiene que estar ahí.

    Es el modo de falla que tuvimos: `import a.b.c` liga `a`, no `c`, así que el import
    pasa y el atributo no aparece hasta que alguien lo llama en runtime.
    """
    mod = importlib.import_module(modulo)
    assert hasattr(mod, atributo), f"{modulo} no expone {atributo}"
