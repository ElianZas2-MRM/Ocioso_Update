"""Tests de BaseFormFiller._decidir_marca_checkbox (lógica pura, sin driver).

Prioridad acordada con el usuario para checkboxes:
1) Excel / IDs únicos (por id o name — uno de los dos siempre existe) manda siempre.
2) Sin preferencia explícita: sólo se marca si es requerido (HTML required/aria-required)
   o es un checkbox de términos/privacidad conocido. Los opcionales quedan como están
   (antes se marcaba cualquier checkbox visible sin distinguir).
"""
import os
import sys

CORE_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "core")
if CORE_DIR not in sys.path:
    sys.path.insert(0, CORE_DIR)

from base_form_filler import BaseFormFiller  # noqa: E402

decidir = BaseFormFiller._decidir_marca_checkbox


def test_excel_no_desmarca_sin_importar_lo_demas():
    assert decidir(is_known=True, is_required=True, pref=False, tiene_identificador=True) == "uncheck"


def test_excel_si_marca_sin_importar_lo_demas():
    assert decidir(is_known=False, is_required=False, pref=True, tiene_identificador=True) == "mark"


def test_sin_preferencia_requerido_se_marca():
    assert decidir(is_known=False, is_required=True, pref=None, tiene_identificador=True) == "mark"


def test_sin_preferencia_termino_conocido_se_marca():
    assert decidir(is_known=True, is_required=False, pref=None, tiene_identificador=True) == "mark"


def test_sin_preferencia_opcional_no_se_toca():
    # El caso que cambia respecto del comportamiento viejo: antes esto se marcaba igual.
    assert decidir(is_known=False, is_required=False, pref=None, tiene_identificador=True) == "skip"


def test_sin_identificador_se_saltea_aunque_sea_requerido():
    assert decidir(is_known=False, is_required=True, pref=None, tiene_identificador=False) == "skip"
