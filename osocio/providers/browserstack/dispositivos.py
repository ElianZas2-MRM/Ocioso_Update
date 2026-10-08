"""
dispositivos.py
===============
Los dispositivos de BrowserStack tal como los ve la app. Es la ÚNICA tabla: la usan los
botones de Envío de Leads y Generar Excels, el despacho de sesiones, "Reintentar
fallidos", los mails y el runner. Antes cada uno tenía su copia de "Mac LT / Android LT".

Cada dispositivo tiene varias "caras" según quién lo pida:
  etiqueta      el botón de la UI                      "iPhone BS"
  clave         cómo se guarda en ui_prefs             "iphone bs"
  plataforma    la de bs_driver.DISPOSITIVOS            "iphone"
  dtype         el tipo de sesión en Envío de Leads     "bs_iphone"
  navegador     para resultados y mails                 "browserstack_iphone"
  sufijo_excel  el Excel de entrada                     ..._Iphone.xlsx
  token_email   el sufijo del email de los leads        ..._bsiphone@mrm.com
"""

from dataclasses import dataclass
from typing import Dict, Optional


@dataclass(frozen=True)
class DispositivoBS:
    etiqueta: str
    plataforma: str
    sufijo_excel: str
    nombre_largo: str

    @property
    def clave(self) -> str:
        return self.etiqueta.lower()

    @property
    def dtype(self) -> str:
        return f"bs_{self.plataforma}"

    @property
    def navegador(self) -> str:
        return f"browserstack_{self.plataforma}"

    @property
    def token_email(self) -> str:
        return f"bs{self.plataforma}"


DISPOSITIVOS_BS = (
    DispositivoBS("iPhone BS", "iphone", "Iphone", "BrowserStack iPhone 14 Pro Max (Safari)"),
    DispositivoBS("Android BS", "android", "Android", "BrowserStack Samsung Galaxy S26 Ultra (Chrome)"),
    DispositivoBS("Mac BS", "mac", "Mac", "BrowserStack Mac (Safari)"),
)

# Selección guardada por versiones anteriores (LambdaTest) -> equivalente en BrowserStack.
_CLAVES_LAMBDATEST = {"mac lt": "mac bs", "android lt": "android bs"}


def por_cualquier_nombre(nombre) -> Optional[DispositivoBS]:
    """El dispositivo, venga como etiqueta, clave, plataforma, dtype o navegador."""
    n = str(nombre or "").strip().lower()
    for d in DISPOSITIVOS_BS:
        if n in (d.clave, d.plataforma, d.dtype, d.navegador):
            return d
    return None


def es_browserstack(nombre) -> bool:
    return por_cualquier_nombre(nombre) is not None


def migrar_seleccion(guardada: Dict[str, bool]) -> Dict[str, bool]:
    """ui_prefs.selected_disp de una versión con LambdaTest: 'mac lt' pasa a 'mac bs'."""
    out = {}
    for k, v in (guardada or {}).items():
        out[_CLAVES_LAMBDATEST.get(k, k)] = v
    return out


# Sesiones de BrowserStack simultáneas que lanza la app. El plan tiene 10 en paralelo
# COMPARTIDAS por todo el equipo: por defecto 2, para no acaparar.
# Configurable en json/config_global.json -> {"browserstack": {"max_paralelo": N}}.
MAX_PARALELO_DEFAULT = 2
_MAX_PARALELO_PLAN = 10


def max_paralelo() -> int:
    try:
        import json
        import os
        from osocio.paths import JSON_DIR
        with open(os.path.join(JSON_DIR, "config_global.json"), "r", encoding="utf-8") as f:
            valor = int((json.load(f).get("browserstack") or {}).get("max_paralelo",
                                                                     MAX_PARALELO_DEFAULT))
    except Exception:
        return MAX_PARALELO_DEFAULT
    return max(1, min(valor, _MAX_PARALELO_PLAN))
