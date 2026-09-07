# Dependencias entre campos del formulario.
# Formato: {parent_id: child_id}
# El campo hijo sólo se vuelve visible/habilitado después de que el padre tiene un valor.
#
# Usado tanto en la validación de campos (selenium_validation_runner) como
# referencia para base_form_filler. NO modificar base_form_filler para importar
# desde aquí: ese módulo ya tiene sus propias dependencias inline.

import os
import json


_HARDCODED: dict[str, str] = {
    "models": "kits[]",
    "model": "kits[]",
    "region": "city",
    "city": "dealer",
}




def _cargar_dependencias_json() -> dict[str, str]:
    """Lee las dependencias definidas por el usuario en ids_dinamicos.json."""
    result: dict[str, str] = dict(_HARDCODED)
    try:
        if getattr(__import__("sys"), "frozen", False):
            base = os.path.dirname(__import__("sys").executable)
        else:
            base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        path = os.path.join(base, "json", "ids_dinamicos.json")
        if not os.path.exists(path):
            return result
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        for dep in data.get("dependencies", []):
            if isinstance(dep, dict):
                padre = str(dep.get("padre") or "").strip()
                hijo = str(dep.get("hijo") or "").strip()
                if padre and hijo:
                    result[padre] = hijo
        return result
    except Exception:
        return result


def get_field_dependencies(country: str | None = None) -> dict[str, str]:
    """Devuelve el mapa combinado hardcodeado + usuario (filtrado por país si se indica)."""
    combined = dict(_HARDCODED)
    # Load user-defined dependencies from JSON
    try:
        if getattr(__import__("sys"), "frozen", False):
            base = os.path.dirname(__import__("sys").executable)
        else:
            base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        path = os.path.join(base, "json", "ids_dinamicos.json")
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            for dep in data.get("dependencies", []):
                if not isinstance(dep, dict):
                    continue
                padre = str(dep.get("padre") or "").strip()
                hijo = str(dep.get("hijo") or "").strip()
                if not padre or not hijo:
                    continue
                paises = dep.get("paises", [])
                if not paises or country is None or country in paises:
                    combined[padre] = hijo
    except Exception:
        pass
    return combined


# Backward-compatible dict (sin filtrado por país) — usado por selenium_validation_runner
FIELD_DEPENDENCIES: dict[str, str] = _cargar_dependencias_json()
