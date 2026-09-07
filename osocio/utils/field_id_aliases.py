"""
field_id_aliases.py — Alias de IDs de campo entre las distintas generaciones de formularios GM.

El mismo campo no siempre tiene el mismo id segun el estandar con el que se construyo el
formulario. Ejemplo concreto: el Nombre es 'firstname' en los forms clasicos y 'name' en los
del estandar visid / gm_frontend (los .../gm_frontend/chevrolet/t3/<pais>/form/<slug>).

Como el mapping por pais esta escrito con los ids clasicos, cuando un id no aparece en el DOM
hay que reintentar con su alias antes de dar el campo por ausente. Esta tabla es la unica
fuente: la usan el motor de Envio de Leads (core/base_form_filler.py) y la Validacion de
Campos (validation/selenium_validation_runner.py).

Si en una migracion aparecen ids nuevos que difieren del estandar actual, se agregan aca.
"""

VISID_ID_ALIASES = {
    "firstname":               "name",
    "models":                  "model",
    "model_1":                 "model",
    "model_2":                 "model",
    "estimated-date-purchase": "estimated-day",
    "estimated-date":          "estimated-day",
    "estimated_date_purchase": "estimated-day",
    # gm_front / gm_frontend / alianzas modernas
    "telephone":               "phone",
    "cellphone":               "phone",
    "ci":                      "document",
}

# Alias en el sentido inverso (id nuevo -> id clasico), para reglas escritas con el id nuevo
# corriendo contra un form viejo.
_REVERSE_ALIASES = {}
for _clasico, _nuevo in VISID_ID_ALIASES.items():
    _REVERSE_ALIASES.setdefault(_nuevo, []).append(_clasico)


def alias_ids_for(element_id):
    """Ids alternativos a probar para element_id, sin repetir el original."""
    if not element_id or not isinstance(element_id, str):
        return []
    candidatos = []
    directo = VISID_ID_ALIASES.get(element_id)
    if directo:
        candidatos.append(directo)
    for inverso in _REVERSE_ALIASES.get(element_id, []):
        if inverso not in candidatos:
            candidatos.append(inverso)
    return [c for c in candidatos if c != element_id]
