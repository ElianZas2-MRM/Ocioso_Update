"""En el iPhone de BrowserStack un elemento devuelto adentro de un objeto JS ({el: c})
no se puede volver a pasar a un script: se reemplaza por el de find_elements."""
from osocio.core.base_form_filler import BaseFormFiller


def test_reemplaza_por_los_de_find_elements():
    crudos = [{"el": "wdc-1", "name": "terms"}, {"el": "wdc-2", "name": "privacy"}]
    assert BaseFormFiller._con_elementos_reales(crudos, ["real-1", "real-2"]) == [
        {"el": "real-1", "name": "terms"}, {"el": "real-2", "name": "privacy"}]


def test_si_no_coincide_la_cantidad_no_toca_nada():
    crudos = [{"el": "wdc-1", "name": "terms"}]
    assert BaseFormFiller._con_elementos_reales(crudos, []) == [{"el": "wdc-1", "name": "terms"}]
