"""Tests del registro de evidencia en el Excel de resultado.

Todo campo que la herramienta llena tiene que quedar registrado con su dato, aunque no
estuviera en el Excel de entrada — de un campo auto-descubierto no se sabe que existe
hasta correr el form, pero el resultado sí tiene que dejar constancia de qué se envió.

El corte esta en `BaseFormFiller._sync_tracked_with_dom_before_submit`: el snapshot barre
TODA la pagina (`document.querySelectorAll`), asi que hay que separar un campo del form de
un input suelto de la landing. Se hace sin driver, mockeando el snapshot.
"""
import json
import os
import sys

import pytest

CORE_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "core")
if CORE_DIR not in sys.path:
    sys.path.insert(0, CORE_DIR)

from base_form_filler import BaseFormFiller  # noqa: E402


def _entry(eid, value, tag="input", visible=True):
    return {"id": eid, "realId": eid, "name": eid, "tag": tag,
            "type": "text" if tag != "select" else "select-one",
            "value": value, "text": value, "visible": visible}


def _filler(snapshot, *, detectados=(), ids_dinamicos=(), trackeado=None,
            historicos=(), tmp_path=None):
    """Instancia sin __init__ (no hay browser en los tests) con lo justo para correr el sync."""
    f = object.__new__(BaseFormFiller)
    f.current_row_field_values = dict(trackeado or {})
    f._campos_nuevos_detectados = [{"id": i} for i in detectados]
    f._ids_propios_cache = None
    f._dropdowns_sin_elegir = []
    f.field_mapping = []
    f._datos_vs_excel = ""
    f._datos_mismatch = False
    f._snapshot_form_state = lambda: snapshot
    f._cargar_ids_dinamicos = lambda: {i: ["x"] for i in ids_dinamicos}

    # json/nuevos_campos_<pais>.json: el histórico de campos auto-descubiertos.
    f.config = {"pais": "Peru" if historicos else ""}
    f.BASE_DIR = str(tmp_path) if tmp_path else ""
    if historicos and tmp_path:
        carpeta = os.path.join(str(tmp_path), "json")
        os.makedirs(carpeta, exist_ok=True)
        with open(os.path.join(carpeta, "nuevos_campos_peru.json"), "w", encoding="utf-8") as fh:
            json.dump({"campos_nuevos": [{"id": i} for i in historicos]}, fh)
    return f


def _finales(filler):
    return {k.split("::", 1)[1]: v
            for k, v in filler.current_row_field_values.items() if k.startswith("Final::")}


def test_campo_auto_descubierto_de_texto_queda_registrado():
    """El caso pedido: `contract` no está en el Excel de entrada porque nadie sabía que el
    form lo tenía, pero se llenó y tiene que aparecer en el resultado con su dato."""
    filler = _filler({"contract": _entry("contract", "5372819044")},
                     detectados=["contract"])
    filler._sync_tracked_with_dom_before_submit()
    assert _finales(filler) == {"contract": "5372819044"}


def test_campo_de_ids_dinamicos_queda_registrado():
    filler = _filler({"adviser": _entry("adviser", "042")}, ids_dinamicos=["adviser"])
    filler._sync_tracked_with_dom_before_submit()
    assert _finales(filler) == {"adviser": "042"}


def test_input_suelto_de_la_landing_no_ensucia_el_excel():
    """El snapshot ve toda la página: un buscador o un newsletter tienen valor pero no son
    parte del lead, y registrarlos agregaría una columna por corrida."""
    filler = _filler({"search": _entry("search", "camioneta"),
                      "newsletter-email": _entry("newsletter-email", "a@b.com")})
    filler._sync_tracked_with_dom_before_submit()
    assert _finales(filler) == {}


def test_select_no_trackeado_se_registra_aunque_no_sea_campo_propio():
    """Comportamiento previo que no debe cambiar: un select con valor siempre viaja."""
    filler = _filler({"region": _entry("region", "Córdoba", tag="select")})
    filler._sync_tracked_with_dom_before_submit()
    assert _finales(filler) == {"region": "Córdoba"}


def test_campo_propio_pero_oculto_no_se_registra():
    filler = _filler({"contract": _entry("contract", "123", visible=False)},
                     detectados=["contract"])
    filler._sync_tracked_with_dom_before_submit()
    assert _finales(filler) == {}


def test_campo_propio_vacio_no_se_registra():
    filler = _filler({"contract": _entry("contract", "")}, detectados=["contract"])
    filler._sync_tracked_with_dom_before_submit()
    assert _finales(filler) == {}


def test_campo_ya_trackeado_no_se_duplica_como_final():
    filler = _filler({"contract": _entry("contract", "5372819044")},
                     detectados=["contract"],
                     trackeado={"Paso1::contract": "5372819044"})
    filler._sync_tracked_with_dom_before_submit()
    assert _finales(filler) == {}
    assert filler.current_row_field_values["Paso1::contract"] == "5372819044"


def test_valor_real_del_dom_pisa_al_trackeado():
    """Si el form terminó con otro valor que el que creíamos, manda el del DOM."""
    filler = _filler({"contract": _entry("contract", "9999999999")},
                     detectados=["contract"],
                     trackeado={"Paso1::contract": "324"})
    filler._sync_tracked_with_dom_before_submit()
    assert filler.current_row_field_values["Paso1::contract"] == "9999999999"


def test_sin_snapshot_no_revienta():
    filler = _filler({})
    filler._sync_tracked_with_dom_before_submit()
    assert _finales(filler) == {}


# --- _ids_llenados_por_la_herramienta ---

def test_junta_auto_descubiertos_e_ids_dinamicos():
    filler = _filler({}, detectados=["contract", "comments"], ids_dinamicos=["adviser"])
    assert filler._ids_llenados_por_la_herramienta() == {"contract", "comments", "adviser"}


def test_usa_el_historico_cuando_el_discovery_ya_no_reporta_nada(tmp_path):
    """A la segunda corrida de un form, `_campos_nuevos_detectados` viene vacío porque el
    auto-discovery sólo reporta lo que todavía no conocía. Sin leer el histórico, la
    evidencia se perdería justo en los formularios ya conocidos — que son la mayoría."""
    filler = _filler({"cc_address": _entry("cc_address", "Av. Javier Prado Este 1234")},
                     detectados=(), historicos=["cc_name", "cc_address"], tmp_path=tmp_path)

    assert "cc_address" in filler._ids_llenados_por_la_herramienta()
    filler._sync_tracked_with_dom_before_submit()
    assert _finales(filler) == {"cc_address": "Av. Javier Prado Este 1234"}


def test_historico_ilegible_no_tumba_la_corrida(tmp_path):
    filler = _filler({}, detectados=["contract"], historicos=["x"], tmp_path=tmp_path)
    with open(os.path.join(str(tmp_path), "json", "nuevos_campos_peru.json"), "w") as fh:
        fh.write("{ esto no es json")
    filler._ids_propios_cache = None
    assert filler._ids_llenados_por_la_herramienta() == {"contract"}


def test_ignora_entradas_sin_id_y_tolera_el_store_roto():
    filler = _filler({}, detectados=["contract"])
    filler._campos_nuevos_detectados = [{"id": "contract"}, {"id": "  "}, {}, None]

    def _explota():
        raise RuntimeError("json corrupto")

    filler._cargar_ids_dinamicos = _explota
    # El registro de evidencia no puede tumbar una corrida por un store ilegible.
    assert filler._ids_llenados_por_la_herramienta() == {"contract"}


# --- dropdowns que quedaron en su placeholder ---

def _select(eid, texto):
    return {"id": eid, "realId": eid, "name": eid, "tag": "select", "type": "select-one",
            "value": texto, "text": texto, "visible": True}


@pytest.mark.parametrize("placeholder", [
    "Seleccionar", "Seleccione", "Selecione", "Seleccioná", "Escolha",
    "Select", "Elija una opción",
])
def test_dropdown_en_placeholder_no_se_registra_como_valor(placeholder):
    """El Excel decía `models = Seleccionar`, que es justo lo contrario de lo que pasó:
    ese dropdown no se eligió. No puede figurar como dato del lead."""
    filler = _filler({"models": _select("models", placeholder)},
                     trackeado={"Paso1::models": placeholder})
    filler._sync_tracked_with_dom_before_submit()

    assert "Paso1::models" not in filler.current_row_field_values
    assert "models" in filler._dropdowns_sin_elegir
    assert "sin elegir" in filler._datos_vs_excel
    assert filler._datos_mismatch is True


def test_dropdown_con_opcion_real_si_se_registra():
    filler = _filler({"models": _select("models", "Captiva Híbrida")},
                     trackeado={"Paso1::models": "Captiva Híbrida"})
    filler._sync_tracked_with_dom_before_submit()

    assert filler.current_row_field_values["Paso1::models"] == "Captiva Híbrida"
    assert filler._dropdowns_sin_elegir == []
    assert filler._datos_vs_excel == "OK"


def test_un_input_de_texto_que_dice_seleccionar_no_se_descarta():
    """El filtro es sólo para <select>: en un campo de texto libre, la palabra podría ser
    un dato legítimo que el usuario escribió."""
    filler = _filler({"comments": _entry("comments", "Seleccionar modelo nuevo")},
                     detectados=["comments"],
                     trackeado={"Paso1::comments": "Seleccionar modelo nuevo"})
    filler._sync_tracked_with_dom_before_submit()

    assert filler.current_row_field_values["Paso1::comments"] == "Seleccionar modelo nuevo"
    assert filler._dropdowns_sin_elegir == []


# --- forms multi-paso: el chequeo tiene que correr al cerrar cada paso ---

def _filler_con_selects(selects):
    f = _filler({})
    f.driver = type("D", (), {"execute_script": staticmethod(lambda *a, **k: selects)})()
    f._current_step = 1
    return f


def test_registra_dropdowns_en_placeholder_al_cerrar_el_paso():
    """En multi-paso el paso anterior puede salir del DOM: si no se mira al cerrarlo, el
    chequeo previo al submit ya no lo alcanza y el dropdown pasa como completado."""
    filler = _filler_con_selects([
        {"id": "region", "texto": "Seleccionar"},
        {"id": "models", "texto": "Captiva Híbrida"},
        {"id": "dealer", "texto": "Selecione"},
    ])
    filler.current_row_field_values = {"Paso1::region": "Seleccionar",
                                       "Paso1::models": "Captiva Híbrida"}

    nuevos = filler._registrar_dropdowns_en_placeholder()

    assert sorted(nuevos) == ["dealer", "region"]
    assert "Paso1::region" not in filler.current_row_field_values, "no puede figurar como dato"
    assert filler.current_row_field_values["Paso1::models"] == "Captiva Híbrida"


def test_no_duplica_un_dropdown_ya_anotado():
    filler = _filler_con_selects([{"id": "region", "texto": "Seleccione"}])
    filler._registrar_dropdowns_en_placeholder()
    filler._registrar_dropdowns_en_placeholder()
    assert filler._dropdowns_sin_elegir == ["region"]


def test_si_el_js_falla_no_tumba_el_paso():
    filler = _filler({})
    def _explota(*a, **k):
        raise RuntimeError("driver caido")
    filler.driver = type("D", (), {"execute_script": staticmethod(_explota)})()
    assert filler._registrar_dropdowns_en_placeholder() == []
