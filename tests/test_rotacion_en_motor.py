"""Rotación de región / ciudad / concesionario dentro del motor (sin navegador): un form
simulado con su cascada Región → Ciudad → Concesionario."""
from osocio.core.base_form_filler import BaseFormFiller
from osocio.utils.rotacion_dealers import CIUDAD, DEALER, REGION, RotacionDelDia


class _Opcion:
    def __init__(self, texto):
        self.text = texto


# Región → ciudades → concesionarios
CASCADA = {
    "VALPARAISO": {"VINA DEL MAR": ["KOVACS VINA"], "VALPARAISO": ["KOVACS VALPO", "SALAZAR VALPO"]},
    "METROPOLITANA": {"SANTIAGO": ["DIVEMOTOR", "AUTOMOTORA SUR"]},
}


def _motor(tmp_path, cascada=CASCADA, config=None):
    m = object.__new__(BaseFormFiller)
    m.config = dict({"pais": "Chile"}, **(config or {}))
    m._rotacion_dia = RotacionDelDia(carpeta=str(tmp_path), fecha="2026-10-23")
    m._rotacion_fila = {}
    m._aviso_rotacion = ""
    estado = {"region": None, "city": None, "dealer": None}
    cambios = []     # cada select que se volvió a elegir, en orden
    listo = {"city": True}   # False: la ciudad "no carga" después de cambiar de región

    def opciones(select_id):
        if select_id == "region":
            return list(cascada)
        if select_id == "city":
            return list(cascada.get(estado["region"], {}))
        return list(cascada.get(estado["region"], {}).get(estado["city"], []))

    def reelegir(select_id, tipo, excluir, solo=None):
        from osocio.utils.rotacion_dealers import _norm
        ops = [o for o in opciones(select_id) if _norm(o) not in {_norm(e) for e in excluir}]
        if solo is not None:
            ops = [o for o in ops if _norm(o) == _norm(solo)]
        cambios.append(select_id)
        if not ops:
            return None
        i = m._rotacion().elegir("Chile", tipo, ops)
        estado[select_id] = ops[i]
        if select_id == "region":
            estado["city"] = estado["dealer"] = None
        if select_id == "city":
            estado["dealer"] = None
        m._rotacion_anotar(tipo, select_id, ops[i], al_azar=True)
        return ops[i]

    m._reelegir_al_azar = reelegir
    m._wait_for_dependent_dropdown_ready = lambda select_id, *a, **k: listo.get(select_id, True)
    m.safe_find_element = lambda *a, **k: object()
    m._get_valid_select_options = lambda el: [_Opcion(t) for t in opciones("dealer")]
    m._estado = estado
    m._cambios = cambios
    m._listo = listo
    m._llenar_al_azar = lambda: [reelegir(s, t, ()) for s, t in (("region", REGION), ("city", CIUDAD), ("dealer", DEALER))]
    return m


def test_elegir_al_azar_rota_y_anota(tmp_path):
    m = _motor(tmp_path)
    m._rotacion_dia.registrar("Chile", {DEALER: "DIVEMOTOR"})
    elegida = m._elegir_al_azar("dealer", "Concesionario", [_Opcion("DIVEMOTOR"), _Opcion("AUTOMOTORA SUR")])
    assert elegida.text == "AUTOMOTORA SUR"
    assert m._rotacion_fila[DEALER] == {"id": "dealer", "texto": "AUTOMOTORA SUR", "al_azar": True}
    # Un campo que no es región/ciudad/concesionario no se anota.
    m._elegir_al_azar("models", "Modelo", [_Opcion("Onix")])
    assert set(m._rotacion_fila) == {DEALER}


def test_cambia_de_ciudad_y_region_hasta_un_concesionario_sin_usar(tmp_path):
    m = _motor(tmp_path)
    # Ya se usaron hoy todos los de Valparaíso: sólo quedan los de Santiago.
    m._rotacion_dia.registrar("Chile", {DEALER: "KOVACS VINA"})
    m._rotacion_dia.registrar("Chile", {DEALER: "KOVACS VALPO"})
    m._rotacion_dia.registrar("Chile", {DEALER: "SALAZAR VALPO"})
    m._estado.update(region="VALPARAISO", city="VINA DEL MAR", dealer="KOVACS VINA")
    m._rotacion_fila = {REGION: {"id": "region", "texto": "VALPARAISO", "al_azar": True},
                        CIUDAD: {"id": "city", "texto": "VINA DEL MAR", "al_azar": True},
                        DEALER: {"id": "dealer", "texto": "KOVACS VINA", "al_azar": True}}
    m._evitar_concesionario_repetido()
    assert m._estado["region"] == "METROPOLITANA"
    assert m._rotacion_fila[DEALER]["texto"] in ("DIVEMOTOR", "AUTOMOTORA SUR")
    assert not m._aviso_rotacion


def test_lo_del_excel_no_se_cambia_y_se_avisa(tmp_path):
    m = _motor(tmp_path)
    m._rotacion_dia.registrar("Chile", {DEALER: "KOVACS VINA"})
    m._estado.update(region="VALPARAISO", city="VINA DEL MAR", dealer="KOVACS VINA")
    m._rotacion_fila = {CIUDAD: {"id": "city", "texto": "VINA DEL MAR", "al_azar": False},   # del Excel
                        DEALER: {"id": "dealer", "texto": "KOVACS VINA", "al_azar": True}}
    m._evitar_concesionario_repetido()
    assert m._estado["city"] == "VINA DEL MAR"
    assert "concesionario repetido hoy: KOVACS VINA" in m._aviso_rotacion
    assert "único" in m._aviso_rotacion


def test_dealer_del_excel_no_se_toca(tmp_path):
    m = _motor(tmp_path)
    m._rotacion_dia.registrar("Chile", {DEALER: "DIVEMOTOR"})
    m._rotacion_fila = {DEALER: {"id": "dealer", "texto": "DIVEMOTOR", "al_azar": False}}
    m._evitar_concesionario_repetido()
    assert m._rotacion_fila[DEALER]["texto"] == "DIVEMOTOR" and not m._aviso_rotacion


def test_cinco_forms_seguidos_no_repiten_concesionario(tmp_path):
    m = _motor(tmp_path)
    usados = []
    for _ in range(5):   # hay 5 concesionarios en total
        m._rotacion_fila, m._aviso_rotacion = {}, ""
        m._llenar_al_azar()
        m._evitar_concesionario_repetido()
        usados.append(m._rotacion_fila[DEALER]["texto"])
        m._registrar_rotacion_fila()
    assert len(set(usados)) == 5


# Cadillac: los mismos 3 concesionarios en todas las ciudades y regiones.
CADILLAC = {r: {c: ["CADILLAC SAO PAULO", "CADILLAC CURITIBA", "CADILLAC BRASILIA"]
                for c in (f"{r} CIUDAD 1", f"{r} CIUDAD 2", f"{r} CIUDAD 3")}
            for r in ("MINAS GERAIS", "MATO GROSSO", "BAHIA")}


def _fila_cadillac(m):
    m._rotacion_fila, m._aviso_rotacion = {}, ""
    m._cambios.clear()
    m._llenar_al_azar()
    m._cambios.clear()


def test_concesionarios_iguales_en_todas_las_ciudades_no_hacen_loop(tmp_path):
    m = _motor(tmp_path, cascada=CADILLAC, config={"pais": "Brasil"})
    m.expected_form_url = "https://secure-developments.com/commonwealth/gm_frontend/cadillac/t1/brasil/form/vistiq"
    for d in ("CADILLAC SAO PAULO", "CADILLAC CURITIBA", "CADILLAC BRASILIA"):
        m._rotacion_dia.registrar("Brasil", {DEALER: d})
    _fila_cadillac(m)
    region, ciudad = m._estado["region"], m._estado["city"]
    m._evitar_concesionario_repetido()
    # Probó una sola ciudad, vio que da lo mismo y volvió a la que tenía.
    assert m._cambios.count("city") <= 2 and "region" not in m._cambios
    assert (m._estado["region"], m._estado["city"]) == (region, ciudad)
    assert m._estado["dealer"] and not m._aviso_rotacion
    # Fila siguiente del mismo form: ya no toca la ciudad.
    _fila_cadillac(m)
    m._evitar_concesionario_repetido()
    assert m._cambios == []


def test_si_no_encuentra_concesionario_libre_vuelve_a_lo_que_habia(tmp_path):
    m = _motor(tmp_path)
    for d in ("KOVACS VINA", "KOVACS VALPO", "SALAZAR VALPO", "DIVEMOTOR", "AUTOMOTORA SUR"):
        m._rotacion_dia.registrar("Chile", {DEALER: d})
    m._estado.update(region="VALPARAISO", city="VINA DEL MAR", dealer="KOVACS VINA")
    m._rotacion_fila = {REGION: {"id": "region", "texto": "VALPARAISO", "al_azar": True},
                        CIUDAD: {"id": "city", "texto": "VINA DEL MAR", "al_azar": True},
                        DEALER: {"id": "dealer", "texto": "KOVACS VINA", "al_azar": True}}
    m._evitar_concesionario_repetido()
    assert m._estado == {"region": "VALPARAISO", "city": "VINA DEL MAR", "dealer": "KOVACS VINA"}
    assert "concesionario repetido hoy" in m._aviso_rotacion


def test_si_la_ciudad_no_carga_tras_cambiar_de_region_no_queda_vacia(tmp_path):
    """Caso real (Cadillac Vistiq, fila 4): cambió de región, la ciudad no cargó a tiempo y
    el form quedó con la ciudad vacía. Ahora vuelve a lo que había."""
    m = _motor(tmp_path)
    m._rotacion_dia.registrar("Chile", {DEALER: "KOVACS VINA"})
    m._estado.update(region="VALPARAISO", city="VINA DEL MAR", dealer="KOVACS VINA")
    m._rotacion_fila = {REGION: {"id": "region", "texto": "VALPARAISO", "al_azar": True},
                        CIUDAD: {"id": "city", "texto": "VINA DEL MAR", "al_azar": True},
                        DEALER: {"id": "dealer", "texto": "KOVACS VINA", "al_azar": True}}
    # La única otra ciudad de Valparaíso tiene concesionarios usados; en Metropolitana,
    # la ciudad "no carga".
    m._rotacion_dia.registrar("Chile", {DEALER: "KOVACS VALPO"})
    m._rotacion_dia.registrar("Chile", {DEALER: "SALAZAR VALPO"})
    m._listo["city"] = False
    m._evitar_concesionario_repetido()
    assert m._estado["region"] and m._estado["city"] and m._estado["dealer"]
    assert m._estado["city"] in CASCADA[m._estado["region"]]
    assert m._estado["dealer"] in CASCADA[m._estado["region"]][m._estado["city"]]


def test_corrida_sin_envio_no_anota(tmp_path):
    m = _motor(tmp_path, config={"chequeo_campos": True})
    m._rotacion_fila = {DEALER: {"id": "dealer", "texto": "DIVEMOTOR", "al_azar": True}}
    m._registrar_rotacion_fila()
    assert m._rotacion_dia.veces("Chile", DEALER, "DIVEMOTOR") == 0 and m._rotacion_fila == {}


def test_iphone_elige_con_rotacion_por_etiqueta(tmp_path):
    from osocio.providers.browserstack.iphone_runner import FormFillerIphone
    m = object.__new__(FormFillerIphone)
    m.config = {"pais": "Chile"}
    m._rotacion_dia = RotacionDelDia(carpeta=str(tmp_path), fecha="2026-10-23")
    m._rotacion_fila, m._rotacion_excluir = {}, {}
    m._rotacion_dia.registrar("Chile", {DEALER: "KOVACS"})
    assert m._elegir_opcion_ios("CONCESIONARIO", ["KOVACS", "DIVEMOTOR"]) == "DIVEMOTOR"
    assert m._rotacion_fila[DEALER] == {"id": "CONCESIONARIO", "texto": "DIVEMOTOR", "al_azar": True}
    m._rotacion_excluir = {CIUDAD: {"VALPARAISO"}}
    assert m._elegir_opcion_ios("CIUDAD", ["VALPARAISO", "QUILLOTA"]) == "QUILLOTA"
    assert m._elegir_opcion_ios("MODELOS", ["Onix", "Tracker"]) is None    # sin rotación


def test_iphone_menu_usa_el_elegidor():
    from osocio.providers.browserstack.ios_nativo import FormNativoIOS
    f = object.__new__(FormNativoIOS)
    f.elegir_opcion = lambda etiqueta, ops: "QUILLOTA" if etiqueta == "CIUDAD" else None
    elegidas = []
    import xml.etree.ElementTree as ET
    T = "XCUIElementType"
    raiz = ET.Element(T + "Application")
    menu = ET.SubElement(raiz, T + "CollectionView")
    for t in ("Seleccione", "VALPARAISO", "QUILLOTA"):
        ET.SubElement(menu, T + "Button", label=t)
    f._fuente = lambda: raiz
    f.log = lambda *a: None

    class _Boton:
        def get_dom_attribute(self, n):
            return "true"
        def click(self):
            elegidas.append("click")

    class _D:
        def find_element(self, by, xp):
            elegidas.append(xp)
            return _Boton()
    f.d = _D()
    import osocio.providers.browserstack.ios_nativo as m
    dormir, m.time.sleep = m.time.sleep, (lambda s: None)
    try:
        assert f._elegir_en_menu(None, "CIUDAD") == "QUILLOTA"
    finally:
        m.time.sleep = dormir
    assert any("QUILLOTA" in str(x) for x in elegidas)


def test_iphone_concesionarios_iguales_no_hace_loop_y_respeta_lo_que_habia(tmp_path):
    """iPhone, form tipo Cadillac: prueba una ciudad, ve los mismos concesionarios, vuelve a
    la ciudad que tenía y en la fila siguiente ya no busca."""
    from osocio.providers.browserstack.iphone_runner import FormFillerIphone
    m = object.__new__(FormFillerIphone)
    m.config = {"pais": "Brasil"}
    m.expected_form_url = "https://x/gm_frontend/cadillac/t1/brasil/form/vistiq"
    m._rotacion_dia = RotacionDelDia(carpeta=str(tmp_path), fecha="2026-11-03")
    m._rotacion_excluir, m._aviso_rotacion = {}, ""
    dealers = ["CADILLAC SAO PAULO", "CADILLAC CURITIBA", "CADILLAC BRASILIA"]
    for d in dealers:
        m._rotacion_dia.registrar("Brasil", {DEALER: d})
    ciudades = ["CUIABA", "SINOP", "SORRISO"]
    toques = []

    class _Nativo:
        def elegir_de_nuevo(self, etiqueta):
            toques.append(etiqueta)
            return m._elegir_opcion_ios(etiqueta, ciudades)

        def llenar(self, valor_para, al_llenar=None):
            m._elegir_opcion_ios("CONCESIONARIO", dealers)
    m._nativo = _Nativo()
    m._registrar_campo_ios = lambda *a, **k: None

    def fila():
        m._rotacion_fila = {}
        m._elegir_opcion_ios("CIUDAD", ciudades)
        m._elegir_opcion_ios("CONCESIONARIO", dealers)
        toques.clear()
        return m._rotacion_fila[CIUDAD]["texto"]

    ciudad = fila()
    m._evitar_concesionario_repetido_ios(lambda *a: None)
    assert m._rotacion_fila[CIUDAD]["texto"] == ciudad and len(toques) == 2   # probó una y volvió
    assert m._rotacion_fila[DEALER]["texto"] in dealers
    fila()
    m._evitar_concesionario_repetido_ios(lambda *a: None)
    assert toques == []
