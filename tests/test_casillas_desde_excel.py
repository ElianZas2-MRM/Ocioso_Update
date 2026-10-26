"""Radios y grupos de checkboxes pedidos desde el Excel (lógica pura, sin driver).

- Radio: columna = name del grupo, UNA opción ("have-interest" = "Renovar").
- Grupo de checkboxes: columna = name, opciones separadas por coma ("WhatsApp, Llamada");
  las demás quedan sin marcar.
- Checkbox suelto: SI/NO como siempre. '-' deja el campo sin tocar.
Las opciones se leen del form en cada corrida: nada de listas fijas en el código.
"""
from osocio.core.base_form_filler import BaseFormFiller

marcar = BaseFormFiller.opciones_a_marcar


def _filler(headers=(), row=()):
    f = BaseFormFiller.__new__(BaseFormFiller)
    f._campos_omitidos = []
    f._opciones_no_encontradas = []
    f.registrados = {}
    f._record_field_value = lambda campo, valor: f.registrados.__setitem__(campo, valor)
    f._cargar_ids_dinamicos = lambda: {}
    f.checkbox_prefs = f.build_checkbox_prefs(list(headers), list(row))
    f.checkbox_opciones = f.build_checkbox_opciones(list(headers), list(row))
    f.radio_prefs = f.build_radio_prefs(list(headers), list(row))
    return f


CANALES = [("sms", "SMS"), ("whatsapp", "WhatsApp"), ("todas", "Todas"),
           ("email", "Correo electrónico"), ("call", "Llamada")]


# --- parseo del Excel ---------------------------------------------------------------

def test_lista_separada_por_coma_punto_y_coma_o_barra():
    assert BaseFormFiller.lista_de_opciones("WhatsApp, Llamada") == ["WhatsApp", "Llamada"]
    assert BaseFormFiller.lista_de_opciones("SMS;Correo | Llamada,") == ["SMS", "Correo", "Llamada"]


def test_si_no_queda_en_prefs_y_listas_en_opciones():
    f = _filler(["product_info", "contact-channels"], ["SI", "WhatsApp, Llamada"])
    assert f.checkbox_prefs == {"product_info": True}
    assert f.checkbox_opciones["contact-channels"] == ["WhatsApp", "Llamada"]
    assert "product_info" not in f.checkbox_opciones


# --- coincidencia con las opciones del form -------------------------------------------

def test_por_texto_o_por_value_sin_acentos_ni_mayusculas():
    assert marcar(["whatsapp", "LLAMADA"], CANALES) == ({1, 4}, [])
    assert marcar(["email"], CANALES) == ({3}, [])                 # por value


def test_contiene_solo_si_no_hay_exacta_y_por_palabra_completa():
    assert marcar(["Correo"], CANALES) == ({3}, [])                # "Correo electrónico"
    opciones = [("a", "Correo"), ("b", "Correo y SMS")]
    assert marcar(["Correo"], opciones) == ({0}, [])               # la exacta gana
    assert marcar(["Si"], [("x", "Sin preferencia")]) == (set(), ["Si"])


def test_opcion_que_el_form_no_tiene_se_informa():
    assert marcar(["WhatsApp", "Telegram"], CANALES) == ({1}, ["Telegram"])


def test_opcion_nueva_en_el_form_se_puede_pedir_sin_tocar_codigo():
    """Si mañana el form agrega "Telegram", el Excel la pide y se encuentra sola."""
    hoy = CANALES + [("telegram", "Telegram")]
    assert marcar(["Telegram"], hoy) == ({5}, [])


# --- grupo de checkboxes en el DOM ----------------------------------------------------

def _checkboxes(name, opciones):
    return [{"name": name, "id": f"{name}-{v}", "value": v, "texto": t} for v, t in opciones]


def test_grupo_marca_las_pedidas_y_desmarca_las_demas():
    f = _filler(["contact-channels"], ["WhatsApp, Llamada"])
    decision = f._grupos_pedidos_checkbox(_checkboxes("contact-channels", CANALES))
    assert decision == {0: False, 1: True, 2: False, 3: False, 4: True}
    assert f.registrados["contact-channels"] == "WhatsApp | Llamada"


def test_grupo_con_guion_no_se_toca():
    f = _filler(["contact-channels"], ["-"])
    decision = f._grupos_pedidos_checkbox(_checkboxes("contact-channels", CANALES))
    assert set(decision.values()) == {None}


def test_grupo_avisa_opcion_no_encontrada_una_sola_vez():
    f = _filler(["contact-channels"], ["Telegram"])
    f._grupos_pedidos_checkbox(_checkboxes("contact-channels", CANALES))
    f._grupos_pedidos_checkbox(_checkboxes("contact-channels", CANALES))   # 2ª pasada
    assert f._opciones_no_encontradas == ["contact-channels: 'Telegram'"]


def test_checkbox_suelto_sigue_con_si_no():
    """Un valor que no es SI/NO en un checkbox suelto no lo desmarca: sigue la regla de siempre."""
    f = _filler(["terms"], ["X"])
    assert f._grupos_pedidos_checkbox(_checkboxes("terms", [("1", "Acepto")])) == {}


# --- radios -------------------------------------------------------------------------

def _grupo_radio(name, opciones):
    return {"name": name, "original_name": name,
            "options": [{"value": v, "texto": t} for v, t in opciones]}


def test_radio_por_texto_del_label_aunque_el_value_sea_otro():
    f = _filler(["have-interest"], ["Suscribir"])
    elegida = f._choose_radio_option(_grupo_radio("have-interest", [("1", "Renovar"), ("2", "Suscribir")]))
    assert elegida["value"] == "2"
    assert f.registrados["have-interest"] == "Suscribir"


def test_radio_con_mas_de_dos_opciones():
    f = _filler(["cc-claim-type"], ["Queja"])
    opciones = [("RECLAMO", "Reclamo"), ("QUEJA", "Queja"), ("SUGERENCIA", "Sugerencia")]
    assert f._choose_radio_option(_grupo_radio("cc-claim-type", opciones))["value"] == "QUEJA"


# --- iPhone: el traductor name <-> texto ----------------------------------------------
#
# La capa nativa ve textos, no el name. El traductor (leído del form abierto solo) dice
# qué name tiene cada grupo y sus opciones; acá se simula con un dict.

def _iphone(headers, row, casillas):
    from osocio.providers.browserstack.iphone_runner import FormFillerIphone
    f = FormFillerIphone.__new__(FormFillerIphone)
    f._campos_omitidos = []
    f._opciones_no_encontradas = []
    f._nombres_ios, f._checkbox_ios_marcadas = {}, {}
    f._cargar_ids_dinamicos = lambda: {}
    f.checkbox_prefs = f.build_checkbox_prefs(headers, row)
    f.checkbox_opciones = f.build_checkbox_opciones(headers, row)
    f.radio_prefs = f.build_radio_prefs(headers, row)
    f._casillas_form = casillas
    return f


ONSTAR = {
    "have-interest": {"tipo": "radio", "titulo": "¿Qué te interesa?",
                      "opciones": [{"valor": "1", "texto": "Renovar"}, {"valor": "2", "texto": "Suscribir"}]},
    "have-chevrolet": {"tipo": "radio", "titulo": "¿Tenés un Chevrolet?",
                       "opciones": [{"valor": "si", "texto": "Sí"}, {"valor": "no", "texto": "No"}]},
    "test-drive": {"tipo": "radio", "titulo": "¿Querés un test drive?",
                   "opciones": [{"valor": "si", "texto": "Sí"}, {"valor": "no", "texto": "No"}]},
}


def test_iphone_radio_por_el_conjunto_de_opciones():
    f = _iphone(["have-interest"], ["Suscribir"], ONSTAR)
    assert f._radio_pedido_ios("¿QUÉ TE INTERESA?", ["Renovar", "Suscribir"]) == "Suscribir"


def test_iphone_dos_grupos_si_no_se_distinguen_por_el_titulo():
    f = _iphone(["have-chevrolet", "test-drive"], ["No", "Sí"], ONSTAR)
    assert f._radio_pedido_ios("¿TENÉS UN CHEVROLET? *", ["Sí", "No"]) == "No"
    assert f._radio_pedido_ios("¿Querés un test drive?", ["Sí", "No"]) == "Sí"
    # Sin título que coincida no se adivina: al azar.
    assert f._radio_pedido_ios("Otra pregunta", ["Sí", "No"]) is None


def test_iphone_radio_en_pedazos_marca_solo_en_el_que_tiene_la_opcion():
    """PE libro de reclamaciones: las opciones largas, una debajo de otra, llegan a la capa
    nativa como preguntas de UNA opción ("radio@13")."""
    from osocio.providers.browserstack.ios_nativo import OMITIR
    reclamo = "RECLAMO: DISCONFORMIDAD RELACIONADA A LOS PRODUCTOS O SERVICIOS"
    queja = "QUEJA: DISCONFORMIDAD NO RELACIONADA A LOS PRODUCTOS O SERVICIOS"
    casillas = {"cc-claim-type": {"tipo": "radio", "titulo": "",
                                  "opciones": [{"valor": "RECLAMO", "texto": reclamo},
                                               {"valor": "QUEJA", "texto": queja}]}}
    f = _iphone(["cc-claim-type"], ["Queja"], casillas)
    assert f._radio_pedido_ios("radio@13", [reclamo]) is OMITIR
    assert f._radio_pedido_ios("radio@2", [queja]) == queja


def test_iphone_pedazos_sin_titulo_con_el_mismo_nombre_son_preguntas_distintas():
    """Con el scroll, el pedazo de QUEJA puede llamarse "radio@2" como uno de RECLAMO ya
    hecho: se recuerda por nombre y opciones, así que QUEJA sigue pendiente."""
    from osocio.providers.browserstack.ios_nativo import FormNativoIOS
    n = object.__new__(FormNativoIOS)
    n._radios_hechos = set()
    reclamo = {"tipo": "radio", "etiqueta": "radio@2", "opciones": [{"texto": "RECLAMO"}]}
    queja = {"tipo": "radio", "etiqueta": "radio@2", "opciones": [{"texto": "QUEJA"}]}
    n._radios_hechos.add(n._clave_radio(reclamo))
    assert not n._pendiente(reclamo) and n._pendiente(queja)


def test_iphone_al_azar_nunca_menor_de_edad():
    from osocio.providers.browserstack.ios_nativo import FormNativoIOS
    n = object.__new__(FormNativoIOS)
    n._radios_hechos = set()
    solo_menor = {"tipo": "radio", "etiqueta": "EN CASO DE MENORES DE EDAD",
                  "opciones": [{"texto": "SOY MENOR DE EDAD", "tap": (0, 0)}]}
    assert n._elegir_radio(solo_menor, None) == "(no se toca)"


def test_lectura_por_html_de_radios_checkboxes_y_listas():
    from osocio.providers.browserstack.iphone_runner import casillas_del_html
    html = """
      <label><input type="radio" name="business_person" value="Sí"> Sí</label>
      <label><input type="radio" name="business_person" value="No"> No</label>
      <label for="cp1">Preferencia de contacto</label><input type="checkbox" id="cp1" name="contact_preference" value="E-mail">
      <label for="cp2">Preferencia de contacto</label><input type="checkbox" id="cp2" name="contact_preference" value="WhatsApp">
      <select name="hobbies[]" multiple><option value="" disabled>Elige</option>
        <option value="Surf">Surf</option><option value="Campo">Campo</option></select>"""
    c = casillas_del_html(html)
    assert [o["texto"] for o in c["business_person"]["opciones"]] == ["Sí", "No"]
    assert [o["valor"] for o in c["contact_preference"]["opciones"]] == ["E-mail", "WhatsApp"]
    assert [o["texto"] for o in c["hobbies[]"]["opciones"]] == ["Surf", "Campo"]
    assert c["hobbies[]"]["tipo"] == "multi"


def test_iphone_radio_con_guion_se_omite():
    from osocio.providers.browserstack.ios_nativo import OMITIR
    f = _iphone(["have-interest"], ["-"], ONSTAR)
    assert f._radio_pedido_ios("¿Qué te interesa?", ["Renovar", "Suscribir"]) is OMITIR


CANALES_FORM = {
    "contact-channels": {"tipo": "checkbox", "titulo": "Preferencia de contacto",
                         "opciones": [{"valor": v, "texto": t} for v, t in CANALES]},
    "product_info": {"tipo": "checkbox", "titulo": "",
                     "opciones": [{"valor": "1", "texto": "Me gustaría recibir información"}]},
}


def _cb(etiqueta, valor="0", grupo=""):
    return {"tipo": "checkbox", "etiqueta": etiqueta, "valor": valor, "grupo": grupo}


def test_iphone_grupo_de_checkboxes_marca_las_pedidas_y_desmarca_el_resto():
    f = _iphone(["contact-channels"], ["WhatsApp, Llamada"], CANALES_FORM)
    assert f._preferencia_checkbox_ios(_cb("WhatsApp", grupo="PREFERENCIA DE CONTACTO")) is True
    assert f._preferencia_checkbox_ios(_cb("SMS", valor="1", grupo="PREFERENCIA DE CONTACTO")) is False
    assert f._checkbox_ios_marcadas == {"contact-channels": ["WhatsApp"]}


def test_iphone_checkbox_suelto_si_no_y_sin_columna():
    f = _iphone(["product_info"], ["NO"], CANALES_FORM)
    assert f._preferencia_checkbox_ios(_cb("Me gustaría recibir información")) is False
    assert f._preferencia_checkbox_ios(_cb("Acepto los términos")) is None   # no está en el form leído


ENVOLVE_CL = [("E-mail", "Preferencia de contacto"), ("Llamada telefónica", "Preferencia de contacto"),
              ("WhatsApp", "Preferencia de contacto")]


def test_label_igual_al_titulo_del_grupo_se_usa_el_value():
    """CL gm_forms/envolve: el <label for> de cada opción es "Preferencia de contacto"."""
    f = _filler(["contact_preference"], ["Llamada, WhatsApp"])
    decision = f._grupos_pedidos_checkbox(_checkboxes("contact_preference", ENVOLVE_CL))
    assert decision == {0: False, 1: True, 2: True}
    assert f.registrados["contact_preference"] == "Llamada telefónica | WhatsApp"
    casillas = {"contact_preference": {"tipo": "checkbox", "titulo": "",
                                       "opciones": [{"valor": v, "texto": t} for v, t in ENVOLVE_CL]}}
    ios = _iphone(["contact_preference"], ["WhatsApp"], casillas)
    assert ios._preferencia_checkbox_ios(_cb("WhatsApp", grupo="Preferencia de contacto")) is True
    assert ios._preferencia_checkbox_ios(_cb("E-mail", valor="1", grupo="Preferencia de contacto")) is False


def test_iphone_lista_multiple_por_orden_del_form():
    casillas = {"hobbies[]": {"tipo": "multi", "titulo": "¿Cuáles son tus pasatiempos?",
                              "opciones": [{"valor": v, "texto": v} for v in
                                           ("Deportes", "Música", "Viajes", "Cine", "Lectura")]}}
    f = _iphone(["hobbies[]"], ["Viajes, Lectura"], casillas)
    plan = f._opciones_multi_ios({"tipo": "multi", "etiqueta": "¿CUÁLES SON TUS PASATIEMPOS?"})
    assert plan["indices"] == [2, 4] and plan["total"] == 5
    sin_excel = _iphone([], [], casillas)._opciones_multi_ios({"tipo": "multi", "etiqueta": "x"})
    assert sin_excel == {"total": 5, "indices": None}     # al azar entre las 5 que hay hoy


def test_iphone_pendiente_respeta_el_excel_para_desmarcar():
    """En la capa nativa: con el Excel pidiendo NO, un checkbox marcado queda pendiente
    (para desmarcarlo) y uno sin marcar no."""
    from osocio.providers.browserstack.ios_nativo import FormNativoIOS
    n = object.__new__(FormNativoIOS)
    n._radios_hechos, n._grupos_hechos, n.marcar_todos = set(), set(), False
    n.preferencia_checkbox = lambda c: False
    assert n._pendiente(_cb("SMS", valor="1"))
    assert not n._pendiente(_cb("SMS", valor="0"))
    n.preferencia_checkbox = lambda c: True
    assert n._pendiente(_cb("WhatsApp", valor="0"))


def test_radio_que_no_existe_va_al_azar_y_se_avisa():
    f = _filler(["have-chevrolet"], ["Tal vez"])
    elegida = f._choose_radio_option(_grupo_radio("have-chevrolet", [("si", "Sí"), ("no", "No")]))
    assert elegida["value"] in ("si", "no")
    assert f._opciones_no_encontradas == ["have-chevrolet: 'Tal vez'"]
