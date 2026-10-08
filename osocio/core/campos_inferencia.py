"""
campos_inferencia.py
====================
Reglas para reconocer un campo o un botón de un form por su texto, compartidas por los
motores remotos (antes vivían en lambdatest_mac/lt_runner.py, que las reexporta).

- Qué dato del lead va en un campo, a partir de su id / name / etiqueta.
- Qué opción de un desplegable es un placeholder ("Seleccionar", "Escolha"...) y no vale.
- Qué texto de botón es "Siguiente" y cuál es "Enviar".

Neutro a propósito: el iPhone de BrowserStack lo usa sin depender de archivos de
LambdaTest, que se van a retirar.
"""

_PLACEHOLDER_KW = (
    "seleccione", "selecciona", "seleccionar", "seleccion",
    "selecione", "selecionar", "selecao", "escolha", "escolher",
    "elija", "elegir", "opcao", "opciones", "opcion",
    "porfavor", "favor", "obrigatorio", "required",
    "select", "choose", "please",
)


def _is_placeholder(text: str) -> bool:
    t = (text or "").lower().strip()
    if not t or "--" in t:
        return True
    return any(kw in t for kw in _PLACEHOLDER_KW)


_INFER_DATAKEY_MAP = {
    "firstname":  ("nombre", "name", "first", "fname", "nome", "given"),
    "lastname":   ("apellido", "lastname", "surname", "lname", "sobrenome", "family"),
    "email":      ("email", "correo", "mail", "e-mail", "emailaddress"),
    "phone":      ("telefono", "celular", "phone", "mobile", "cel", "tel", "fono", "movil",
                   "whatsapp", "celphone", "telefone"),
    "document":   ("documento", "doc", "cedula", "dni", "rut", "ci", "cpf",
                   "rfc", "identif", "passport", "pasaporte"),
    "model":      ("modelo", "model", "vehiculo", "veiculo", "auto", "car"),
}


def _normalize_for_infer(text):
    import unicodedata
    nfkd = unicodedata.normalize("NFKD", (text or "").lower())
    ascii_str = "".join(c for c in nfkd if not unicodedata.combining(c))
    return ascii_str.replace(" ", "").replace("-", "").replace("_", "")


def _infer_data_key(field_id, field_name, field_placeholder=""):
    tokens = [
        _normalize_for_infer(field_id),
        _normalize_for_infer(field_name),
        _normalize_for_infer(field_placeholder),
    ]
    for data_key, keywords in _INFER_DATAKEY_MAP.items():
        for token in tokens:
            if token and any(kw in token or token in kw for kw in keywords):
                return data_key
    return None


def _get_inferred_value(inferred_key: str, lead) -> str:
    if not lead or not inferred_key:
        return ""
    lead_lower = {k.lower().strip(): v for k, v in lead.data.items()}
    INFERRED_TO_HEADERS = {
        "firstname": ["firstname", "name", "nombre", "nombres", "fullname", "full_name"],
        "lastname":  ["lastname", "apellido", "apellidos", "last_name"],
        "email":     ["email", "correo", "e-mail"],
        "phone":     ["phone", "celular", "telefono", "cellphone", "telephone", "telephone-mask"],
        "document":  ["document", "documento", "cpf", "cnpj", "dni", "rut", "ci"],
        "model":     ["model", "modelo", "models"],
    }
    headers = INFERRED_TO_HEADERS.get(inferred_key, [inferred_key])
    for h in headers:
        val = lead_lower.get(h, "")
        if val and str(val).strip().lower() not in ("", "none"):
            return str(val)
    return str(lead_lower.get(inferred_key, ""))


_NEXT_KEYWORDS = (
    "siguiente", "seguinte", "continuar", "continuacao",
    "proximo", "próximo", "next",
)

_SUBMIT_KEYWORDS = (
    "enviar", "submit", "confirmar", "finalizar", "enviar dados",
)
