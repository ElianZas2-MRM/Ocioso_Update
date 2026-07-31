"""Generadores de datos válidos por país para el tab 'Generar Excels'."""

import random
import unicodedata

FIRST_NAMES = [
    "Carlos", "Juan", "Luis", "Diego", "Miguel", "Andrés", "Pablo", "Sebastián",
    "Matías", "Alejandro", "Fernando", "Gabriel", "Martín", "Ricardo", "Eduardo",
    "Rafael", "Felipe", "Cristian", "Rodrigo", "Javier", "Marcelo", "Gonzalo",
    "Nicolás", "Emilio", "Ramón", "Héctor", "Ignacio", "Tomás", "Agustín",
    "María", "Ana", "Sofia", "Valentina", "Camila", "Isabella", "Daniela", "Laura",
    "Carolina", "Andrea", "Fernanda", "Paula", "Gabriela", "Natalia", "Catalina",
    "Lucía", "Valeria", "Claudia", "Monica", "Patricia", "Florencia", "Agustina",
    "Romina", "Verónica", "Susana", "Lorena", "Paola", "Sandra", "Cecilia",
]

LAST_NAMES = [
    "García", "Rodriguez", "Martinez", "Lopez", "Gonzalez", "Perez", "Sanchez",
    "Ramirez", "Torres", "Flores", "Morales", "Herrera", "Mendoza", "Ortega",
    "Castro", "Vargas", "Rojas", "Diaz", "Cruz", "Reyes", "Gutierrez", "Jimenez",
    "Romero", "Alvarez", "Ruiz", "Navarro", "Molina", "Delgado", "Vega", "Soto",
    "Rios", "Paredes", "Campos", "Fuentes", "Quispe", "Mamani", "Silva", "Medina",
    "Aguilar", "Suarez", "Castillo", "Ramos", "Espinoza", "Acosta", "Miranda",
    "Pinto", "Salazar", "Muñoz", "Ibáñez", "Figueroa", "Bravo", "Contreras",
    "Villanueva", "Serrano", "Mora", "Guerrero", "Cabrera", "Nuñez", "Valdez",
]

PAIS_ABREV = {
    "Argentina": "ar", "Bolivia": "bo", "Brasil": "br", "Chile": "cl",
    "Colombia": "co", "Ecuador": "ec", "Paraguay": "py", "Peru": "pe", "Uruguay": "uy",
}

SPECIAL_CHARS = [".", "_"]


def _strip_accents(text):
    return "".join(
        c for c in unicodedata.normalize("NFD", text) if unicodedata.category(c) != "Mn"
    )


def _clean(text):
    """Elimina acentos, espacios y caracteres no alfanuméricos."""
    return _strip_accents(text).lower().replace(" ", "")


def generar_nombre():
    count = random.choice([1, 2])
    return " ".join(random.sample(FIRST_NAMES, count))


def generar_apellido():
    count = random.choice([1, 2])
    return " ".join(random.sample(LAST_NAMES, count))


def generar_email(nombre, apellido, pais="", device=None):
    """
    Formato: {nombre_partes}{apellido_partes}{special}{pais_abrev}{nn}{_device}@mrm.com
    Incluye todos los tokens de nombre y apellido, abreviatura de país, 2 dígitos, un carácter especial y dispositivo.
    """
    # Incluir todos los tokens (1 o 2) según cuántas palabras haya
    nombre_parts = "".join(_clean(p) for p in nombre.split())
    apellido_parts = "".join(_clean(p) for p in apellido.split())
    abrev = PAIS_ABREV.get(pais, "")
    special = random.choice(SPECIAL_CHARS)
    numero = f"{random.randint(1, 99):02d}"
    
    device_suffix = ""
    if device:
        dev_clean = device.lower().replace("lambdatest_", "").replace("emulado-navegador", "").replace("desktop", "").strip()
        if dev_clean:
            device_suffix = f"_{dev_clean}"
            
    return f"{nombre_parts}{apellido_parts}{special}{abrev}{numero}{device_suffix}@mrm.com"


# ── Generadores de documento ─────────────────────────────────────────────────

def generar_rut_chile():
    """RUT chileno válido con dígito verificador (8-9 chars totales: 7-8 dígitos + DV)."""
    while True:
        num = random.randint(5_000_000, 25_000_000)
        digits = [int(d) for d in str(num)]
        multiplicadores = [2, 3, 4, 5, 6, 7]
        suma = sum(d * multiplicadores[i % 6] for i, d in enumerate(reversed(digits)))
        resto = 11 - (suma % 11)
        dv = "0" if resto == 11 else ("K" if resto == 10 else str(resto))
        rut = f"{num}{dv}"
        if 8 <= len(rut) <= 9:
            return rut


def generar_rut_chile_con_k():
    """RUT chileno válido cuyo dígito verificador es K."""
    multiplicadores = [2, 3, 4, 5, 6, 7]
    while True:
        num = random.randint(5_000_000, 25_000_000)
        digits = [int(d) for d in str(num)]
        suma = sum(d * multiplicadores[i % 6] for i, d in enumerate(reversed(digits)))
        if (11 - suma % 11) == 10 and 8 <= len(str(num)) + 1 <= 9:
            return f"{num}K"


def generar_ci_ecuador():
    """Cédula de identidad ecuatoriana válida (10 dígitos)."""
    provincia = random.randint(1, 24)
    tercero = random.randint(0, 6)
    resto = [random.randint(0, 9) for _ in range(6)]
    dgts = [int(d) for d in f"{provincia:02d}"] + [tercero] + resto
    coef = [2, 1, 2, 1, 2, 1, 2, 1, 2]
    suma = 0
    for d, c in zip(dgts, coef):
        p = d * c
        suma += (p - 9) if p >= 10 else p
    check = (10 - (suma % 10)) % 10
    return "".join(str(d) for d in dgts) + str(check)


def generar_cpf_brasil():
    """CPF brasileño válido (11 dígitos). Fallback local — rechaza dígitos todos iguales."""
    while True:
        nums = [random.randint(0, 9) for _ in range(9)]
        if len(set(nums)) == 1:
            continue
        soma = sum((10 - i) * nums[i] for i in range(9))
        resto = soma % 11
        d1 = 0 if resto < 2 else 11 - resto
        nums.append(d1)
        soma = sum((11 - i) * nums[i] for i in range(10))
        resto = soma % 11
        d2 = 0 if resto < 2 else 11 - resto
        nums.append(d2)
        return "".join(str(d) for d in nums)


def generar_cnpj_brasil():
    """CNPJ brasileño válido (14 dígitos sin puntuación). Fallback local con algoritmo oficial."""
    b = [random.randint(0, 9) for _ in range(8)] + [0, 0, 0, 1]
    w1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    r1 = sum(v * w for v, w in zip(b, w1)) % 11
    c1 = 0 if r1 < 2 else 11 - r1
    w2 = [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    r2 = sum(v * w for v, w in zip(b + [c1], w2)) % 11
    c2 = 0 if r2 < 2 else 11 - r2
    return "".join(str(d) for d in b + [c1, c2])


# CEPs válidos conocidos como último recurso si la API falla
_CEPS_FALLBACK = [
    "01310100", "04538133", "20040020", "30112010",
    "40015970", "60060100", "80010010", "90010280",
]


def generar_documento_brasil_4devs(tipo: str = "cpf") -> str:
    """Genera CPF, CNPJ o CEP via API de 4devs. Fallback local correcto por tipo si falla la red."""
    import urllib.request, urllib.parse, re as _re
    def _fetch(acao, extra=None):
        data = {"acao": acao}
        if extra:
            data.update(extra)
        body = urllib.parse.urlencode(data).encode()
        req = urllib.request.Request(
            "https://www.4devs.com.br/ferramentas_online.php",
            data=body,
            headers={"Content-Type": "application/x-www-form-urlencoded",
                     "User-Agent": "Mozilla/5.0"},
        )
        with urllib.request.urlopen(req, timeout=10) as r:
            return r.read().decode("utf-8").strip()

    tipo = (tipo or "cpf").lower()
    try:
        if tipo == "cnpj":
            raw = _fetch("gerar_cnpj", {"pontuacao": "N"})
            digits = "".join(c for c in raw if c.isdigit())
            if len(digits) != 14:
                raise ValueError(f"CNPJ inválido de API: {raw!r}")
            return digits
        if tipo == "cep":
            html = _fetch("gerar_cep", {"estado": "", "cidade": "São Paulo",
                                         "bairro": "", "tipo_cep": "residencial"})
            m = _re.search(r'(\d{5}-\d{3})', html)
            if m:
                return m.group(1).replace("-", "")
            m2 = _re.search(r'\d{8}', html)
            if m2:
                return m2.group(0)
            raise ValueError("CEP inválido de API: " + repr(html[:100]))
        # CPF (default)
        raw = _fetch("gerar_cpf", {"pontuacao": "N"})
        digits = "".join(c for c in raw if c.isdigit())
        if len(digits) != 11:
            raise ValueError(f"CPF inválido de API: {raw!r}")
        return digits
    except Exception:
        # Fallback local con formato correcto por tipo (antes siempre devolvía CPF para todo)
        from utils.popup_logger import log_runtime
        log_runtime(f"⚠️ API 4devs falló para '{tipo}', usando generador local como fallback", level="WARNING")
        if tipo == "cnpj":
            return generar_cnpj_brasil()
        if tipo == "cep":
            return random.choice(_CEPS_FALLBACK)
        return generar_cpf_brasil()


# Países cuyos formularios pueden tener MÁS DE UN campo de documento a la vez
# (ej. Brasil: CPF y CEP, o CNPJ y CEP). Estos usan columnas semánticas separadas en
# el Excel en vez de una sola "Documento". Los países fuera de este dict siguen igual.
DOC_TYPES_BY_COUNTRY = {
    "Brasil": ["CPF", "CNPJ", "CEP", "VIN"],
}


# Valores fijos por formulario (match por substring de la URL del form). Se aplican al
# generar el Excel, sobrescribiendo los aleatorios. Útil cuando un form exige un vehículo
# real (VIN + modelo que coinciden). El usuario puede editarlos a mano en el Excel.
FIXED_FORM_VALUES = {
    "clubemyev": {"Modelo": "Spark EUV", "VIN": "LK6ADAE3XTB075680"},
}


def fixed_values_for_url(url):
    """Devuelve el dict {columna: valor} fijo si la URL matchea un form conocido, o {}."""
    u = str(url or "").lower()
    for key, vals in FIXED_FORM_VALUES.items():
        if key in u:
            return vals
    return {}


# VIN: 17 caracteres, sin I/O/Q. Dígito verificador en la posición 9 (ISO 3779 / NHTSA).
_VIN_CHARS = "ABCDEFGHJKLMNPRSTUVWXYZ0123456789"
_VIN_TRANSLIT = {
    **{str(d): d for d in range(10)},
    'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5, 'F': 6, 'G': 7, 'H': 8,
    'J': 1, 'K': 2, 'L': 3, 'M': 4, 'N': 5, 'P': 7, 'R': 9,
    'S': 2, 'T': 3, 'U': 4, 'V': 5, 'W': 6, 'X': 7, 'Y': 8, 'Z': 9,
}
_VIN_WEIGHTS = [8, 7, 6, 5, 4, 3, 2, 10, 0, 9, 8, 7, 6, 5, 4, 3, 2]


def generar_vin():
    """Genera un VIN válido de 17 caracteres con dígito verificador correcto."""
    chars = [random.choice(_VIN_CHARS) for _ in range(17)]
    total = sum(_VIN_TRANSLIT[chars[i]] * _VIN_WEIGHTS[i] for i in range(17))
    resto = total % 11
    chars[8] = 'X' if resto == 10 else str(resto)  # posición 9 = verificador (peso 0)
    return "".join(chars)


def _generar_doc_por_tipo(tipo):
    """Genera un documento brasileño del tipo pedido. CPF/CNPJ/VIN locales (instantáneos);
    CEP vía 4devs con fallback a lista de CEPs válidos."""
    t = (tipo or "").lower()
    if t == "cpf":
        return generar_cpf_brasil()
    if t == "cnpj":
        return generar_cnpj_brasil()
    if t == "cep":
        return generar_documento_brasil_4devs("cep")
    if t == "vin":
        return generar_vin()
    return ""


def generar_documento(pais):
    if pais == "Chile":
        return generar_rut_chile()
    if pais == "Ecuador":
        return generar_ci_ecuador()
    if pais == "Brasil":
        return generar_cpf_brasil()
    if pais == "Uruguay":
        return str(random.randint(10_000_000, 99_999_999))
    if pais == "Paraguay":
        # La CI paraguaya en los forms actuales acepta exactamente 7 dígitos
        # (minlength = maxlength = 7): con 8 el form recorta el último.
        return str(random.randint(1_000_000, 9_999_999))
    if pais == "Peru":
        # DNI peruano: exactamente 8 dígitos (los forms validan "Ingresa al menos 8 dígitos").
        # El rango genérico podía devolver 7 y el form lo rechazaba.
        return str(random.randint(10_000_000, 99_999_999))
    return str(random.randint(1_000_000, 99_999_999))


# ── Generadores de celular ───────────────────────────────────────────────────

def generar_celular(pais):
    def digits(n):
        return "".join(str(random.randint(0, 9)) for _ in range(n))

    if pais == "Argentina":
        area = random.choice(["11", "351", "261", "221", "341", "381", "299"])
        return area + digits(10 - len(area))

    if pais == "Bolivia":
        prefix = random.choice(["6", "7"])
        return prefix + digits(7)

    if pais == "Brasil":
        ddd = str(random.randint(11, 98))
        return ddd + "9" + digits(8)

    if pais == "Chile":
        return "9" + digits(8)

    if pais == "Colombia":
        return "3" + digits(9)

    if pais == "Ecuador":
        return "09" + digits(8)

    if pais == "Paraguay":
        return "09" + digits(8)

    if pais == "Peru":
        return "9" + digits(8)

    if pais == "Uruguay":
        return "09" + digits(7)

    return digits(9)


# ── Generador de fila completa ───────────────────────────────────────────────

def generar_fila_datos(pais, device=None, doc_types=None):
    """Devuelve dict {nombre_columna: valor} para el país dado.
    doc_types: solo para países en DOC_TYPES_BY_COUNTRY (ej. Brasil). dict {tipo: bool}
    para elegir qué documentos generar (CPF/CNPJ/CEP); default = todos. Esos países usan
    columnas semánticas (CPF/CNPJ/CEP) en vez de la única 'Documento'."""
    nombre = generar_nombre()
    apellido = generar_apellido()
    fila = {
        "Modelo": "",
        "Nombre": nombre,
        "Apellido": apellido,
        "Documento": generar_documento(pais),
        "Celular": generar_celular(pais),
        "Email": generar_email(nombre, apellido, pais, device),
        # Campos de selección aleatoria en el formulario → vacíos
        "Región": "",
        "Region": "",
        "Ciudad": "",
        "Concesionario": "",
        "Fecha estimada": "",
        "Fecha estimada de compra": "",
        "Fecha de compra": "",
        # Campos extra opcionales → vacíos
        "Patente": "",
        "Evento": "",
        "Chasis": "",
        "Chasis/Vin": "",
        "Código asesor": "",
        "Comentario": "",
        "Tipo de documento": "",
        "Tipo de documento (Perú)": "",
        # Perú: el select 'document-type' define cuántos dígitos exige el nº de documento
        # (DNI 8 / RUC 11 / Carné de Extranjería y Pasaporte 12). Si la columna va vacía, el
        # form filler termina eligiendo una opción al azar DESPUÉS de haber escrito el
        # documento y el form rechaza el lead por longitud. Se fija DNI, que es el tipo que
        # matchea el 'Documento' de 8 dígitos que genera generar_documento("Peru").
        # El nombre de la columna sale de country_configs (Peru → 'Tipo de Documento').
        "Tipo de Documento": "DNI" if pais == "Peru" else "",
        "Kilometraje": "",
        "Año de adquisición": "",
        "Seguro": "",
        "Color": "",
        "Kit": "",
        "Dispositivo": device or "",
    }

    # Países con múltiples documentos (Brasil): columnas semánticas CPF/CNPJ/CEP en vez
    # de "Documento". Se generan solo los tipos tildados (doc_types); default = todos.
    tipos = DOC_TYPES_BY_COUNTRY.get(pais)
    if tipos:
        fila.pop("Documento", None)
        sel = doc_types if isinstance(doc_types, dict) else {t: True for t in tipos}
        for t in tipos:
            fila[t] = _generar_doc_por_tipo(t) if sel.get(t, False) else ""

    return fila
