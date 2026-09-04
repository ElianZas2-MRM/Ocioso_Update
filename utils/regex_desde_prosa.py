"""Deriva un `regex_full` desde la prosa del excel de validaciones del CRM.

La columna DESCRIPCION y las columnas MENSAJE DE ERROR del excel son prosa libre en
español (portugués en la hoja de Brasil), no regex. Este modulo la traduce a un patron
concreto para que los campos que hoy quedan sin validacion puedan generar un valor de
prueba plausible en vez de quedarse vacios o con un valor inventado.

Contrato, calcado del de `utils.valor_campo_generator`: si la prosa no alcanza para
derivar un patron con confianza, devuelve "" — nunca adivina. Todo regex emitido cae
dentro de la gramatica que `GeneradorValorCampo` sabe parsear, y se verifica con un
round-trip real (se le pide un valor al generador) antes de devolverlo.

Precedencia deliberada, porque el excel se contradice a si mismo en varios campos:
- la CANTIDAD sale del mensaje de error, que es el texto literal que el form muestra
  cuando rechaza;
- la CLASE de caracteres sale de la descripcion, que es la que la califica
  ("caracteres numericos" / "alfanumericos" / "alfabeticos").

Caso testigo — NUMERO DE CONTRATO (Argentina):
    descripcion: "De 1 a 10 caracteres numéricos. No permite iniciar con 0."
    mensaje:     "Ingresá 10 caracteres que no comiencen con 0."
    -> ^[1-9][0-9]{9}$   (10 exactos: gana el mensaje; la clase la pone la descripcion)
"""
import re
import unicodedata

from utils.valor_campo_generator import GeneradorValorCampo

__all__ = ["derivar_regex", "ajustar_largo_desde_prosa", "regex_por_semantica"]

# Clases emitidas. Se mantienen dentro de lo que `_expandir_clase()` de
# valor_campo_generator sabe expandir (rangos simples + chars sueltos).
_CLASE_DIGITOS = "[0-9]"
_CLASE_LETRAS = "[a-zA-ZáéíóúÁÉÍÓÚñÑ ]"
_CLASE_ALNUM = "[a-zA-Z0-9]"

# Patron de email tal cual vive en los field_validation_rules_*.json: usarlo textual
# hace que el generador entre por su rama `_generar_email` y devuelva @mrm.com.
_REGEX_EMAIL = (
    r"^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?"
    r"(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)+$"
)

# Unidad de conteo. Sin ella no se deriva cantidad: "entre el 001 y 100" es un rango de
# VALORES, no de longitud, y tomarlo como largo produciria un patron absurdo.
_UNIDAD = r"(?:digitos?|numeros?|caracteres?|letras?|caraceteres?)"

# Prosa que no describe la forma del campo y solo mete ruido en la derivacion.
_RUIDO = (
    re.compile(r"\bmensaje de error \d+\b"),        # placeholders del propio excel
    re.compile(r"\bmayor(?:es)? de \d+ anos?\b"),   # regla de negocio (edad), no forma
    re.compile(r"\bmayor de edad\b"),
)


def _normalizar(texto):
    """lower + sin acentos + espacios colapsados. El excel mezcla 'mínimo'/'minimo' y
    'dígitos'/'digitos' segun la hoja, asi que el matcheo se hace sobre texto plano."""
    txt = str(texto or "").strip().lower()
    txt = unicodedata.normalize("NFD", txt)
    txt = "".join(c for c in txt if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", txt)


def _es_util(texto):
    """False si el texto esta vacio o es puro ruido conocido."""
    t = _normalizar(texto)
    if not t:
        return False
    return not any(rx.search(t) for rx in _RUIDO)


def _clase_de(texto):
    """Clase de caracteres que el texto califica, o "" si no la nombra."""
    t = _normalizar(texto)
    # "alfanumerico" contiene "numerico": el chequeo mas especifico va primero.
    if "alfanumeric" in t:
        return _CLASE_ALNUM
    if re.search(r"\b(?:digitos?|numeros?)\b|numeric", t):
        return _CLASE_DIGITOS
    if re.search(r"\b(?:letras?)\b|alfabetic", t):
        return _CLASE_LETRAS
    return ""


def _cantidad_de(texto):
    """(forma, lo, hi) o None. Se prueba de la forma mas especifica a la mas general."""
    t = _normalizar(texto)

    # "Insira 11/14 digitos" -> dos largos alternativos
    m = re.search(rf"(\d+)\s*/\s*(\d+)\s*{_UNIDAD}", t)
    if m:
        return ("alt", int(m.group(1)), int(m.group(2)))

    # "entre 6 y 7 caracteres" / "De 5 a 12 digitos" / "De 1 a 10 caracteres numericos"
    m = re.search(rf"(?:entre|de)\s+(\d+)\s+(?:y|a)\s+(\d+)\s*{_UNIDAD}", t)
    if m:
        return ("rango", int(m.group(1)), int(m.group(2)))

    # "al menos 7 digitos" / "minimo 2 letras" / "un minimo de 7 digitos"
    m = re.search(rf"(?:al menos|como minimo|minimo)\s+(?:de\s+)?(\d+)\s*{_UNIDAD}", t)
    if m:
        return ("min", int(m.group(1)), None)

    # "maximo 12 caracteres" / "hasta 12 caracteres" / "tiene un maximo de 120 caracteres"
    # El "de" opcional importa: sin el, "un maximo de 120 caracteres" cae al patron generico
    # de mas abajo y se lee como 120 EXACTOS en vez de hasta 120.
    m = re.search(rf"(?:maximo|hasta)\s+(?:de\s+)?(\d+)\s*{_UNIDAD}", t)
    if m:
        return ("max", 1, int(m.group(1)))

    # "2 caracteres numericos maximo"
    m = re.search(rf"(\d+)\s*{_UNIDAD}[^.]*\bmaximo\b", t)
    if m:
        return ("max", 1, int(m.group(1)))

    # "Ingresa 17 digitos" / "Completa 9 numeros" / "El CUIT debe tener 11 digitos"
    m = re.search(rf"(\d+)\s*{_UNIDAD}", t)
    if m:
        return ("exacto", int(m.group(1)), int(m.group(1)))

    return None


_RX_NO_INICIA = re.compile(
    r"no\s+(?:se\s+)?(?:puede|permite|debe|deben|pueden)?\s*"
    r"(?:empezar|comenzar|iniciar|comiencen|empiecen|inicie|inicien)\s+con\s+(?:el\s+)?(\d)"
)

_RX_PREFIJO = re.compile(
    r"(?:comenzar|comienzan|comienza|comenzando|comenza|empezar|empiezan|empieza|iniciar)"
    r"\s+con\s+(?:los\s+numeros?\s+|el\s+numero\s+|el\s+)?(\d+)(?:\s+o\s+(\d+))?"
)


def _inicio_de(textos):
    """(prohibido, prefijos): el digito que no puede ir al inicio y/o los prefijos forzados.

    El negativo se extrae y se BORRA del texto antes de buscar el positivo: "no puede
    empezar con 0" tambien matchea el patron de prefijo forzado, y sin esto quedaria
    interpretado al reves."""
    prohibido = ""
    prefijos = []
    for texto in textos:
        t = _normalizar(texto)
        if not t:
            continue
        m = _RX_NO_INICIA.search(t)
        if m:
            prohibido = prohibido or m.group(1)
            t = t[:m.start()] + " " + t[m.end():]
        m = _RX_PREFIJO.search(t)
        if m:
            for g in (m.group(1), m.group(2)):
                if g and g not in prefijos:
                    prefijos.append(g)
    return prohibido, prefijos


def _cuerpo(clase, forma, lo, hi, prohibido):
    """Cuerpo del patron. `prohibido` solo aplica a digitos: "no empezar con 0" no tiene
    lectura razonable sobre una clase alfabetica."""
    fija_primero = bool(prohibido) and clase == _CLASE_DIGITOS
    primero = f"[{int(prohibido) + 1}-9]" if fija_primero and prohibido == "0" else "[1-9]"

    def _tramo(n_lo, n_hi):
        if not fija_primero:
            if n_hi is None:
                return f"{clase}{{{n_lo},}}"
            if n_lo == n_hi:
                return f"{clase}{{{n_lo}}}"
            return f"{clase}{{{n_lo},{n_hi}}}"
        # Con el primer char fijado, el resto del largo baja en 1.
        r_lo, r_hi = max(n_lo - 1, 0), (None if n_hi is None else max(n_hi - 1, 0))
        if r_hi == 0 and r_lo == 0:
            return primero
        if r_hi is None:
            return f"{primero}{clase}{{{r_lo},}}"
        if r_lo == r_hi:
            return f"{primero}{clase}{{{r_lo}}}"
        return f"{primero}{clase}{{{r_lo},{r_hi}}}"

    if forma == "alt":
        return f"(?:{_tramo(lo, lo)}|{_tramo(hi, hi)})"
    if forma == "min":
        return _tramo(lo, None)
    return _tramo(lo, hi)


def _componer(clase, cantidad, prohibido, prefijos):
    forma, lo, hi = cantidad
    # Un prefijo forzado y un primer char restringido se pisan entre si; el prefijo es
    # el dato mas concreto, asi que gana y se descarta el prohibido.
    if prefijos:
        prohibido = ""
    cuerpo = _cuerpo(clase, forma, lo, hi, prohibido)
    if prefijos:
        # Forma ya soportada por el generador (ver `_analizar`: lookahead al inicio).
        return f"^(?=(?:{'|'.join(prefijos)})){cuerpo}$"
    return f"^{cuerpo}$"


def _validar(regex):
    """Devuelve `regex` si compila y el generador puede producir un valor con el; "" si no.

    Es el guardarraíl del modulo: un patron que el generador no sabe satisfacer no sirve
    de nada rio abajo, y es preferible dejar el campo sin regla que cargar una inservible.
    """
    if not regex:
        return ""
    try:
        re.compile(regex)
    except re.error:
        return ""
    if not GeneradorValorCampo(semilla=0).generar_desde_regex(regex, max_variantes=1):
        return ""
    return regex


def derivar_regex(descripcion, mensajes_error=None):
    """Deriva `(regex, motivo)` de la prosa del excel. `regex` es "" si no alcanza.

    `motivo` es el fragmento de prosa que justifico la cantidad — va al reporte de
    auditoria del import para que se pueda revisar de donde salio cada regla.
    """
    mensajes = [m for m in (mensajes_error or []) if _es_util(m)]
    desc = descripcion if _es_util(descripcion) else ""

    # CLASE: la descripcion la califica mejor; el mensaje es el respaldo.
    clase = _clase_de(desc) or next((c for c in (_clase_de(m) for m in mensajes) if c), "")

    # CANTIDAD: gana el mensaje de error (ver docstring del modulo).
    cantidad, motivo = None, ""
    for mensaje in mensajes:
        cantidad = _cantidad_de(mensaje)
        if cantidad:
            motivo = str(mensaje).strip()
            break
    cantidad_desc = _cantidad_de(desc) if desc else None
    if not cantidad and cantidad_desc:
        cantidad = cantidad_desc
        motivo = str(desc).strip()

    if not clase or not cantidad:
        return "", ""

    # El mensaje gana, pero cuando solo fija un piso ("Ingresa al menos 2 letras") y la
    # descripcion si tiene techo ("De 2 a 50 caracteres alfabeticos"), quedarse con el
    # piso suelto emitiria valores mas largos que el maxlength real del campo.
    if cantidad[0] == "min" and cantidad_desc and cantidad_desc[0] == "rango":
        techo = cantidad_desc[2]
        if techo is not None and techo >= cantidad[1]:
            cantidad = ("rango", cantidad[1], techo)

    prohibido, prefijos = _inicio_de([*mensajes, desc])
    regex = _validar(_componer(clase, cantidad, prohibido, prefijos))
    return (regex, motivo) if regex else ("", "")


_RX_CUANTIFICADOR = re.compile(r"\{(\d+)(?:,(\d*))?\}")

_INF = float("inf")


def _rango_prosa(descripcion, mensajes_error):
    """(lo, hi, motivo) del largo que pide la prosa, con hi=inf si es abierto."""
    mensajes = [m for m in (mensajes_error or []) if _es_util(m)]
    desc = descripcion if _es_util(descripcion) else ""

    cantidad, motivo = None, ""
    for mensaje in mensajes:
        cantidad = _cantidad_de(mensaje)
        if cantidad:
            motivo = str(mensaje).strip()
            break
    cantidad_desc = _cantidad_de(desc) if desc else None
    if not cantidad and cantidad_desc:
        cantidad, motivo = cantidad_desc, str(desc).strip()
    if not cantidad:
        return None
    if cantidad[0] == "min" and cantidad_desc and cantidad_desc[0] == "rango":
        techo = cantidad_desc[2]
        if techo is not None and techo >= cantidad[1]:
            cantidad = ("rango", cantidad[1], techo)
    # "alt" (dos largos alternativos) no se puede expresar tocando un solo cuantificador.
    if cantidad[0] == "alt":
        return None
    lo, hi = cantidad[1], cantidad[2]
    return lo, (_INF if hi is None else hi), motivo


def ajustar_largo_desde_prosa(regex_actual, descripcion, mensajes_error=None):
    """Corrige SOLO la cantidad de caracteres de un regex ya cargado. `("", "")` si no toca.

    Un regex escrito a mano suele saber cosas que la prosa del excel no menciona: que un
    comentario admite espacios y puntuacion, que un nombre necesita vocal y consonante, que
    una cedula ecuatoriana arranca con el codigo de provincia. Regenerarlo entero desde la
    prosa perderia todo eso. Lo que la prosa si acierta —y donde el JSON se equivoca— es el
    LARGO, asi que se ajusta unicamente el cuantificador y se preserva el resto intacto.

    Se exige un unico cuantificador: con dos o mas no hay forma de saber a cual corresponde
    el largo que describe la prosa, y con ninguno no hay nada que ajustar. El largo fijo que
    aportan los demas atomos se mide empiricamente pidiendole un valor al generador.
    """
    regex_actual = str(regex_actual or "").strip()
    if not regex_actual:
        return "", ""

    rango = _rango_prosa(descripcion, mensajes_error)
    if not rango:
        return "", ""
    lo_prosa, hi_prosa, motivo = rango

    cuantificadores = list(_RX_CUANTIFICADOR.finditer(regex_actual))
    if len(cuantificadores) != 1:
        return "", ""
    m = cuantificadores[0]
    lo_act = int(m.group(1))
    if m.group(2) is None:                      # {n}
        hi_act = lo_act
    elif m.group(2) == "":                      # {n,}
        hi_act = _INF
    else:                                       # {n,m}
        hi_act = int(m.group(2))

    gen = GeneradorValorCampo(semilla=0)

    # Cuántos caracteres aportan los átomos que NO están bajo el cuantificador: se mide
    # fijándolo en su mínimo y viendo qué largo sale, en vez de re-parsear el patrón.
    sonda = f"{regex_actual[:m.start()]}{{{lo_act}}}{regex_actual[m.end():]}"
    valores = gen.generar_desde_regex(sonda, max_variantes=1)
    if not valores:
        return "", ""
    fijo = len(valores[0]) - lo_act
    if fijo < 0:
        return "", ""

    total_lo, total_hi = lo_act + fijo, hi_act + fijo
    nuevo_lo, nuevo_hi = max(total_lo, lo_prosa), min(total_hi, hi_prosa)
    if nuevo_lo > nuevo_hi:
        # Se contradicen de plano (el JSON pide 15 y el form acepta hasta 10): manda la
        # prosa, que es lo que el form muestra al rechazar.
        nuevo_lo, nuevo_hi = lo_prosa, hi_prosa
    if (nuevo_lo, nuevo_hi) == (total_lo, total_hi):
        return "", ""

    q_lo, q_hi = nuevo_lo - fijo, (_INF if nuevo_hi == _INF else nuevo_hi - fijo)
    if q_lo < 0 or (q_hi != _INF and q_hi < q_lo):
        return "", ""
    if q_hi == _INF:
        cuant = f"{{{q_lo},}}"
    elif q_lo == q_hi:
        cuant = f"{{{q_lo}}}"
    else:
        cuant = f"{{{q_lo},{q_hi}}}"

    candidato = regex_actual[:m.start()] + cuant + regex_actual[m.end():]
    return (candidato, motivo) if _validar(candidato) else ("", "")


# --- fallback semantico (sin prosa: solo el nombre/label del campo) ---

# El orden importa: las claves mas especificas primero, porque "numero"/"number" aparecen
# dentro de nombres como "numero de patente" y se los comerian. Se incluyen las formas en
# ingles porque los element_id reales del repo estan en ingles (contract, patent, vin,
# firstname, telephone...) mientras que los labels vienen en español.
_SEMANTICA = (
    (("mail", "email", "correo"), _REGEX_EMAIL),
    (("comment", "comentario", "mensaje", "message", "opinion", "sugerencia",
      "suggestion", "consulta", "descripcion"),
     r"^[a-zA-ZñÑ .,;:!?()_\-+/0-9]{10,200}$"),
    # El VIN son 17 alfanuméricos por norma ISO 3779, no es una heurística: va antes que
    # patente/placa, que comparten campo pero tienen largo de patente (6-8).
    (("vin", "chasis", "chassi"), r"^[a-zA-Z0-9]{17}$"),
    (("patente", "patent", "placa", "plate"), r"^[a-zA-Z0-9]{6,8}$"),
    (("nombre", "apellido", "firstname", "lastname", "surname", "name"),
     r"^[a-zA-Z]{3,20}$"),
    (("contrato", "contract", "orden", "order", "telefono", "telephone", "celular",
      "phone", "movil", "documento", "document", "cedula", "ci", "dni", "cuit", "cuil",
      "rut", "codigo postal", "zipcode", "numero", "number"),
     r"^[1-9][0-9]{5,9}$"),
)

# Claves tan cortas que como substring matchearian cualquier cosa ("ci" esta dentro de
# "ciudad", "nacimiento", "direccion"): para estas se exige token completo.
_CLAVES_EXACTAS = {"ci", "dni", "rut", "cp", "vin"}


# Campos que las keywords de abajo matchearian por su raiz pero que piden algo muy
# distinto: un "telephone_prefix" no es un telefono, es un prefijo de 1-3 digitos, y
# llenarlo con un numero largo dispara la validacion del form. Sin saber el formato
# exacto, quedarse callado es mejor que inventar: el contrato del modulo es no adivinar.
_SIN_INFERENCIA = ("prefix", "prefijo", "codigo de area", "area code", "lada")


def regex_por_semantica(field_id, label=""):
    """Patron plausible inferido del nombre/label del campo, o "" si ninguno aplica.

    Ultimo recurso, para campos sin regla y sin prosa en el excel: no reemplaza a una
    validacion real, solo evita mandar el form con el campo vacio. Es deliberadamente
    agnostico del mercado — "contrato -> numerico" vale en los nueve.
    """
    texto = _normalizar(f"{field_id or ''} {label or ''}")
    if not texto.strip():
        return ""
    # Los ids reales vienen en kebab/snake/camel ("regret-reason", "first_name"): se
    # trocean para poder exigir token completo en las claves cortas.
    if any(x in texto for x in _SIN_INFERENCIA):
        return ""
    tokens = set(re.split(r"[^a-z0-9]+", re.sub(r"(?<=[a-z])(?=[A-Z])", " ", texto).lower()))
    for claves, regex in _SEMANTICA:
        for clave in claves:
            if clave in _CLAVES_EXACTAS:
                if clave in tokens:
                    return _validar(regex)
            elif clave in texto:
                return _validar(regex)
    return ""
