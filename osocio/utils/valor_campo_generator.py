"""Generador de valores de prueba para campos detectados en vivo sin mapeo ni dato en el Excel.

`GeneradorValorCampo` produce hasta N strings que satisfacen el `regex_full` de la
regla de validacion del campo (json/field_validation_rules_<pais>.json), o elige
opciones reales de un dropdown. El contrato es estricto: todo valor devuelto pasa
`re.fullmatch` contra el patron original; si el patron cae fuera de la gramatica
soportada o no se puede satisfacer, devuelve `[]` y el caller usa su fallback.

No depende de librerias externas de regex-a-string (exrex/rstr): los patrones
reales del repo son regulares y acotados, asi que se parsea una gramatica chica
a mano. El `re.fullmatch` final garantiza correccion aunque el emisor sea
aproximado.
"""
import random
import re
import string

__all__ = ["GeneradorValorCampo"]

_ALNUM = set(string.ascii_letters + string.digits)
_VOCALES_ASCII = set("aeiouAEIOU")


class _NoSoportado(Exception):
    """El patron usa construcciones fuera del alcance del generador."""


def _expandir_clase(cuerpo):
    """Convierte el interior de un `[...]` en la lista concreta de chars admitidos."""
    negada = cuerpo.startswith("^")
    if negada:
        cuerpo = cuerpo[1:]

    chars = set()
    i = 0
    while i < len(cuerpo):
        c = cuerpo[i]
        if c == "\\" and i + 1 < len(cuerpo):
            nxt = cuerpo[i + 1]
            if nxt == "d":
                chars |= set(string.digits)
            elif nxt == "w":
                chars |= _ALNUM | {"_"}
            elif nxt == "s":
                chars |= {" ", "\t"}
            else:
                chars.add(nxt)
            i += 2
            continue
        # Rango tipo a-z / 0-9 (el guion final o inicial se toma literal).
        if i + 2 < len(cuerpo) and cuerpo[i + 1] == "-" and cuerpo[i + 2] != "]":
            lo, hi = c, cuerpo[i + 2]
            if ord(lo) <= ord(hi) <= ord(lo) + 256:
                chars.update(chr(o) for o in range(ord(lo), ord(hi) + 1))
            i += 3
            continue
        chars.add(c)
        i += 1

    if negada:
        chars = set(string.ascii_letters + string.digits) - chars
    return sorted(chars)


def _elegir_char(rnd, permitidos):
    """Prefiere alfanumerico ASCII para que el valor quede legible; cae al resto si no hay."""
    legibles = [c for c in permitidos if c in _ALNUM]
    return rnd.choice(legibles or permitidos)


class _Clase:
    def __init__(self, permitidos):
        self.permitidos = permitidos

    def emitir(self, rnd):
        if not self.permitidos:
            raise _NoSoportado("clase vacia")
        return _elegir_char(rnd, self.permitidos)


class _Literal:
    def __init__(self, ch):
        self.ch = ch

    def emitir(self, rnd):
        return self.ch


class _Grupo:
    def __init__(self, alternativas):
        self.alternativas = alternativas  # list[_Secuencia]

    def emitir(self, rnd):
        return rnd.choice(self.alternativas).emitir(rnd)


class _Repeticion:
    # Cota superior cuando el cuantificador es abierto ({m,} / + / *) o muy amplio:
    # no necesitamos el maximo real, solo un valor que matchee y sea corto.
    _PAD = 8

    def __init__(self, base, lo, hi):
        self.base = base
        self.lo = lo
        self.hi = hi  # None => abierto

    def emitir(self, rnd):
        hi = self.hi if self.hi is not None else self.lo + self._PAD
        techo = min(hi, self.lo + self._PAD)
        techo = max(techo, self.lo)
        cuenta = rnd.randint(self.lo, techo)
        return "".join(self.base.emitir(rnd) for _ in range(cuenta))


class _Secuencia:
    def __init__(self, nodos):
        self.nodos = nodos

    def emitir(self, rnd):
        return "".join(n.emitir(rnd) for n in self.nodos)

    def clase_unica(self):
        """Si la secuencia es una sola clase (con o sin repeticion), devuelve sus chars."""
        if len(self.nodos) != 1:
            return None
        nodo = self.nodos[0]
        if isinstance(nodo, _Repeticion):
            nodo = nodo.base
        if isinstance(nodo, _Clase):
            return nodo.permitidos
        return None

    # --- parser recursivo descendente sobre la gramatica acotada ---

    @classmethod
    def parsear(cls, patron):
        nodos, i = cls._parsear_seq(patron, 0)
        if i != len(patron):
            raise _NoSoportado(f"resto sin parsear en pos {i}: {patron[i:]!r}")
        return cls(nodos)

    @staticmethod
    def _parsear_seq(s, i):
        nodos = []
        while i < len(s) and s[i] not in "|)":
            base, i = _Secuencia._parsear_atomo(s, i)
            lo, hi, i = _Secuencia._parsear_cuant(s, i)
            if (lo, hi) == (1, 1):
                nodos.append(base)
            else:
                nodos.append(_Repeticion(base, lo, hi))
        return nodos, i

    @staticmethod
    def _parsear_atomo(s, i):
        c = s[i]
        if c == "(":
            j = i + 1
            if s[j:j + 2] == "?:":
                j += 2
            elif s[j:j + 1] == "?":
                raise _NoSoportado("grupo especial no soportado")
            alternativas = []
            nodos, j = _Secuencia._parsear_seq(s, j)
            alternativas.append(_Secuencia(nodos))
            while j < len(s) and s[j] == "|":
                nodos, j = _Secuencia._parsear_seq(s, j + 1)
                alternativas.append(_Secuencia(nodos))
            if j >= len(s) or s[j] != ")":
                raise _NoSoportado("parentesis sin cerrar")
            return _Grupo(alternativas), j + 1
        if c == "[":
            j = i + 1
            buf = []
            while j < len(s) and s[j] != "]":
                if s[j] == "\\" and j + 1 < len(s):
                    buf.append(s[j:j + 2])
                    j += 2
                    continue
                buf.append(s[j])
                j += 1
            if j >= len(s):
                raise _NoSoportado("clase sin cerrar")
            return _Clase(_expandir_clase("".join(buf))), j + 1
        if c == "\\":
            if i + 1 >= len(s):
                raise _NoSoportado("escape colgante")
            nxt = s[i + 1]
            if nxt.isdigit():
                raise _NoSoportado("backreference")
            if nxt == "d":
                return _Clase(list(string.digits)), i + 2
            if nxt == "w":
                return _Clase(sorted(_ALNUM | {"_"})), i + 2
            if nxt == "s":
                return _Clase([" "]), i + 2
            if nxt in "bBAZ":
                raise _NoSoportado("ancla de palabra")
            return _Literal(nxt), i + 2
        if c == ".":
            return _Clase(sorted(_ALNUM)), i + 1
        if c in "*+?{":
            raise _NoSoportado(f"cuantificador sin atomo en pos {i}")
        return _Literal(c), i + 1

    @staticmethod
    def _parsear_cuant(s, i):
        if i >= len(s):
            return 1, 1, i
        c = s[i]
        if c == "+":
            return 1, None, i + 1
        if c == "*":
            return 0, None, i + 1
        if c == "?":
            return 0, 1, i + 1
        if c == "{":
            cierre = s.find("}", i)
            if cierre == -1:
                raise _NoSoportado("llave sin cerrar")
            cuerpo = s[i + 1:cierre]
            if "," in cuerpo:
                lo_txt, hi_txt = cuerpo.split(",", 1)
                lo = int(lo_txt) if lo_txt.strip() else 0
                hi = int(hi_txt) if hi_txt.strip() else None
            else:
                lo = hi = int(cuerpo)
            return lo, hi, cierre + 1
        return 1, 1, i


class GeneradorValorCampo:
    """Arma valores de prueba a partir de un `regex_full` o de las opciones de un dropdown."""

    def __init__(self, semilla=None):
        self._rnd = random.Random(semilla)

    def generar_desde_regex(self, regex_full, max_variantes=2):
        """Hasta `max_variantes` strings distintos que cumplen `re.fullmatch(regex_full, x)`.

        Devuelve `[]` si el patron esta vacio, no compila, usa construcciones fuera
        de alcance, o no se pudo satisfacer en los intentos disponibles.
        """
        if not isinstance(regex_full, str):
            return []
        patron = regex_full.strip()
        if not patron or max_variantes < 1:
            return []
        try:
            compilado = re.compile(patron)
        except re.error:
            return []
        try:
            return self._generar(patron, compilado, max_variantes)[:max_variantes]
        except _NoSoportado:
            return []
        except Exception:  # noqa: BLE001 - el contrato es no propagar nunca
            return []

    def elegir_opciones_dropdown(self, opciones, k=2):
        """Devuelve hasta `k` opciones distintas al azar, descartando vacios y duplicados."""
        limpias = []
        for opt in opciones or []:
            texto = str(opt).strip() if opt is not None else ""
            if texto and texto not in limpias:
                limpias.append(texto)
        if not limpias:
            return []
        k = max(1, min(k, len(limpias)))
        return self._rnd.sample(limpias, k)

    # --- interno ---

    def _generar(self, patron, compilado, max_variantes):
        # Un `@` dentro de una clase `[...]` (campo de texto largo que lo admite) no
        # convierte al patron en un email: solo cuenta si esta fuera de toda clase.
        if "@" in re.sub(r"\[(?:[^\]\\]|\\.)*\]", "", patron):
            return self._generar_email(compilado, max_variantes)

        prefijo, debe_contener, cuerpo = self._analizar(patron)
        arbol = _Secuencia.parsear(cuerpo)

        vistos = []
        intentos = max(80, max_variantes * 40)
        for _ in range(intentos):
            candidato = arbol.emitir(self._rnd)
            candidato = self._inyectar_requeridos(candidato, debe_contener, arbol, len(prefijo))
            if prefijo:
                candidato = prefijo + candidato[len(prefijo):]
            if candidato and compilado.fullmatch(candidato) and candidato not in vistos:
                vistos.append(candidato)
                if len(vistos) >= max_variantes:
                    break
        return vistos

    def _generar_email(self, compilado, max_variantes):
        # Dominio de prueba pedido por el usuario: los leads de verificacion usan @mrm.com.
        bases = [f"qa.prueba{n}@mrm.com" for n in range(1, 12)] + \
                [f"demo.osocio{n}@mrm.com" for n in range(1, 12)]
        out = []
        for cand in bases:
            if compilado.fullmatch(cand) and cand not in out:
                out.append(cand)
            if len(out) >= max_variantes:
                break
        return out

    def _analizar(self, patron):
        """Separa lookaheads del frente y devuelve (prefijo_forzado, [clases_requeridas], cuerpo)."""
        s = patron
        if s.startswith("^"):
            s = s[1:]
        if s.endswith("$"):
            s = s[:-1]

        prefijos = []
        debe = []
        while True:
            m = re.match(r"\(\?!\.\*\(\.\)\(\\1\)\{2\}\)", s)  # sin char triplicado
            if m:
                s = s[m.end():]
                continue
            m = re.match(r"\(\?!\(\.\)\(\\1\+\$\)\)", s)  # no todos iguales
            if m:
                s = s[m.end():]
                continue
            m = re.match(r"\(\?=\.\*(\[(?:[^\]\\]|\\.)*\])\)", s)  # debe contener [clase]
            if m:
                debe.append(m.group(1))
                s = s[m.end():]
                continue
            m = re.match(r"\(\?=\(\?:([^()]*)\)\)", s)  # prefijo forzado por alternacion literal
            if m:
                prefijos = [tok for tok in m.group(1).split("|") if tok]
                s = s[m.end():]
                continue
            break

        if re.search(r"\(\?[!=]", s):
            # Queda un lookahead en el medio: fuera de alcance.
            raise _NoSoportado("lookahead no posicionado al inicio")

        prefijo = self._rnd.choice(prefijos) if prefijos else ""
        return prefijo, debe, s

    def _inyectar_requeridos(self, candidato, debe_contener, arbol, offset_prefijo):
        """Si el cuerpo es una sola clase, fuerza que aparezca un char de cada clase requerida."""
        if not debe_contener:
            return candidato
        permitidos_cuerpo = arbol.clase_unica()
        if not permitidos_cuerpo or not candidato:
            return candidato

        chars = list(candidato)
        for clase_txt in debe_contener:
            requeridos = set(_expandir_clase(clase_txt[1:-1]))
            if any(c in requeridos for c in chars):
                continue
            opciones = [c for c in permitidos_cuerpo if c in requeridos]
            legibles = [c for c in opciones if c in _ALNUM]
            opciones = legibles or opciones
            if not opciones:
                continue
            posiciones = list(range(offset_prefijo, len(chars))) or list(range(len(chars)))
            chars[self._rnd.choice(posiciones)] = self._rnd.choice(opciones)
        return "".join(chars)
