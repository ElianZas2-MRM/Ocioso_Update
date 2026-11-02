"""
mercado_por_url.py
==================
De qué mercado es una URL, para que el generador de Excels elija el mercado solo.

El país aparece en distintos lugares según el sitio:
- en el dominio: www.chevrolet.com.br, www.chevrolet.cl
- en la ruta: secure-developments.com/commonwealth/chile/..., .../cadillac/t1/brasil/...
- dentro del dominio de preview de AEM: www-cadillacbrasil-com-br-preview.prod.acp.gm.com

Una URL sin ninguna de esas señales no tiene mercado (None): no se adivina.
"""

from typing import Dict, List, Optional, Set, Tuple
from urllib.parse import unquote, urlparse

# Terminación del dominio de cada mercado (con los puntos alrededor, ver _dominio).
_DOMINIOS = {
    ".com.ar.": "Argentina", ".com.bo.": "Bolivia", ".com.br.": "Brasil", ".cl.": "Chile",
    ".com.co.": "Colombia", ".com.ec.": "Ecuador", ".com.py.": "Paraguay", ".com.pe.": "Peru",
    ".com.uy.": "Uruguay",
}
# Segmento de la ruta con el nombre del mercado.
_SEGMENTOS = {
    "argentina": "Argentina", "bolivia": "Bolivia", "brasil": "Brasil", "brazil": "Brasil",
    "chile": "Chile", "colombia": "Colombia", "ecuador": "Ecuador", "paraguay": "Paraguay",
    "peru": "Peru", "uruguay": "Uruguay",
}


def _dominio(host: str) -> str:
    """El dominio con guiones como puntos y un punto en cada punta: así ".com.br." se
    encuentra igual en www.chevrolet.com.br que en www-cadillacbrasil-com-br-preview..."""
    return "." + host.lower().replace("-", ".").split(":")[0] + "."


def mercados_de_url(url: str) -> Set[str]:
    """Mercados que señala una URL (normalmente uno; vacío si no dice ninguno)."""
    texto = (url or "").strip()
    if not texto:
        return set()
    if "://" not in texto:
        texto = "https://" + texto
    partes = urlparse(texto)
    encontrados = set()
    dominio = _dominio(partes.netloc)
    for fin, mercado in _DOMINIOS.items():
        if fin in dominio:
            encontrados.add(mercado)
    for seg in unquote(partes.path).lower().split("/"):
        if seg in _SEGMENTOS:
            encontrados.add(_SEGMENTOS[seg])
    return encontrados


def detectar_mercado(lineas: List[str]) -> Tuple[Optional[str], Dict[str, List[str]]]:
    """(mercado, por_mercado) para las URLs pegadas en el generador.

    mercado: el único mercado de todas las URLs que dicen alguno, o None si no hay
    ninguno o si hay más de uno. por_mercado: {mercado: [urls]} para explicarle a la
    persona qué URL es de cuál cuando se mezclan. Las URLs que no dicen mercado (por
    ejemplo un form suelto sin país en la ruta) no cuentan para ninguno.
    """
    por_mercado: Dict[str, List[str]] = {}
    for linea in lineas:
        for mercado in sorted(mercados_de_url(linea)):
            por_mercado.setdefault(mercado, []).append(linea.strip())
    mercado = next(iter(por_mercado)) if len(por_mercado) == 1 else None
    return mercado, por_mercado


def aviso_varios_mercados(por_mercado: Dict[str, List[str]], max_por_mercado: int = 3) -> str:
    """Texto del error cuando las URLs son de más de un mercado."""
    lineas = [f"Las URLs son de {len(por_mercado)} mercados distintos. "
              "Se genera un Excel por mercado a la vez: dejá solo las de uno.\n"]
    for mercado, urls in sorted(por_mercado.items()):
        lineas.append(f"{mercado} ({len(urls)}):")
        lineas += [f"  • {u}" for u in urls[:max_por_mercado]]
        if len(urls) > max_por_mercado:
            lineas.append(f"  … y {len(urls) - max_por_mercado} más")
    return "\n".join(lineas)


def pares_inconsistentes(pares: List[Tuple[str, str]]) -> List[Tuple[int, str, str, str, str]]:
    """Pares (landing, form) donde la landing y el form dicen mercados distintos.

    Devuelve (número de par desde 1, landing, mercado landing, form, mercado form). Si una
    de las dos URLs no dice mercado (ej. un preview sin país), no se puede comparar y no
    cuenta como inconsistencia.
    """
    malos = []
    for n, (landing, form) in enumerate(pares, start=1):
        ml, mf = mercados_de_url(landing), mercados_de_url(form)
        if ml and mf and not (ml & mf):
            malos.append((n, landing, ", ".join(sorted(ml)), form, ", ".join(sorted(mf))))
    return malos


def aviso_pares_inconsistentes(malos, max_pares: int = 5) -> str:
    """Texto del error: en qué par la landing y el form son de mercados distintos."""
    lineas = [f"En {len(malos)} par(es) la landing y el form son de mercados distintos. "
              "Revisá que cada landing vaya con el form de su mercado:\n"]
    for n, landing, ml, form, mf in malos[:max_pares]:
        lineas += [f"Par {n}:", f"  • Landing ({ml}): {landing}", f"  • Form ({mf}): {form}"]
    if len(malos) > max_pares:
        lineas.append(f"… y {len(malos) - max_pares} par(es) más")
    return "\n".join(lineas)
