"""
url_status.py
=============
Estado HTTP real de una URL (landing o formulario), para poder distinguir en el Excel
de resultados por QUÉ falló un lead: si el form nunca cargó porque la URL da 404, si
redirige a otro lado, o si el servidor está caído (503).

Selenium no expone el status code de la navegación, así que se consulta aparte con
requests. Es informativo: si el sitio bloquea el request (403 de un WAF) se reporta el
código igual, con la aclaración de que puede ser bot-protection.

Lo usan los tres runners (escritorio, LambdaTest Mac y LambdaTest Android).
"""

from typing import Dict

# User-Agent de browser real: varios sitios de GM devuelven 403 a un cliente sin UA.
_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
       "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")

_HEADERS = {
    "User-Agent": _UA,
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "es-ES,es;q=0.9,en;q=0.8",
}

# Texto legible por código, para que la celda del Excel se entienda sin buscar el número
_CODE_LABELS = {
    400: "Solicitud inválida",
    401: "No autorizado",
    403: "Prohibido (puede ser bot-protection, no necesariamente rota)",
    404: "NO ENCONTRADA (404)",
    408: "Timeout del servidor",
    410: "Eliminada permanentemente",
    429: "Demasiadas solicitudes",
    500: "Error interno del servidor",
    502: "Bad Gateway",
    503: "SERVICIO NO DISPONIBLE (503)",
    504: "Gateway Timeout",
}


def check_url_status(url: str, timeout: int = 12) -> Dict:
    """Consulta el estado HTTP de una URL.

    Devuelve:
        {
          "code": int|None,     # status final (None si no hubo respuesta)
          "label": str,         # texto para la celda del Excel
          "ok": bool,           # True si respondió < 400
          "final_url": str,     # URL final tras redirects
          "redirected": bool,
        }
    """
    out = {"code": None, "label": "-", "ok": False, "final_url": "", "redirected": False}
    url = (url or "").strip()
    if not url:
        return out

    if not url.lower().startswith(("http://", "https://")):
        out["label"] = f"URL inválida (no empieza con http): {url[:60]}"
        return out

    try:
        import requests
    except ImportError:
        out["label"] = "-"
        return out

    try:
        resp = requests.get(url, headers=_HEADERS, timeout=timeout,
                            allow_redirects=True, stream=True)
        try:
            resp.close()
        except Exception:
            pass
    except Exception as e:
        motivo = type(e).__name__
        detalle = str(e)
        if "Timeout" in motivo or "timed out" in detalle.lower():
            out["label"] = f"SIN RESPUESTA (timeout {timeout}s)"
        elif "NameResolution" in detalle or "getaddrinfo" in detalle:
            out["label"] = "SIN RESPUESTA (dominio no resuelve)"
        elif "SSL" in motivo or "SSL" in detalle:
            out["label"] = "SIN RESPUESTA (error SSL/certificado)"
        else:
            out["label"] = f"SIN RESPUESTA ({motivo})"
        return out

    code = resp.status_code
    out["code"] = code
    out["ok"] = code < 400
    out["final_url"] = resp.url or ""
    out["redirected"] = bool(resp.history)

    base = _CODE_LABELS.get(code, "")
    if code == 200 and not resp.history:
        out["label"] = "200 OK"
        return out

    if resp.history:
        # Cadena de redirects: interesa el primer código y a dónde terminó
        cadena = " → ".join(str(h.status_code) for h in resp.history)
        destino = out["final_url"]
        if code == 200:
            out["label"] = f"{cadena} → 200 REDIRIGE a {destino}"
        else:
            out["label"] = f"{cadena} → {code} {base or ''}".strip() + f" ({destino})"
        return out

    out["label"] = f"{code} {base}".strip() if base else str(code)
    return out


def format_status_pair(landing: Dict, form: Dict) -> str:
    """Resumen corto de ambos estados para el motivo del fallo. '' si los dos están OK."""
    problemas = []
    if landing and landing.get("code") is not None and not landing.get("ok"):
        problemas.append(f"landing {landing.get('label')}")
    elif landing and landing.get("code") is None and landing.get("label") not in ("-", ""):
        problemas.append(f"landing {landing.get('label')}")
    if form and form.get("code") is not None and not form.get("ok"):
        problemas.append(f"form {form.get('label')}")
    elif form and form.get("code") is None and form.get("label") not in ("-", ""):
        problemas.append(f"form {form.get('label')}")
    return " ; ".join(problemas)
