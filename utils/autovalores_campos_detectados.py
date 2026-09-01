"""Siembra valores de prueba para campos detectados en vivo sin mapeo ni dato en el Excel.

`AutovaloresCamposDetectados` toma los campos que el auto-discovery encontro en el
form, arma hasta 2 valores plausibles por campo (regex de la regla de validacion
para inputs de texto; opciones reales del dropdown para selects) y los persiste en
json/ids_dinamicos.json scopeados al pais. El flujo de llenado ya existente los
levanta por su prioridad "IDs dinamicos".

Reglas duras:
- El Excel siempre gana: si el campo trae valor en el Excel, no se siembra nada.
- Nunca se pisa una entry que el usuario ya cargo para ese id/pais.
- Se preservan `dependencies` y las entries ajenas del store.
"""
import json
import os

from utils.valor_campo_generator import GeneradorValorCampo

__all__ = ["AutovaloresCamposDetectados"]

_TIPOS_DROPDOWN = {"select", "dropdown"}


class AutovaloresCamposDetectados:
    def __init__(self, pais, *, json_dir=None, generador=None):
        self.pais = str(pais or "").strip()
        self.json_dir = json_dir or os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "json"
        )
        self.generador = generador or GeneradorValorCampo()
        self._reglas_cache = None

    # --- reglas de validacion ---

    def _ruta_reglas(self):
        clave = self.pais.lower().replace(" ", "_")
        return os.path.join(self.json_dir, f"field_validation_rules_{clave}.json")

    def _cargar_reglas(self):
        """Mapa {clave_lower -> regla} indexando por KEY, `campo` y `element_id`."""
        if self._reglas_cache is not None:
            return self._reglas_cache

        mapa = {}
        try:
            with open(self._ruta_reglas(), encoding="utf-8") as fh:
                data = json.load(fh)
            for key, regla in (data.get("fields") or {}).items():
                if not isinstance(regla, dict):
                    continue
                for alias in (key, regla.get("campo"), regla.get("element_id")):
                    alias = str(alias or "").strip().lower()
                    if alias:
                        mapa.setdefault(alias, regla)
        except Exception:
            mapa = {}

        self._reglas_cache = mapa
        return mapa

    def regex_para(self, field_id):
        """`regex_full` de la regla que matchea `field_id`, o "" si no hay."""
        regla = self._cargar_reglas().get(str(field_id or "").strip().lower())
        if not regla:
            return ""
        return str(regla.get("regex_full") or "").strip()

    def _es_dropdown(self, field_id, tipo):
        if str(tipo or "").strip().lower() in _TIPOS_DROPDOWN:
            return True
        regla = self._cargar_reglas().get(str(field_id or "").strip().lower())
        return bool(regla and regla.get("dropdown"))

    # --- generacion de valores ---

    def valores_para_campo(self, field_id, tipo, *, opciones_dropdown=None,
                           valor_excel=None, max_variantes=2):
        """Hasta `max_variantes` valores para el campo, o [] si el Excel manda o no hay con que."""
        if valor_excel is not None and str(valor_excel).strip():
            return []

        if self._es_dropdown(field_id, tipo):
            return self.generador.elegir_opciones_dropdown(opciones_dropdown or [], k=max_variantes)

        regex_full = self.regex_para(field_id)
        if not regex_full:
            return []
        return self.generador.generar_desde_regex(regex_full, max_variantes=max_variantes)

    # --- persistencia en ids_dinamicos.json ---

    def _ruta_store(self):
        return os.path.join(self.json_dir, "ids_dinamicos.json")

    def _cargar_store(self):
        try:
            with open(self._ruta_store(), encoding="utf-8") as fh:
                data = json.load(fh)
            if isinstance(data, dict):
                data.setdefault("version", 2)
                data.setdefault("entries", [])
                if not isinstance(data["entries"], list):
                    data["entries"] = []
                return data
        except Exception:
            pass
        return {"version": 2, "entries": []}

    def _campo_ya_tiene_valor(self, entries, field_id):
        """True si ya hay una entry para ese id aplicable al pais actual (global o mismo pais)."""
        for entry in entries:
            if not isinstance(entry, dict) or str(entry.get("id") or "").strip() != field_id:
                continue
            paises = entry.get("paises") or entry.get("countries") or []
            if isinstance(paises, str):
                paises = [paises]
            if not paises or not self.pais or self.pais in paises:
                return True
        return False

    def sembrar(self, campos):
        """Genera y persiste valores para `campos`. Devuelve {field_id: [valores]} sembrados.

        `campos`: iterable de dicts con `id`, `type`, opcional `opciones` (list[str] de un
        dropdown vivo), opcional `valor_excel`, opcional `label`.
        """
        store = self._cargar_store()
        entries = store["entries"]

        sembrados = {}
        for campo in campos or []:
            if not isinstance(campo, dict):
                continue
            field_id = str(campo.get("id") or "").strip()
            if not field_id or field_id in sembrados:
                continue
            if self._campo_ya_tiene_valor(entries, field_id):
                continue

            valores = self.valores_para_campo(
                field_id,
                campo.get("type"),
                opciones_dropdown=campo.get("opciones"),
                valor_excel=campo.get("valor_excel"),
            )
            if not valores:
                continue

            entries.append({
                "id": field_id,
                "valor": valores if len(valores) > 1 else valores[0],
                "paises": [self.pais] if self.pais else [],
                # label puede venir como " " desde el DOM: si queda vacío al limpiar, usar el id.
                "nombre_campo": (str(campo.get("label") or "").strip() or field_id),
                "origen": "autovalor",
            })
            sembrados[field_id] = valores

        if sembrados:
            self._guardar_store(store)
        return sembrados

    def _guardar_store(self, store):
        os.makedirs(self.json_dir, exist_ok=True)
        ruta = self._ruta_store()
        tmp = ruta + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(store, fh, ensure_ascii=False, indent=2)
        os.replace(tmp, ruta)
