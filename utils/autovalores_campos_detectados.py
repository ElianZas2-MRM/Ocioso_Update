"""Siembra valores de prueba para campos detectados en vivo sin mapeo ni dato en el Excel.

`AutovaloresCamposDetectados` toma los campos que el auto-discovery encontro en el
form, arma hasta 2 valores plausibles por campo (regex de la regla de validacion
para inputs de texto; opciones reales del dropdown para selects) y los persiste en
json/ids_dinamicos.json scopeados al pais. El flujo de llenado ya existente los
levanta por su prioridad "IDs dinamicos".

Reglas duras:
- El Excel siempre gana: si el campo trae valor en el Excel, no se siembra nada.
- Una entry ya cargada NO se pisa mientras su valor cumpla el regex del campo para el
  mercado en curso. Si no lo cumple si se reemplaza: un valor invalido no protege nada,
  solo produce leads fallidos (caso real: `contract` = "324" en un form que pide 10
  digitos, que ademas bloqueaba la generacion del valor correcto).
- Se preservan `dependencies` y las entries ajenas del store.

Sobre mercados: el mismo element_id tiene validaciones distintas segun el pais (`ci`,
`telephone`, `patent`, `vin`, `firstname`, `lastname` divergen hoy), asi que toda
validacion resuelve la regla por el pais de esta instancia. Ver `_corregir_entry` para
el caso espinoso: una entry sin `paises` aplica a los nueve mercados a la vez.
"""
import json
import os
import re

from utils.regex_desde_prosa import regex_por_semantica
from utils.valor_campo_generator import GeneradorValorCampo

__all__ = ["AutovaloresCamposDetectados"]

_TIPOS_DROPDOWN = {"select", "dropdown"}

# Duplicado deliberado de utils.crm_excel_importer.PAISES_SOPORTADOS: importarlo de alla
# arrastraria openpyxl a un modulo que corre dentro del flujo de llenado. Si se agrega un
# pais a la app, hay que sumarlo aca tambien.
PAISES_CONOCIDOS = (
    "Argentina", "Bolivia", "Brasil", "Chile", "Colombia",
    "Ecuador", "Paraguay", "Peru", "Uruguay",
)


class AutovaloresCamposDetectados:
    def __init__(self, pais, *, json_dir=None, generador=None):
        self.pais = str(pais or "").strip()
        self.json_dir = json_dir or os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "json"
        )
        self.generador = generador or GeneradorValorCampo()
        self._reglas_cache = {}

    # --- reglas de validacion ---

    def _ruta_reglas(self, pais=None):
        clave = str(pais or self.pais).lower().replace(" ", "_")
        return os.path.join(self.json_dir, f"field_validation_rules_{clave}.json")

    def _cargar_reglas(self, pais=None):
        """`(mapa, ambiguos)` del pais pedido. `mapa` indexa {clave_lower -> regla} por KEY,
        `campo` y `element_id`; `ambiguos` son los alias que varias reglas se disputan.

        Un mismo element_id puede tener varias reglas condicionadas a otro campo: en Peru,
        `ci` vale 8 digitos con document-type=DNI, 11 con RUC y 12 con Pasaporte o Carne.
        Quedarse con la primera (lo que hacia un setdefault suelto) daba por invalido un DNI
        correcto por compararlo contra la regla del RUC. Sin saber que eligio el dropdown
        padre no se puede decidir, asi que esos alias se marcan y no se usan. La KEY del
        campo sigue siendo unica, asi que pedir por ella resuelve igual.

        Cachea por pais: el desdoble de una entry global consulta los nueve mercados.
        """
        pais = str(pais or self.pais)
        if pais in self._reglas_cache:
            return self._reglas_cache[pais]

        mapa, ambiguos = {}, set()
        try:
            with open(self._ruta_reglas(pais), encoding="utf-8") as fh:
                data = json.load(fh)
            for key, regla in (data.get("fields") or {}).items():
                if not isinstance(regla, dict):
                    continue
                for alias in (key, regla.get("campo"), regla.get("element_id")):
                    alias = str(alias or "").strip().lower()
                    if not alias:
                        continue
                    previa = mapa.get(alias)
                    if previa is not None and previa is not regla:
                        ambiguos.add(alias)
                    else:
                        mapa[alias] = regla
        except Exception:
            mapa, ambiguos = {}, set()

        self._reglas_cache[pais] = (mapa, ambiguos)
        return mapa, ambiguos

    def _regla_para(self, field_id, pais=None):
        """La regla que matchea `field_id` sin ambigüedad, o None."""
        clave = str(field_id or "").strip().lower()
        mapa, ambiguos = self._cargar_reglas(pais)
        if not clave or clave in ambiguos:
            return None
        return mapa.get(clave)

    def regex_para(self, field_id, pais=None):
        """`regex_full` de la regla que matchea `field_id` en ese mercado, o "" si no hay.

        "" también cuando varias reglas condicionadas comparten el id: ver `_cargar_reglas`.
        """
        regla = self._regla_para(field_id, pais)
        if not regla:
            return ""
        return str(regla.get("regex_full") or "").strip()

    def _es_dropdown(self, field_id, tipo):
        if str(tipo or "").strip().lower() in _TIPOS_DROPDOWN:
            return True
        regla = self._regla_para(field_id)
        return bool(regla and regla.get("dropdown"))

    # --- generacion de valores ---

    def valores_para_campo(self, field_id, tipo, *, opciones_dropdown=None,
                           valor_excel=None, label=None, max_variantes=2):
        """Hasta `max_variantes` valores para el campo, o [] si el Excel manda o no hay con que."""
        if valor_excel is not None and str(valor_excel).strip():
            return []

        if self._es_dropdown(field_id, tipo):
            return self.generador.elegir_opciones_dropdown(opciones_dropdown or [], k=max_variantes)

        regex_full = self.regex_para(field_id)
        if not regex_full:
            # "No hay regla" y "hay varias y no se cual aplica" se parecen pero piden lo
            # contrario: en el segundo caso el campo SI esta validado, y un valor inferido
            # del nombre no cumpliria ninguna de las reglas en disputa.
            _, ambiguos = self._cargar_reglas()
            if str(field_id or "").strip().lower() in ambiguos:
                return []
            regex_full = regex_por_semantica(field_id, label)
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

    @staticmethod
    def _paises_de(entry):
        paises = entry.get("paises") or entry.get("countries") or []
        if isinstance(paises, str):
            paises = [paises]
        return [str(p).strip() for p in paises if str(p).strip()]

    @staticmethod
    def _valores_de(raw):
        if isinstance(raw, (list, tuple)):
            return [str(v).strip() for v in raw if str(v).strip()]
        texto = str(raw).strip() if raw is not None else ""
        return [texto] if texto else []

    def _entry_aplicable(self, entries, field_id):
        """La entry de ese id que aplica al pais actual (scopeada o global), o None."""
        for entry in entries:
            if not isinstance(entry, dict) or str(entry.get("id") or "").strip() != field_id:
                continue
            paises = self._paises_de(entry)
            if not paises or not self.pais or self.pais in paises:
                return entry
        return None

    def _valor_valido(self, raw, field_id, pais=None):
        """True si el valor cumple el regex del campo en ese mercado.

        Sin regla, sin regex o con un regex que no compila no hay con que invalidar: se da
        por bueno. La duda favorece al valor que ya estaba cargado.
        """
        regex = self.regex_para(field_id, pais=pais)
        if not regex:
            return True
        valores = self._valores_de(raw)
        if not valores:
            return False
        try:
            compilado = re.compile(regex)
        except re.error:
            return True
        return all(compilado.fullmatch(v) for v in valores)

    def _corregir_entry(self, entries, entry, field_id, valores):
        """Reemplaza un valor invalido, sin romper los mercados donde si servia.

        Una entry **scopeada** se corrige in-place. Una entry **global** (sin `paises`) se
        desdobla: al llenar, los valores de todas las entries de un id se ACUMULAN en una
        lista (core/base_form_filler.py, `_cargar_ids_dinamicos`) — no hay precedencia de
        pais sobre global — asi que agregar una entry del pais no alcanzaria: la global
        invalida seguiria en juego. Entonces se restringe la global a los mercados donde su
        valor si valida y el mercado actual se lleva una entry nueva.
        """
        valor = valores if len(valores) > 1 else valores[0]
        paises = self._paises_de(entry)

        if paises:
            entry["valor_previo"] = entry.get("valor")
            entry["valor"] = valor
            entry["origen"] = "autovalor_corregido"
            return

        otros_validos = [
            p for p in PAISES_CONOCIDOS
            if p != self.pais and self._valor_valido(entry.get("valor"), field_id, pais=p)
        ]
        if not otros_validos:
            # No sirve en ningun mercado: se reemplaza y sigue siendo global.
            entry["valor_previo"] = entry.get("valor")
            entry["valor"] = valor
            entry["origen"] = "autovalor_corregido"
            return

        entry["paises"] = otros_validos
        entries.append({
            "id": field_id,
            "valor": valor,
            "paises": [self.pais] if self.pais else [],
            "nombre_campo": entry.get("nombre_campo") or field_id,
            "origen": "autovalor_corregido",
            "valor_previo": entry.get("valor"),
        })

    def sembrar(self, campos):
        """Genera y persiste valores para `campos`. Devuelve `(sembrados, corregidos)`,
        ambos `{field_id: [valores]}`.

        `campos`: iterable de dicts con `id`, `type`, opcional `opciones` (list[str] de un
        dropdown vivo), opcional `valor_excel`, opcional `label`.
        """
        store = self._cargar_store()
        entries = store["entries"]

        sembrados, corregidos = {}, {}
        for campo in campos or []:
            if not isinstance(campo, dict):
                continue
            field_id = str(campo.get("id") or "").strip()
            if not field_id or field_id in sembrados or field_id in corregidos:
                continue

            entry = self._entry_aplicable(entries, field_id)
            if entry is not None and self._valor_valido(entry.get("valor"), field_id):
                continue

            valores = self.valores_para_campo(
                field_id,
                campo.get("type"),
                opciones_dropdown=campo.get("opciones"),
                valor_excel=campo.get("valor_excel"),
                label=campo.get("label"),
            )
            if not valores:
                continue

            if entry is None:
                entries.append({
                    "id": field_id,
                    "valor": valores if len(valores) > 1 else valores[0],
                    "paises": [self.pais] if self.pais else [],
                    # label puede venir como " " desde el DOM: si queda vacío al limpiar, usar el id.
                    "nombre_campo": (str(campo.get("label") or "").strip() or field_id),
                    "origen": "autovalor",
                })
                sembrados[field_id] = valores
            else:
                self._corregir_entry(entries, entry, field_id, valores)
                corregidos[field_id] = valores

        if sembrados or corregidos:
            self._guardar_store(store)
        return sembrados, corregidos

    def _guardar_store(self, store):
        os.makedirs(self.json_dir, exist_ok=True)
        ruta = self._ruta_store()
        tmp = ruta + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(store, fh, ensure_ascii=False, indent=2)
        os.replace(tmp, ruta)
