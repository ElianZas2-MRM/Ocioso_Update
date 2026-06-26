import json
import os
import re
import sys

from utils.paths import BASE_DIR, BUNDLE_DIR

CORE_DIR = os.path.join(BUNDLE_DIR, "core")
JSON_DIR = os.path.join(BASE_DIR, "json")
FIXED_FIELD_MAPPINGS_PATH = os.path.join(JSON_DIR, "fixed_field_mappings.json")

def _normalize_country_name(country_name):
    return str(country_name or "").strip()


def _normalize_field_id(raw_id):
    if isinstance(raw_id, list):
        normalized_ids = [str(item).strip() for item in raw_id if str(item).strip()]
        return normalized_ids or [""]
    return str(raw_id or "").strip()


def _normalize_field_type(raw_type):
    normalized = str(raw_type or "text").strip().lower()
    return normalized if normalized in {"text", "select"} else "text"


def _normalize_data_index(raw_index):
    if isinstance(raw_index, bool):
        raise ValueError("data_index inválido")
    try:
        value = int(raw_index)
    except Exception as exc:
        raise ValueError("data_index inválido") from exc
    if value < 0:
        raise ValueError("data_index debe ser mayor o igual a 0")
    return value


def _normalize_requested_data_index(raw_index, fallback_index):
    """Normaliza el índice solicitado por usuario para ordenar/reindexar luego."""
    if raw_index is None:
        return fallback_index
    return _normalize_data_index(raw_index)


def _normalize_field_entry(entry):
    if not isinstance(entry, dict):
        raise ValueError("Cada entry de field_mapping debe ser un dict")

    field_id = _normalize_field_id(entry.get("id"))
    field_name = str(entry.get("name") or "").strip()
    field_type = _normalize_field_type(entry.get("type"))
    data_index = _normalize_data_index(entry.get("data_index", 0))
    requested_data_index = _normalize_requested_data_index(entry.get("requested_data_index"), data_index)

    if isinstance(field_id, list):
        if not all(field_id):
            raise ValueError("Los IDs alternativos no pueden estar vacíos")
    elif not field_id:
        raise ValueError("El ID del campo es obligatorio")

    normalized = {
        "type": field_type,
        "id": field_id,
        "data_index": data_index,
        "requested_data_index": requested_data_index,
        "name": field_name or (field_id[0] if isinstance(field_id, list) else field_id),
    }

    if entry.get("data_key"):
        normalized["data_key"] = str(entry.get("data_key")).strip()

    return normalized


def _normalize_required_fields(required_fields):
    if not isinstance(required_fields, list):
        return []
    normalized = []
    for field_id in required_fields:
        text = str(field_id or "").strip()
        if text and text not in normalized:
            normalized.append(text)
    return normalized


def load_base_country_form_configs():
    # Lee configs desde country_configs.py en lugar de parsear archivos Base por AST
    try:
        if CORE_DIR not in sys.path:
            sys.path.insert(0, CORE_DIR)
        from country_configs import COUNTRY_CONFIGS
    except ImportError:
        return {}

    result = {}
    for country_name, raw in COUNTRY_CONFIGS.items():
        normalized = _normalize_country_name(country_name)
        if not normalized:
            continue

        field_mapping = []
        for entry in raw.get("field_mapping") or []:
            try:
                field_mapping.append(_normalize_field_entry(entry))
            except Exception:
                continue

        country_fields = dict(raw.get("country_fields") or {})
        country_fields["required_fields"] = _normalize_required_fields(country_fields.get("required_fields"))

        result[normalized] = {
            "pais": normalized,
            "excel_file": str(raw.get("excel_file") or "").strip(),
            "data_start_index": int(raw.get("data_start_index", 2)),
            "field_mapping": field_mapping,
            "country_fields": country_fields,
        }

    return result


def _normalize_override_country_config(country_name, raw_config):
    if not isinstance(raw_config, dict):
        return None

    field_mapping = []
    for entry in raw_config.get("field_mapping") or []:
        field_mapping.append(_normalize_field_entry(entry))

    required_fields = _normalize_required_fields(raw_config.get("required_fields"))
    return {
        "pais": _normalize_country_name(country_name),
        "field_mapping": field_mapping,
        "required_fields": required_fields,
    }


def load_fixed_field_mapping_overrides():
    if not os.path.exists(FIXED_FIELD_MAPPINGS_PATH):
        return {"version": 1, "countries": {}}

    try:
        with open(FIXED_FIELD_MAPPINGS_PATH, "r", encoding="utf-8") as file_handle:
            data = json.load(file_handle)
    except Exception:
        return {"version": 1, "countries": {}}

    countries = {}
    for country_name, raw_config in dict(data.get("countries") or {}).items():
        normalized = _normalize_override_country_config(country_name, raw_config)
        if normalized is not None:
            countries[_normalize_country_name(country_name)] = normalized

    return {
        "version": 1,
        "countries": countries,
    }


def save_fixed_field_mapping_overrides(data):
    if not isinstance(data, dict):
        raise ValueError("data debe ser dict")

    normalized_countries = {}
    for country_name, raw_config in dict(data.get("countries") or {}).items():
        normalized = _normalize_override_country_config(country_name, raw_config)
        if normalized is not None:
            normalized_countries[_normalize_country_name(country_name)] = normalized

    os.makedirs(JSON_DIR, exist_ok=True)
    tmp_path = FIXED_FIELD_MAPPINGS_PATH + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as file_handle:
        json.dump({"version": 1, "countries": normalized_countries}, file_handle, indent=2, ensure_ascii=False)
    os.replace(tmp_path, FIXED_FIELD_MAPPINGS_PATH)

    return True


def load_effective_country_form_config(country_name, fallback_config=None):
    normalized_country = _normalize_country_name(country_name)
    base_config = load_base_country_form_configs().get(normalized_country, {})

    if not base_config and isinstance(fallback_config, dict):
        country_fields = dict(fallback_config.get("country_fields") or {})
        base_config = {
            "pais": normalized_country,
            "excel_file": str(fallback_config.get("excel_file") or "").strip(),
            "data_start_index": int(fallback_config.get("data_start_index", 2)),
            "field_mapping": [_normalize_field_entry(entry) for entry in (fallback_config.get("field_mapping") or [])],
            "country_fields": {
                **country_fields,
                "required_fields": _normalize_required_fields(country_fields.get("required_fields")),
            },
        }

    effective_config = {
        "pais": normalized_country,
        "excel_file": base_config.get("excel_file", ""),
        "data_start_index": base_config.get("data_start_index", 2),
        "field_mapping": list(base_config.get("field_mapping") or []),
        "country_fields": dict(base_config.get("country_fields") or {}),
    }

    override_config = load_fixed_field_mapping_overrides().get("countries", {}).get(normalized_country)
    if override_config:
        effective_config["field_mapping"] = list(override_config.get("field_mapping") or effective_config["field_mapping"])
        if override_config.get("required_fields"):
            effective_config["country_fields"]["required_fields"] = list(override_config.get("required_fields") or [])

    if "required_fields" not in effective_config["country_fields"]:
        effective_config["country_fields"]["required_fields"] = []

    return effective_config


def list_available_fixed_mapping_countries():
    return sorted(load_base_country_form_configs().keys())


def get_effective_fixed_field_mapping(country_name):
    return list(load_effective_country_form_config(country_name).get("field_mapping") or [])


def save_country_fixed_field_mapping(country_name, field_mapping, required_fields=None):
    normalized_country = _normalize_country_name(country_name)
    if not normalized_country:
        raise ValueError("El país es obligatorio")

    normalized_mapping = [_normalize_field_entry(entry) for entry in (field_mapping or [])]
    normalized_required = _normalize_required_fields(required_fields)

    if not normalized_required:
        base_required = load_effective_country_form_config(normalized_country).get("country_fields", {}).get("required_fields") or []
        base_required_set = set(base_required)
        normalized_required = []
        for entry in normalized_mapping:
            field_id = entry.get("id")
            candidate_ids = field_id if isinstance(field_id, list) else [field_id]
            if any(candidate in base_required_set for candidate in candidate_ids):
                selected_id = candidate_ids[0]
                if selected_id and selected_id not in normalized_required:
                    normalized_required.append(selected_id)

    data = load_fixed_field_mapping_overrides()
    countries = dict(data.get("countries") or {})
    countries[normalized_country] = {
        "pais": normalized_country,
        "field_mapping": normalized_mapping,
        "required_fields": normalized_required,
    }
    data["countries"] = countries
    save_fixed_field_mapping_overrides(data)
    return True


def infer_country_from_excel_filename(file_name):
    normalized_name = str(file_name or "").strip().lower()
    if not normalized_name:
        return ""

    match = re.search(r"formulario_([a-záéíóúñ]+)_main", normalized_name)
    if not match:
        return ""

    guessed = match.group(1)
    for country_name in list_available_fixed_mapping_countries():
        if country_name.lower() == guessed:
            return country_name
    return ""


def build_excel_columns_for_country(country_name):
    effective_config = load_effective_country_form_config(country_name)
    field_mapping = list(effective_config.get("field_mapping") or [])
    if not field_mapping:
        return ["URL", "Formulario"]

    normalized_entries = []
    for pos, entry in enumerate(field_mapping):
        header_name = str(entry.get("name") or "").strip()
        if not header_name:
            field_id = entry.get("id")
            if isinstance(field_id, list):
                header_name = str(field_id[0] if field_id else "").strip()
            else:
                header_name = str(field_id or "").strip()
        if not header_name:
            continue

        requested_idx = entry.get("requested_data_index")
        data_idx = entry.get("data_index")

        if not isinstance(requested_idx, int) or requested_idx < 0:
            requested_idx = data_idx if isinstance(data_idx, int) and data_idx >= 0 else pos
        if not isinstance(data_idx, int) or data_idx < 0:
            data_idx = requested_idx

        normalized_entries.append((requested_idx, data_idx, pos, header_name))

    if not normalized_entries:
        return ["URL", "Formulario"]

    normalized_entries.sort(key=lambda item: (item[0], item[1], item[2]))
    headers = [item[3] for item in normalized_entries]
    return ["URL", "Formulario", *headers]


def get_fixed_mapping_ids(country_name=None):
    if country_name:
        country_names = [_normalize_country_name(country_name)]
    else:
        country_names = list_available_fixed_mapping_countries()

    all_ids = set()
    for current_country in country_names:
        for entry in get_effective_fixed_field_mapping(current_country):
            field_id = entry.get("id")
            candidate_ids = field_id if isinstance(field_id, list) else [field_id]
            for candidate in candidate_ids:
                normalized_id = str(candidate or "").strip()
                if normalized_id:
                    all_ids.add(normalized_id)
    return all_ids