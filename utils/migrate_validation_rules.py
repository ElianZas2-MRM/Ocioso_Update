"""
Script de migración: divide field_validation_rules.json en un archivo por país.
Ejecutar una sola vez: python utils/migrate_validation_rules.py
"""
import json
import os
import sys

# Mapeo de abreviación → nombre normalizado para el nombre de archivo
ABBREV_TO_FILENAME = {
    "ar": "argentina",
    "bo": "bolivia",
    "br": "brasil",
    "ch": "chile",
    "cl": "chile",
    "co": "colombia",
    "ec": "ecuador",
    "py": "paraguay",
    "pe": "peru",
    "uy": "uruguay",
}

# Mapeo de nombre en el JSON (campo "paises") → clave de archivo
PAIS_NOMBRE_TO_KEY = {
    "Argentina": "argentina",
    "Bolivia": "bolivia",
    "Brasil": "brasil",
    "Chile": "chile",
    "Colombia": "colombia",
    "Ecuador": "ecuador",
    "Paraguay": "paraguay",
    "Peru": "peru",
    "Perú": "peru",
    "Uruguay": "uruguay",
}


def _get_base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _atomic_write(path, data):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def migrate():
    base_dir = _get_base_dir()
    json_dir = os.path.join(base_dir, "json")
    src_path = os.path.join(json_dir, "field_validation_rules.json")

    if not os.path.exists(src_path):
        print(f"Archivo fuente no encontrado: {src_path}")
        return

    with open(src_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # Parallel arrays: urls, landing_urls, form_urls
    urls = data.get("urls", [])
    landing_urls = data.get("landing_urls", [])
    form_urls = data.get("form_urls", [])
    fields_all = data.get("fields", {})
    global_settings = {
        k: v for k, v in data.items()
        if k not in {"url", "urls", "landing_urls", "form_urls", "fields", "selected_paises"}
    }

    # Agrupar índices de las listas paralelas por país
    pais_indices: dict[str, list[int]] = {}
    for idx, abbrev in enumerate(urls):
        key = ABBREV_TO_FILENAME.get(abbrev.lower())
        if key:
            pais_indices.setdefault(key, []).append(idx)

    # Separar fields por país (cada field tiene campo "paises")
    pais_fields: dict[str, dict] = {}
    for field_name, field_cfg in fields_all.items():
        paises_field = field_cfg.get("paises", [])
        if not paises_field:
            # Sin filtro de país → incluir en todos
            for key in ABBREV_TO_FILENAME.values():
                pais_fields.setdefault(key, {})[field_name] = field_cfg
        else:
            for pais_str in paises_field:
                key = PAIS_NOMBRE_TO_KEY.get(pais_str)
                if key:
                    pais_fields.setdefault(key, {})[field_name] = field_cfg

    # Escribir un archivo por país
    paises_escritos = set()
    for key in set(list(pais_indices.keys()) + list(pais_fields.keys())):
        indices = pais_indices.get(key, [])
        country_urls = [urls[i] for i in indices]
        country_landing = [landing_urls[i] for i in indices if i < len(landing_urls)]
        country_forms = [form_urls[i] for i in indices if i < len(form_urls)]

        # Nombre capitalizado para selected_paises
        pais_nombre = key.capitalize()
        if key == "colombia":
            pais_nombre = "Colombia"

        country_data = {
            **global_settings,
            "url": country_urls[0] if country_urls else key[:2],
            "urls": country_urls,
            "landing_urls": country_landing,
            "form_urls": country_forms,
            "selected_paises": [pais_nombre],
            "fields": pais_fields.get(key, {}),
        }

        out_path = os.path.join(json_dir, f"field_validation_rules_{key}.json")
        _atomic_write(out_path, country_data)
        print(f"  Escrito: {os.path.basename(out_path)} ({len(country_forms)} URLs, {len(pais_fields.get(key, {}))} campos)")
        paises_escritos.add(key)

    print(f"\nMigración completada: {len(paises_escritos)} archivos generados en {json_dir}")
    print("El archivo original no fue modificado.")


if __name__ == "__main__":
    migrate()
