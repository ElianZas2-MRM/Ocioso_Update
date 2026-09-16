"""Campos que la herramienta descubre sola y persiste en json/ids_dinamicos.json.

Metodos extraidos de BaseFormFiller sin tocarles una linea. Siguen siendo parte de
la misma clase via herencia, asi que pueden llamar a sus hermanos y usar el mismo self.
"""
from selenium.webdriver.common.by import By
import json
import os


class IdsDinamicosMixin:
    """Se usa solo como mixin de BaseFormFiller."""

    def _ids_llenados_por_la_herramienta(self):
        """Ids que esta herramienta llena aunque no estén en el mapping del país.

        Tres fuentes: los campos que el auto-discovery encontró en esta corrida, los que
        quedaron registrados en corridas anteriores (json/nuevos_campos_<pais>.json) y los
        que tienen valor en IDs dinámicos. Se usa para decidir qué input de texto merece
        quedar registrado como evidencia en el Excel de resultado.

        El histórico hace falta: el auto-discovery sólo reporta lo que todavía no conocía,
        así que a la segunda corrida `_campos_nuevos_detectados` viene vacío y sin esto la
        evidencia se perdería justo en los formularios que ya se corrieron antes.
        """
        if getattr(self, "_ids_propios_cache", None) is not None:
            return self._ids_propios_cache

        ids = set()
        for campo in getattr(self, "_campos_nuevos_detectados", []) or []:
            fid = str((campo or {}).get("id") or "").strip()
            if fid:
                ids.add(fid)

        pais = str(self.config.get("pais") or "").strip().lower().replace(" ", "_")
        if pais:
            ruta = os.path.join(self.BASE_DIR, "json", f"nuevos_campos_{pais}.json")
            try:
                with open(ruta, "r", encoding="utf-8") as fh:
                    for campo in (json.load(fh) or {}).get("campos_nuevos", []):
                        fid = str((campo or {}).get("id") or "").strip()
                        if fid:
                            ids.add(fid)
            except Exception:
                pass

        try:
            ids.update(self._cargar_ids_dinamicos().keys())
        except Exception:
            pass

        self._ids_propios_cache = ids
        return ids

    def _cargar_nombres_ids_dinamicos(self):
        """Carga nombres de campo para IDs dinámicos aplicables al país actual."""
        nombres = {}
        pais_actual = str(self.config.get("pais") or "").strip()

        def _normalizar_paises(raw):
            if raw is None:
                return []
            if isinstance(raw, str):
                candidatos = [raw]
            elif isinstance(raw, (list, tuple, set)):
                candidatos = list(raw)
            else:
                return []
            salida = []
            for item in candidatos:
                texto = str(item).strip()
                if texto and texto not in salida:
                    salida.append(texto)
            return salida

        try:
            path = os.path.join(self.BASE_DIR, "json", "ids_dinamicos.json")
            if not os.path.exists(path):
                return nombres

            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)

            if isinstance(data, dict) and isinstance(data.get("entries"), list):
                for entry in data.get("entries", []):
                    if not isinstance(entry, dict):
                        continue
                    entry_id = str(entry.get("id") or "").strip()
                    if not entry_id or entry_id in nombres:
                        continue

                    paises_entry = _normalizar_paises(entry.get("paises", entry.get("countries")))
                    if paises_entry and pais_actual not in paises_entry:
                        continue

                    nombre = str(
                        entry.get("nombre_campo")
                        or entry.get("nombre")
                        or entry.get("campo")
                        or ""
                    ).strip()
                    if nombre:
                        nombres[entry_id] = nombre
            elif isinstance(data, dict):
                # Legacy: {id: valor} o {id: {nombre_campo, valor}}
                for entry_id, raw_value in data.items():
                    if entry_id in {"version", "entries"}:
                        continue
                    entry_id = str(entry_id).strip()
                    if not entry_id or entry_id in nombres:
                        continue

                    if isinstance(raw_value, dict):
                        nombre = str(
                            raw_value.get("nombre_campo")
                            or raw_value.get("nombre")
                            or raw_value.get("campo")
                            or ""
                        ).strip()
                        if nombre:
                            nombres[entry_id] = nombre
        except Exception as e:
            print(f" No se pudo cargar nombres de ids_dinamicos: {e}")

        return nombres

    def _persistir_campos_nuevos(self, nuevos):
        """Agrega los campos nuevos al mapping del país y genera el reporte JSON."""
        if not nuevos:
            return

        pais = str(self.config.get("pais") or "").strip()
        if not pais:
            return

        # Agregar al mapping activo en memoria (para que _auto_fill los reconozca)
        ids_ya_en_mapping = self._get_mapped_select_ids()
        for campo in nuevos:
            if campo["id"] not in ids_ya_en_mapping:
                self.field_mapping.append(campo)

        # Persistir en fixed_field_mappings.json
        try:
            from osocio.utils.fixed_field_mapping_store import (
                save_country_fixed_field_mapping,
                load_effective_country_form_config,
            )
            effective = load_effective_country_form_config(pais)
            existing_mapping = list(effective.get("field_mapping") or [])
            existing_ids = set()
            for entry in existing_mapping:
                fid = entry.get("id")
                if isinstance(fid, list):
                    existing_ids.update(fid)
                elif fid:
                    existing_ids.add(fid)
            campos_a_agregar = [c for c in nuevos if c["id"] not in existing_ids]
            if campos_a_agregar:
                merged = existing_mapping + campos_a_agregar
                save_country_fixed_field_mapping(pais, merged)
                print(f"Auto-mapeo: {len(campos_a_agregar)} campo(s) nuevo(s) guardado(s) para {pais}")
        except Exception as e:
            print(f"No se pudo persistir campos nuevos en fixed_field_mappings: {e}")

        # Agregar campos nuevos al field_validation_rules_{pais}.json para que aparezcan en la UI de validación
        try:
            _pais_key_vr = pais.lower().replace(" ", "_")
            _vr_path = os.path.join(self.BASE_DIR, "json", f"field_validation_rules_{_pais_key_vr}.json")
            if os.path.exists(_vr_path):
                with open(_vr_path, "r", encoding="utf-8") as _f:
                    _vr_data = json.load(_f)
                _vr_fields = _vr_data.get("fields") or {}
                _existing_element_ids = {v.get("element_id") for v in _vr_fields.values() if v.get("element_id")}
                _added_vr = False
                for _c in nuevos:
                    if _c["id"] in _existing_element_ids:
                        continue
                    _key = _c["id"].upper()
                    _vr_fields[_key] = {
                        "descripcion": _c.get("name") or _c["id"],
                        "campo": _c["id"],
                        "element_id": _c["id"],
                        "regex_full": "",
                        "regex_char": "",
                        "test_text": "",
                        "dropdown": _c.get("type") == "select",
                        "dropdown_error_message": "",
                        "dependencies": [],
                        "paises": [pais],
                        "teclado_mobile": False,
                        "rules": {},
                        "error_messages": {},
                        "error_message_patterns": [],
                        "error_config": {},
                        "error_priority": [],
                    }
                    _added_vr = True
                if _added_vr:
                    _vr_data["fields"] = _vr_fields
                    _tmp_vr = _vr_path + ".tmp"
                    with open(_tmp_vr, "w", encoding="utf-8") as _f:
                        json.dump(_vr_data, _f, ensure_ascii=False, indent=2)
                    os.replace(_tmp_vr, _vr_path)
                    print(f"Campos nuevos agregados a {os.path.basename(_vr_path)}")
        except Exception as _e_vr:
            print(f"No se pudo actualizar field_validation_rules: {_e_vr}")

        # Generar reporte json/nuevos_campos_<pais>.json
        try:
            from datetime import datetime as _dt
            pais_key = pais.lower().replace(" ", "_")
            reporte_path = os.path.join(self.BASE_DIR, "json", f"nuevos_campos_{pais_key}.json")
            existing_report = {}
            if os.path.exists(reporte_path):
                try:
                    with open(reporte_path, "r", encoding="utf-8") as _f:
                        existing_report = json.load(_f)
                except Exception:
                    pass
            campos_previos = existing_report.get("campos_nuevos", [])
            ids_previos = {c.get("id") for c in campos_previos}
            campos_realmente_nuevos = [c for c in nuevos if c["id"] not in ids_previos]
            if campos_realmente_nuevos:
                reporte = {
                    "pais": pais,
                    "ultima_deteccion": _dt.now().isoformat(timespec="seconds"),
                    "campos_nuevos": campos_previos + [
                        {
                            "id": c["id"],
                            "type": c["type"],
                            "required": c["required"],
                            "label": c.get("name", c["id"]),
                            "origen": "runtime_discovery",
                        }
                        for c in campos_realmente_nuevos
                    ],
                }
                tmp = reporte_path + ".tmp"
                with open(tmp, "w", encoding="utf-8") as _f:
                    json.dump(reporte, _f, ensure_ascii=False, indent=2)
                os.replace(tmp, reporte_path)
                print(f"Reporte de campos nuevos guardado: {os.path.basename(reporte_path)}")
        except Exception as e:
            print(f"No se pudo escribir reporte de campos nuevos: {e}")

    def _sembrar_valores_campos_detectados(self):
        """Genera valores de prueba para los campos recién auto-descubiertos y los
        persiste en json/ids_dinamicos.json (scopeados al país). Para selects toma
        opciones reales del DOM; para inputs, el regex_full de la regla de validación (y
        si no hay, la forma inferida del nombre del campo). El Excel siempre tiene
        prioridad; un valor ya cargado se respeta salvo que no cumpla su propia validación
        en este mercado, en cuyo caso se reemplaza y se avisa por consola."""
        campos = list(getattr(self, "_campos_nuevos_detectados", []) or [])
        pais = str(self.config.get("pais") or "").strip()
        if not campos or not pais:
            return

        payload = []
        for campo in campos:
            fid = str(campo.get("id") or "").strip()
            if not fid:
                continue
            tipo = str(campo.get("type") or "").strip().lower()
            item = {
                "id": fid,
                "type": tipo,
                "label": campo.get("name") or fid,
                "valor_excel": self._excel_value_for_field_id(fid) or None,
            }
            if tipo == "select":
                item["opciones"] = []
                try:
                    encontrados = (self.driver.find_elements(By.ID, fid)
                                   or self.driver.find_elements(By.NAME, fid))
                    if encontrados:
                        item["opciones"] = [
                            (o.text or "").strip()
                            for o in self._get_valid_select_options(encontrados[0])
                        ]
                except Exception:
                    pass
            payload.append(item)

        from osocio.utils.autovalores_campos_detectados import AutovaloresCamposDetectados

        sembrados, corregidos = AutovaloresCamposDetectados(pais).sembrar(payload)
        if sembrados:
            print(f"Autovalores: sembrados {len(sembrados)} campo(s) en IDs dinámicos "
                  f"({pais}): {', '.join(sembrados)}")
        for fid, valores in corregidos.items():
            print(f"Autovalores: '{fid}' tenía un valor que no cumple su validación en "
                  f"{pais} → reemplazado por {valores[0]} (el anterior queda en valor_previo)")

    def _cargar_ids_dinamicos(self):
        """Carga IDs dinámicos aplicables al país actual: globales + específicos de país."""
        ids_por_id = {}
        pais_actual = str(self.config.get("pais") or "").strip()

        def _agregar_valores(entry_id, raw_value):
            valores = self._resolve_dynamic_id_values(raw_value)
            if not valores:
                return
            if entry_id not in ids_por_id:
                ids_por_id[entry_id] = []
            for val in valores:
                if val not in ids_por_id[entry_id]:
                    ids_por_id[entry_id].append(val)

        def _normalizar_paises(raw):
            if raw is None:
                return []
            if isinstance(raw, str):
                candidatos = [raw]
            elif isinstance(raw, (list, tuple, set)):
                candidatos = list(raw)
            else:
                return []
            salida = []
            for item in candidatos:
                texto = str(item).strip()
                if texto and texto not in salida:
                    salida.append(texto)
            return salida

        try:
            path = os.path.join(self.BASE_DIR, "json", "ids_dinamicos.json")
            if os.path.exists(path):
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)

                if isinstance(data, dict) and isinstance(data.get("entries"), list):
                    for entry in data.get("entries", []):
                        if not isinstance(entry, dict):
                            continue
                        entry_id = str(entry.get("id") or "").strip()
                        if not entry_id:
                            continue

                        paises_entry = _normalizar_paises(entry.get("paises", entry.get("countries")))
                        if paises_entry and pais_actual not in paises_entry:
                            continue

                        raw_value = (
                            entry.get("valor")
                            if "valor" in entry
                            else entry.get("valores", entry.get("value", entry.get("values")))
                        )
                        _agregar_valores(entry_id, raw_value)
                elif isinstance(data, dict):
                    # Legacy: {id: valor}
                    for entry_id, raw_value in data.items():
                        if entry_id in {"version", "entries"}:
                            continue
                        entry_id = str(entry_id).strip()
                        if not entry_id:
                            continue
                        _agregar_valores(entry_id, raw_value)
        except Exception as e:
            print(f" No se pudo cargar ids_dinamicos: {e}")
        return ids_por_id
