"""Un smoke test por funcionalidad: que el camino principal arranque y deje lo que promete.

Estos tests aíslan la carpeta de configuración con `tmp_path` para no pisar los datos
reales del usuario. Esa es la misma técnica que ocultó el bug de rutas, así que conviene
ser explícito sobre la división de trabajo:

  - `tests/test_rutas.py` verifica la resolución REAL de rutas, sin parchear nada.
  - estos tests verifican la LÓGICA, con las rutas aisladas.

Hacen falta los dos. El bug se coló porque solo existía el segundo tipo.
"""
import json
import os

import pytest


# ============================================================================
# Programación de Tests
# ============================================================================

class TestProgramacionDeTests:
    """Lo que el usuario arma en la pestaña 'Programación de Tests'.

    Esta funcionalidad guardaba en una carpeta que nadie leía: era uno de los 10 sitios
    que resolvían mal la raíz.
    """

    @pytest.fixture
    def json_aislado(self, tmp_path, monkeypatch):
        carpeta = tmp_path / "json"
        carpeta.mkdir()
        monkeypatch.setattr("osocio.utils.scheduling.JSON_DIR", str(carpeta))
        return carpeta

    def test_guardar_y_releer_devuelve_lo_mismo(self, json_aislado):
        from osocio.utils import scheduling

        programacion = {
            "tipo": "semanal",
            "modo_tarea": "leads",
            "horarios": {"lunes": ["09:00"], "miercoles": ["14:30"]},
            "paises": ["Argentina", "Chile"],
            "navegadores": ["chrome"],
            "viewports": ["fullscreen"],
            "dispositivo": "local",
        }

        assert scheduling.guardar_programacion(programacion) is True
        leida = scheduling.cargar_programacion()

        assert leida is not None, "se guardó pero no se pudo releer"
        assert leida["horarios"] == programacion["horarios"]
        assert leida["paises"] == programacion["paises"]

    def test_se_escribe_en_el_json_dir_configurado(self, json_aislado):
        """Que el archivo caiga donde se espera, no en cualquier lado."""
        from osocio.utils import scheduling

        scheduling.guardar_programacion({
            "tipo": "semanal", "horarios": {"lunes": ["08:00"]}, "paises": ["Peru"],
        })
        assert (json_aislado / "programacion_test.json").is_file()

    def test_sin_programacion_guardada_no_revienta(self, json_aislado):
        from osocio.utils import scheduling
        assert scheduling.cargar_programacion() in (None, {}, [])

    def test_limpiar_borra_el_archivo(self, json_aislado):
        from osocio.utils import scheduling

        scheduling.guardar_programacion({
            "tipo": "semanal", "horarios": {"lunes": ["08:00"]}, "paises": ["Peru"],
        })
        scheduling.limpiar_programacion()
        assert not (json_aislado / "programacion_test.json").is_file()


# ============================================================================
# IDs Dinámicos y dependencias entre campos
# ============================================================================

class TestDependenciasEntreCampos:
    """La cadena region -> city -> dealer, que dejó de aplicarse por el bug de rutas."""

    def test_las_dependencias_base_siempre_estan(self):
        from osocio.core.field_dependencies import get_field_dependencies

        deps = get_field_dependencies()
        assert deps["region"] == "city"
        assert deps["city"] == "dealer"

    def test_un_pais_sin_reglas_propias_hereda_las_generales(self):
        from osocio.core.field_dependencies import get_field_dependencies

        deps = get_field_dependencies("PaisQueNoExiste")
        assert deps.get("region") == "city"

    def test_no_revienta_si_el_json_esta_roto(self, tmp_path, monkeypatch):
        """Un ids_dinamicos.json corrupto no puede tumbar una corrida entera."""
        carpeta = tmp_path / "json"
        carpeta.mkdir()
        (carpeta / "ids_dinamicos.json").write_text("{ esto no es json", encoding="utf-8")
        monkeypatch.setattr("osocio.paths.JSON_DIR", str(carpeta))

        from osocio.core.field_dependencies import get_field_dependencies
        deps = get_field_dependencies("Peru")
        assert deps["region"] == "city", "tiene que caer a las dependencias hardcodeadas"


# ============================================================================
# Autovalores para campos detectados
# ============================================================================

class TestAutovalores:
    """Cuando el motor encuentra un campo que no estaba en el Excel, inventa un valor."""

    def test_se_le_puede_inyectar_la_carpeta_de_json(self, tmp_path):
        from osocio.utils.autovalores_campos_detectados import AutovaloresCamposDetectados

        autov = AutovaloresCamposDetectados("Peru", json_dir=str(tmp_path))
        assert autov.json_dir == str(tmp_path)

    def test_por_defecto_usa_el_json_del_proyecto(self):
        from osocio.utils.autovalores_campos_detectados import AutovaloresCamposDetectados
        from osocio import paths

        autov = AutovaloresCamposDetectados("Peru")
        assert os.path.normpath(autov.json_dir) == os.path.normpath(paths.JSON_DIR)


# ============================================================================
# Reglas de validación de campos
# ============================================================================

class TestReglasDeValidacion:
    """Las reglas por país que usa la pestaña 'Validación de Campos'."""

    def _paises_con_reglas(self):
        from osocio import paths
        if not os.path.isdir(paths.JSON_DIR):
            return []
        return [n for n in os.listdir(paths.JSON_DIR)
                if n.startswith("field_validation_rules_") and n.endswith(".json")]

    def test_hay_reglas_cargadas(self):
        assert self._paises_con_reglas(), "no hay ningún field_validation_rules_<pais>.json"

    def test_cada_archivo_de_reglas_tiene_forma_de_diccionario(self):
        from osocio import paths

        for nombre in self._paises_con_reglas():
            with open(os.path.join(paths.JSON_DIR, nombre), "r", encoding="utf-8") as fh:
                datos = json.load(fh)
            assert isinstance(datos, dict), f"{nombre} no es un objeto JSON"


# ============================================================================
# Actualización de drivers
# ============================================================================

class TestDrivers:
    """La app actualiza sola chromedriver/msedgedriver al arrancar."""

    def test_expone_el_punto_de_entrada(self):
        from osocio.utils import driver_updater
        assert callable(driver_updater.ensure_drivers_ready)

    def test_apunta_a_la_carpeta_drivers_del_proyecto(self):
        from osocio.utils import driver_updater
        from osocio import paths

        assert os.path.normpath(driver_updater.DRIVERS_DIR) == os.path.normpath(paths.DRIVERS_DIR)


# ============================================================================
# Ejecutor autónomo
# ============================================================================

class TestEjecutorAutonomo:
    """Corre los tests programados en background, sin que nadie abra la app."""

    def test_expone_sus_puntos_de_entrada(self):
        from osocio import autonomous_runner

        assert callable(autonomous_runner.run_once)
        assert callable(autonomous_runner.main)

    def test_sus_rutas_son_las_del_proyecto(self):
        """Era uno de los 10 sitios rotos: usaba json/ y resultados/ equivocados."""
        from osocio import autonomous_runner
        from osocio import paths

        assert os.path.normpath(autonomous_runner.JSON_DIR) == os.path.normpath(paths.JSON_DIR)
        assert os.path.normpath(autonomous_runner.RESULTS_DIR) == os.path.normpath(paths.RESULTS_DIR)
        assert autonomous_runner.LOG_FILE.startswith(paths.JSON_DIR)
