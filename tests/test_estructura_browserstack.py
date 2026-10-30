"""El proveedor BrowserStack, sin red.

Lo que se rompe en silencio acá no tira error al importar: un dispositivo mal escrito o
un bloque bstack:options incompleto recién falla al abrir la sesión, contra el hub. Estos
tests vigilan eso, más el orden de las credenciales y el comando de estado.
"""
import json

import pytest

from osocio.providers.browserstack import bs_driver


@pytest.mark.parametrize("plataforma", ["iphone", "android", "mac"])
def test_las_tres_plataformas_del_proyecto_estan(plataforma):
    assert plataforma in bs_driver.plataformas()


def test_dispositivos_pedidos():
    assert bs_driver.DISPOSITIVOS["iphone"]["bstack"]["deviceName"] == "iPhone 14 Pro Max"
    assert bs_driver.DISPOSITIVOS["android"]["bstack"]["deviceName"].startswith("Samsung Galaxy S")
    assert bs_driver.DISPOSITIVOS["android"]["browserName"] == "chrome"
    assert bs_driver.DISPOSITIVOS["mac"]["browserName"] == "safari"


@pytest.mark.parametrize("plataforma", ["iphone", "android"])
def test_celulares_piden_device_real(plataforma):
    opts = bs_driver.build_bstack_options(plataforma, "build", "sesion")
    assert opts["realMobile"] == "true"
    assert opts["deviceName"] and opts["osVersion"]


def test_bstack_options_lleva_nombres_y_video():
    opts = bs_driver.build_bstack_options("mac", "Mi build", "Mi sesion")
    assert opts["buildName"] == "Mi build"
    assert opts["sessionName"] == "Mi sesion"
    assert opts["video"] == "true"


def test_bstack_options_no_pisa_el_catalogo():
    bs_driver.build_bstack_options("iphone", "b", "s")
    assert "buildName" not in bs_driver.DISPOSITIVOS["iphone"]["bstack"]


def test_plataforma_desconocida_falla_claro():
    with pytest.raises(ValueError, match="no soportada"):
        bs_driver.build_bstack_options("windows", "b", "s")


def test_status_script_es_json_valido():
    script = bs_driver.status_script(False, 'campo "email" con error')
    prefijo = "browserstack_executor: "
    assert script.startswith(prefijo)
    payload = json.loads(script[len(prefijo):])
    assert payload["action"] == "setSessionStatus"
    assert payload["arguments"] == {"status": "failed", "reason": 'campo "email" con error'}


def test_status_script_recorta_el_motivo():
    payload = json.loads(bs_driver.status_script(True, "x" * 1000).split(": ", 1)[1])
    assert len(payload["arguments"]["reason"]) == 255


def test_credenciales_del_entorno_tienen_prioridad(monkeypatch, tmp_path):
    archivo = tmp_path / bs_driver.CREDENTIALS_FILENAME
    archivo.write_text("username=del_archivo\naccess_key=key_archivo\n", encoding="utf-8")
    monkeypatch.setattr(bs_driver, "_OSOCIO_DIR", str(tmp_path))
    monkeypatch.setenv("BROWSERSTACK_USERNAME", "del_entorno")
    monkeypatch.setenv("BROWSERSTACK_ACCESS_KEY", "key_entorno")
    assert bs_driver.load_credentials() == ("del_entorno", "key_entorno")


def test_credenciales_desde_archivo(monkeypatch, tmp_path):
    archivo = tmp_path / bs_driver.CREDENTIALS_FILENAME
    archivo.write_text("# comentario\nusername = yo\naccess_key = abc123\n", encoding="utf-8")
    monkeypatch.delenv("BROWSERSTACK_USERNAME", raising=False)
    monkeypatch.delenv("BROWSERSTACK_ACCESS_KEY", raising=False)
    monkeypatch.setattr(bs_driver, "_OSOCIO_DIR", str(tmp_path))
    monkeypatch.setattr(bs_driver, "_THIS_DIR", str(tmp_path / "no_existe"))
    assert bs_driver.load_credentials() == ("yo", "abc123")


def test_sin_credenciales_explica_como_cargarlas(monkeypatch, tmp_path):
    monkeypatch.delenv("BROWSERSTACK_USERNAME", raising=False)
    monkeypatch.delenv("BROWSERSTACK_ACCESS_KEY", raising=False)
    monkeypatch.setattr(bs_driver, "_OSOCIO_DIR", str(tmp_path))
    monkeypatch.setattr(bs_driver, "_THIS_DIR", str(tmp_path))
    with pytest.raises(FileNotFoundError, match="BROWSERSTACK_USERNAME"):
        bs_driver.load_credentials()


def test_api_caida_devuelve_vacio_sin_romper(monkeypatch):
    import requests

    def _falla(*_a, **_k):
        raise requests.ConnectionError("sin red")
    monkeypatch.setattr(requests, "get", _falla)
    info = bs_driver.fetch_bs_session("abc", "u", "k")
    assert info["video_url"] == ""


def test_iphone_no_usa_switch_to_frame():
    # Safari en iPhone bloquea entrar a un iframe de otro dominio: la rama tiene que ser
    # explícita, no depender de que falle switch_to.frame() en cada lead.
    assert bs_driver.puede_entrar_a_iframe("iphone") is False


@pytest.mark.parametrize("plataforma", ["android", "mac"])
def test_android_y_mac_entran_al_iframe(plataforma):
    assert bs_driver.puede_entrar_a_iframe(plataforma) is True


@pytest.mark.parametrize("plataforma,esperado", [("iphone", "iOS"), ("android", "android"), ("mac", "mac")])
def test_platform_name_correcto(plataforma, esperado):
    # SafariOptions trae "mac" de fábrica: el iPhone tiene que pedirse como iOS.
    assert bs_driver.DISPOSITIVOS[plataforma]["platformName"] == esperado



def test_tunel_local_apagado_no_pide_local():
    opts = bs_driver.build_bstack_options("iphone", "b", "s", local=(False, ""))
    assert "local" not in opts and "localIdentifier" not in opts


def test_tunel_local_prendido_pide_local_y_su_identificador():
    opts = bs_driver.build_bstack_options("mac", "b", "s", local=(True, "vpn-elian"))
    assert opts["local"] == "true" and opts["localIdentifier"] == "vpn-elian"
    assert "localIdentifier" not in bs_driver.build_bstack_options("mac", "b", "s", local=(True, ""))


def test_tunel_local_se_lee_de_config_global(tmp_path, monkeypatch):
    import json
    import osocio.paths
    monkeypatch.setattr(osocio.paths, "JSON_DIR", str(tmp_path))
    assert bs_driver.tunel_local() == (False, "")          # sin config: apagado
    (tmp_path / "config_global.json").write_text(
        json.dumps({"browserstack": {"local": True, "local_identifier": "vpn"}}), encoding="utf-8")
    assert bs_driver.tunel_local() == (True, "vpn")


def test_mensaje_tunel_cerrado_no_muestra_la_key():
    m = bs_driver.mensaje_tunel_cerrado("vpn")
    assert "BrowserStackLocal.exe --key <tu access key> --force-local --local-identifier vpn" in m
