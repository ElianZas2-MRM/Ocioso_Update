# -*- coding: utf-8 -*-
"""
driver_updater.py — Verifica y actualiza los drivers de navegador al iniciar la app.

Por que existe: el driver (chromedriver/msedgedriver/geckodriver) tiene que coincidir con la
version del navegador instalado, y Chrome/Edge se autoactualizan solos cada pocas semanas.
Hasta ahora los .exe se versionaban en el repo y se reemplazaban a mano, asi que cada salto de
major dejaba TODA corrida fallando -- incluidas las programadas de madrugada, donde no hay
nadie para leer el aviso -- hasta que alguien bajaba el driver nuevo.

Como funciona:
- El chequeo normal es OFFLINE: compara el major del navegador instalado contra el `--version`
  del driver que ya esta en /drivers/. Si coinciden no se toca la red, asi abrir la app no
  depende de internet ni suma latencia.
- Solo se sale a internet cuando hay algo para bajar. La excepcion es geckodriver, que versiona
  aparte de Firefox y no se puede deducir del navegador: ese se consulta como mucho una vez por
  semana (el resultado queda cacheado en json/driver_check.json).
- Solo se gestionan los drivers de los navegadores REALMENTE instalados: si no hay Firefox en la
  PC, no se baja geckodriver.
- Nada de esto puede romper el arranque. Cualquier fallo (sin internet, proxy, permisos, driver
  en uso) se loguea en temporales/runtime.log y la app sigue con el driver que ya tenia.

La verificacion TLS va ACTIVADA. En oficinas con proxy que inspecciona TLS (Netskope/Zscaler)
el certificado lo firma el proxy: run.py ya inyecta `truststore`, que hace que Python use la
misma lista de confianza que Windows. Por eso aca no hace falta ningun verify=False.
"""
import json
import os
import re
import shutil
import subprocess
import zipfile
from datetime import datetime, timedelta

try:
    from utils.paths import DRIVERS_DIR, JSON_DIR
except Exception:  # pragma: no cover - fallback si se ejecuta suelto
    _BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    DRIVERS_DIR = os.path.join(_BASE, "drivers")
    JSON_DIR = os.path.join(_BASE, "json")

try:
    from utils.popup_logger import log_runtime
except Exception:  # pragma: no cover
    def log_runtime(message, level="INFO"):
        print(f"[{level}] {message}")


# CREATE_NO_WINDOW: sin esto, cada `driver.exe --version` abre una consola negra que parpadea
# en la cara del usuario al arrancar la app.
_NO_WINDOW = 0x08000000 if os.name == "nt" else 0

# (connect, read). El read es generoso porque el zip de chromedriver pesa ~10 MB y detras de un
# proxy corporativo puede ir lento; el connect es corto para no colgar el arranque si no hay red.
_HTTP_TIMEOUT = (10, 60)

# geckodriver no se puede deducir de la version de Firefox, hay que preguntarle a Mozilla. Se
# consulta una vez por semana para no pegarle a la red en cada arranque.
_GECKO_CHECK_EVERY_DAYS = 7

_CHECK_STATE_FILE = os.path.join(JSON_DIR, "driver_check.json")

# Estados posibles de cada driver despues del chequeo.
OK = "ok"                    # el driver local sirve
NEEDS_UPDATE = "update"      # esta pero desfasado
MISSING = "missing"          # no esta el archivo
NO_BROWSER = "no_browser"    # el navegador no esta instalado -> no nos interesa
UNKNOWN = "unknown"          # no se pudo determinar (version ilegible, etc.)


def _arch_suffix():
    """win64 salvo que estemos en un Windows de 32 bits."""
    arch = (os.environ.get("PROCESSOR_ARCHITECTURE") or "").upper()
    arch_w6432 = (os.environ.get("PROCESSOR_ARCHITEW6432") or "").upper()
    if arch == "X86" and not arch_w6432:
        return "win32"
    return "win64"


BROWSERS = {
    "chrome": {
        "label": "Google Chrome",
        "driver_name": "chromedriver.exe",
        "app_paths_key": "chrome.exe",
        "beacon_keys": (r"Software\Google\Chrome\BLBeacon",),
        "env_binary": None,
        "common_paths": (
            ("ProgramFiles", r"Google\Chrome\Application\chrome.exe"),
            ("ProgramFiles(x86)", r"Google\Chrome\Application\chrome.exe"),
            ("LOCALAPPDATA", r"Google\Chrome\Application\chrome.exe"),
        ),
    },
    "edge": {
        "label": "Microsoft Edge",
        "driver_name": "msedgedriver.exe",
        "app_paths_key": "msedge.exe",
        "beacon_keys": (r"Software\Microsoft\Edge\BLBeacon",),
        "env_binary": None,
        "common_paths": (
            ("ProgramFiles(x86)", r"Microsoft\Edge\Application\msedge.exe"),
            ("ProgramFiles", r"Microsoft\Edge\Application\msedge.exe"),
        ),
    },
    "firefox": {
        "label": "Mozilla Firefox",
        "driver_name": "geckodriver.exe",
        "app_paths_key": "firefox.exe",
        "beacon_keys": (),
        "env_binary": "FIREFOX_BINARY",
        "common_paths": (
            ("ProgramFiles", r"Mozilla Firefox\firefox.exe"),
            ("ProgramFiles(x86)", r"Mozilla Firefox\firefox.exe"),
        ),
    },
}


class DriverStatus:
    """Resultado del chequeo de un driver. Lo consume tambien la UI de progreso."""

    def __init__(self, key, spec):
        self.key = key
        self.label = spec["label"]
        self.driver_name = spec["driver_name"]
        self.browser_version = None
        self.local_version = None
        self.target_version = None
        self.state = UNKNOWN
        self.detail = ""
        self.updated = False
        self.error = ""

    @property
    def needs_download(self):
        return self.state in (NEEDS_UPDATE, MISSING)

    def __repr__(self):  # pragma: no cover - solo para logs
        return (f"<DriverStatus {self.driver_name} {self.state} "
                f"local={self.local_version} target={self.target_version}>")


def _version_tuple(version):
    parts = []
    for chunk in str(version).split("."):
        try:
            parts.append(int(chunk))
        except ValueError:
            parts.append(0)
    return tuple(parts)


def _major(version):
    if not version:
        return None
    try:
        return int(str(version).split(".")[0])
    except (ValueError, IndexError):
        return None


# ── Deteccion del navegador instalado ─────────────────────────────────────────
def _find_browser_exe(spec):
    """Ruta del .exe del navegador, o None si no esta instalado."""
    env_name = spec.get("env_binary")
    if env_name:
        env_value = os.environ.get(env_name)
        if env_value and os.path.exists(env_value):
            return env_value

    for env_var, relative in spec.get("common_paths", ()):
        base = os.environ.get(env_var, "")
        if not base:
            continue
        candidate = os.path.join(base, relative)
        if os.path.exists(candidate):
            return candidate

    if os.name != "nt":
        return None

    try:
        import winreg

        key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{}".format(
            spec["app_paths_key"]
        )
        hives = (
            (winreg.HKEY_LOCAL_MACHINE, key_path),
            (winreg.HKEY_LOCAL_MACHINE, key_path.replace("SOFTWARE\\", "SOFTWARE\\WOW6432Node\\", 1)),
            (winreg.HKEY_CURRENT_USER, key_path),
        )
        for hive, path in hives:
            try:
                with winreg.OpenKey(hive, path) as key:
                    value, _ = winreg.QueryValueEx(key, None)
                    if value and os.path.exists(value):
                        return value
            except OSError:
                continue
    except Exception:
        pass
    return None


def _read_beacon_version(beacon_keys):
    """Version que Chrome/Edge dejan en el registro (BLBeacon). Es la del navegador que
    realmente va a arrancar, incluso si el .exe todavia no se actualizo en disco."""
    if os.name != "nt" or not beacon_keys:
        return None
    try:
        import winreg
    except Exception:
        return None
    for path in beacon_keys:
        for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
            try:
                with winreg.OpenKey(hive, path) as key:
                    value, _ = winreg.QueryValueEx(key, "version")
                    if value:
                        return str(value)
            except OSError:
                continue
    return None


def _read_file_version(exe_path):
    """Version del .exe leida del recurso VERSIONINFO de Windows."""
    if not exe_path or os.name != "nt":
        return None
    try:
        import win32api

        info = win32api.GetFileVersionInfo(exe_path, "\\")
        ms, ls = info["FileVersionMS"], info["FileVersionLS"]
        return "{}.{}.{}.{}".format(ms >> 16, ms & 0xFFFF, ls >> 16, ls & 0xFFFF)
    except Exception:
        pass
    # Fallback sin pywin32: preguntarle a PowerShell.
    try:
        quoted = exe_path.replace("'", "''")
        out = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command",
             "(Get-Item -LiteralPath '{}').VersionInfo.ProductVersion".format(quoted)],
            capture_output=True, text=True, timeout=20, creationflags=_NO_WINDOW,
        )
        value = (out.stdout or "").strip()
        return value or None
    except Exception:
        return None


def _browser_version(spec):
    version = _read_beacon_version(spec.get("beacon_keys"))
    if version:
        return version
    exe = _find_browser_exe(spec)
    if not exe:
        return None
    return _read_file_version(exe)


# ── Version del driver local ──────────────────────────────────────────────────
def _local_driver_version(driver_path):
    """Corre `driver.exe --version` y devuelve solo el numero."""
    if not os.path.exists(driver_path):
        return None
    try:
        out = subprocess.run(
            [driver_path, "--version"],
            capture_output=True, text=True, timeout=25, creationflags=_NO_WINDOW,
        )
    except Exception as exc:
        log_runtime(f"No se pudo leer la version de {os.path.basename(driver_path)}: {exc}",
                    level="WARN")
        return None
    text = (out.stdout or "") + " " + (out.stderr or "")
    match = re.search(r"(\d+\.\d+(?:\.\d+){0,2})", text)
    return match.group(1) if match else None


# ── Estado cacheado (solo para geckodriver) ───────────────────────────────────
def _load_check_state():
    try:
        with open(_CHECK_STATE_FILE, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _save_check_state(state):
    try:
        os.makedirs(JSON_DIR, exist_ok=True)
        with open(_CHECK_STATE_FILE, "w", encoding="utf-8") as fh:
            json.dump(state, fh, indent=2)
    except Exception as exc:
        log_runtime(f"No se pudo guardar {_CHECK_STATE_FILE}: {exc}", level="WARN")


def _local_driver_version_cached(driver_name, driver_path, state):
    """Igual que _local_driver_version pero cacheando contra mtime+size del .exe.

    Lanzar `driver.exe --version` cuesta cerca de un segundo por driver (son binarios de ~40 MB)
    y eso se pagaba en CADA arranque de la app. El .exe solo cambia cuando lo reemplazamos
    nosotros, asi que si mtime y size son los mismos la version tambien lo es.

    Devuelve (version, cache_actualizado) para que el llamador sepa si tiene que guardar.
    """
    try:
        stat = os.stat(driver_path)
    except OSError:
        return None, False

    cache = state.setdefault("local_versions", {})
    entry = cache.get(driver_name) or {}
    if entry.get("mtime") == int(stat.st_mtime) and entry.get("size") == stat.st_size:
        return entry.get("version"), False

    version = _local_driver_version(driver_path)
    if not version:
        # No cachear el fallo: si se cachea None, el driver queda marcado como faltante para
        # siempre (el mtime no cambia) y se re-descargaria en cada arranque sin volver a
        # preguntarle la version al .exe que ya esta en disco.
        return None, False

    cache[driver_name] = {
        "mtime": int(stat.st_mtime),
        "size": stat.st_size,
        "version": version,
    }
    return version, True


def _checked_recently(state, key, days):
    """Ultima version conocida si el chequeo es reciente; None si hay que volver a preguntar."""
    entry = state.get(key) or {}
    stamp = entry.get("checked_at")
    if not stamp:
        return None
    try:
        when = datetime.fromisoformat(stamp)
    except ValueError:
        return None
    if datetime.now() - when > timedelta(days=days):
        return None
    return entry.get("latest")


# ── Resolucion de la version a descargar ──────────────────────────────────────
def _requests():
    """Import diferido: si el chequeo da todo OK no hace falta cargar requests."""
    import requests
    return requests


def _resolve_chrome_target(browser_version):
    """Chrome for Testing publica un chromedriver por build de Chrome. Se busca el de la build
    instalada; si esa build puntual no esta publicada, se usa la ultima del mismo major (que es
    lo unico que chromedriver exige para arrancar)."""
    requests = _requests()
    url = ("https://googlechromelabs.github.io/chrome-for-testing/"
           "latest-patch-versions-per-build-with-downloads.json")
    resp = requests.get(url, timeout=_HTTP_TIMEOUT)
    resp.raise_for_status()
    builds = (resp.json() or {}).get("builds") or {}
    platform = _arch_suffix()

    def _download_url(entry):
        for item in ((entry.get("downloads") or {}).get("chromedriver") or []):
            if item.get("platform") == platform:
                return item.get("url")
        return None

    build_key = ".".join(str(browser_version).split(".")[:3])
    entry = builds.get(build_key)
    if entry and _download_url(entry):
        return entry.get("version") or build_key, _download_url(entry)

    major = _major(browser_version)
    best = None
    for key, item in builds.items():
        if _major(key) != major:
            continue
        download = _download_url(item)
        if not download:
            continue
        version = item.get("version") or key
        if best is None or _version_tuple(version) > _version_tuple(best[0]):
            best = (version, download)
    if best:
        return best
    raise RuntimeError(f"Chrome for Testing no publica chromedriver para Chrome {browser_version}")


def _resolve_edge_target(browser_version):
    """Edge publica LATEST_RELEASE_<major> como texto plano (en UTF-16)."""
    requests = _requests()
    major = _major(browser_version)
    resp = requests.get(
        f"https://msedgedriver.microsoft.com/LATEST_RELEASE_{major}_WINDOWS",
        timeout=_HTTP_TIMEOUT,
    )
    resp.raise_for_status()
    # El endpoint responde UTF-16 con BOM; requests adivina mal si se confia en resp.text.
    version = resp.content.decode("utf-16", errors="ignore").strip().lstrip("﻿")
    if not re.match(r"^\d+\.\d+", version or ""):
        raise RuntimeError(f"Respuesta inesperada de msedgedriver para Edge {major}: {version!r}")
    arch = "win64" if _arch_suffix() == "win64" else "win32"
    return version, f"https://msedgedriver.microsoft.com/{version}/edgedriver_{arch}.zip"


def _resolve_gecko_target(_browser_version=None):
    """geckodriver versiona aparte de Firefox: se toma el ultimo release de Mozilla.
    Se resuelve siguiendo el redirect de /releases/latest en vez de la API de GitHub, que
    limita a 60 llamadas por hora por IP (y en una oficina salen todos por la misma)."""
    requests = _requests()
    resp = requests.get(
        "https://github.com/mozilla/geckodriver/releases/latest",
        timeout=_HTTP_TIMEOUT, allow_redirects=True,
    )
    resp.raise_for_status()
    match = re.search(r"/releases/tag/v?([0-9]+\.[0-9]+(?:\.[0-9]+)?)", resp.url)
    if not match:
        raise RuntimeError(
            f"No se pudo deducir la ultima version de geckodriver desde {resp.url}"
        )
    version = match.group(1)
    arch = "win64" if _arch_suffix() == "win64" else "win32"
    url = (f"https://github.com/mozilla/geckodriver/releases/download/"
           f"v{version}/geckodriver-v{version}-{arch}.zip")
    return version, url


_RESOLVERS = {
    "chrome": _resolve_chrome_target,
    "edge": _resolve_edge_target,
    "firefox": _resolve_gecko_target,
}


# ── Descarga e instalacion ────────────────────────────────────────────────────
def _download_and_install(status, url, on_progress=None):
    """Baja el zip, saca el .exe y lo deja en /drivers/. La descarga va a una carpeta temporal
    dentro de /drivers/ para que el reemplazo final sea un os.replace() en el mismo volumen
    (atomico): si algo se corta a mitad, el driver viejo sigue intacto."""
    requests = _requests()
    os.makedirs(DRIVERS_DIR, exist_ok=True)
    tmp_dir = os.path.join(DRIVERS_DIR, ".tmp_update")
    shutil.rmtree(tmp_dir, ignore_errors=True)
    os.makedirs(tmp_dir, exist_ok=True)
    zip_path = os.path.join(tmp_dir, f"{status.key}.zip")

    try:
        with requests.get(url, stream=True, timeout=_HTTP_TIMEOUT) as resp:
            resp.raise_for_status()
            total = int(resp.headers.get("Content-Length") or 0)
            done = 0
            with open(zip_path, "wb") as fh:
                for chunk in resp.iter_content(chunk_size=65536):
                    if not chunk:
                        continue
                    fh.write(chunk)
                    done += len(chunk)
                    if on_progress:
                        on_progress(status.key, done, total)

        with zipfile.ZipFile(zip_path) as zf:
            member = next(
                (n for n in zf.namelist()
                 if os.path.basename(n).lower() == status.driver_name.lower()),
                None,
            )
            if not member:
                raise RuntimeError(f"El zip descargado no contiene {status.driver_name}")
            extracted = zf.extract(member, tmp_dir)

        final_path = os.path.join(DRIVERS_DIR, status.driver_name)
        try:
            os.replace(extracted, final_path)
        except PermissionError as exc:
            # Windows no deja pisar un .exe en ejecucion. Pasa si quedo un driver zombie de una
            # corrida anterior; el mensaje tiene que decir eso y no "acceso denegado" pelado.
            raise RuntimeError(
                f"{status.driver_name} esta en uso y no se pudo reemplazar. Cerra las ventanas "
                "de navegador que dejo abiertas una corrida anterior y volve a abrir la app."
            ) from exc
        return final_path
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ── API publica ───────────────────────────────────────────────────────────────
def check_drivers():
    """Chequeo OFFLINE de los tres drivers. No toca la red: compara lo que hay en /drivers/
    contra el navegador instalado."""
    results = []
    state = _load_check_state()
    dirty = False
    for key, spec in BROWSERS.items():
        status = DriverStatus(key, spec)
        driver_path = os.path.join(DRIVERS_DIR, spec["driver_name"])
        status.local_version, changed = _local_driver_version_cached(
            spec["driver_name"], driver_path, state
        )
        dirty = dirty or changed
        status.browser_version = _browser_version(spec)

        if not status.browser_version:
            status.state = NO_BROWSER
            status.detail = "No esta instalado en esta PC"
            results.append(status)
            continue

        if not status.local_version:
            status.state = MISSING
            status.detail = "Falta el driver"
            results.append(status)
            continue

        if key == "firefox":
            # geckodriver no se compara contra la version de Firefox: sigue su propio ciclo.
            # El estado real lo define ensure_drivers_ready consultando el ultimo release.
            status.state = OK
            status.detail = f"geckodriver {status.local_version}"
        elif _major(status.local_version) == _major(status.browser_version):
            status.state = OK
            status.detail = f"Compatible con {spec['label']} {_major(status.browser_version)}"
        else:
            status.state = NEEDS_UPDATE
            status.detail = (f"Driver {_major(status.local_version)} vs "
                             f"{spec['label']} {_major(status.browser_version)}")
        results.append(status)

    if dirty:
        _save_check_state(state)
    return results


def needs_network(statuses=None):
    """True si hay algo para bajar (o para consultar) -> la UI muestra la ventana de progreso.
    Si devuelve False la app abre derecho, sin ventana intermedia y sin tocar la red."""
    statuses = statuses if statuses is not None else check_drivers()
    state = None
    for status in statuses:
        if status.needs_download:
            return True
        if status.key == "firefox" and status.state == OK:
            if state is None:
                state = _load_check_state()
            latest = _checked_recently(state, status.driver_name, _GECKO_CHECK_EVERY_DAYS)
            if latest is None:
                return True
            if _version_tuple(latest) > _version_tuple(status.local_version or "0"):
                return True
    return False


def ensure_drivers_ready(on_status=None, on_progress=None, should_cancel=None):
    """Verifica y, si hace falta, descarga los drivers de los navegadores instalados.

    on_status(status)        -> se llama cada vez que cambia el estado de un driver.
    on_progress(key, n, tot) -> bytes descargados / totales (tot puede ser 0 si el servidor no
                                manda Content-Length).
    should_cancel()          -> si devuelve True se corta antes del proximo driver; lo usa el
                                boton "Omitir" de la UI para que el arranque nunca quede
                                atrapado esperando a la red.

    Devuelve la lista de DriverStatus. Nunca lanza: los errores quedan en status.error.
    """
    statuses = check_drivers()
    state = _load_check_state()

    def _notify(status):
        if on_status:
            try:
                on_status(status)
            except Exception:
                pass

    for status in statuses:
        _notify(status)

    for status in statuses:
        if status.state == NO_BROWSER:
            continue
        if should_cancel and should_cancel():
            log_runtime("Actualizacion de drivers cancelada por el usuario", level="WARN")
            break

        try:
            # geckodriver: hay que preguntar cual es el ultimo release, pero como mucho una vez
            # por semana. Si el cache dice que ya estamos al dia, no se toca la red.
            if status.key == "firefox" and status.state == OK:
                cached = _checked_recently(state, status.driver_name, _GECKO_CHECK_EVERY_DAYS)
                if cached is not None:
                    if _version_tuple(cached) <= _version_tuple(status.local_version or "0"):
                        continue
                    status.state = NEEDS_UPDATE
                    status.target_version = cached
                    status.detail = f"Hay geckodriver {cached}"
                    _notify(status)
                else:
                    status.detail = "Buscando actualizaciones..."
                    _notify(status)
                    latest, _url = _resolve_gecko_target()
                    state[status.driver_name] = {
                        "checked_at": datetime.now().isoformat(timespec="seconds"),
                        "latest": latest,
                    }
                    _save_check_state(state)
                    if _version_tuple(latest) <= _version_tuple(status.local_version or "0"):
                        status.detail = f"geckodriver {status.local_version} al dia"
                        _notify(status)
                        continue
                    status.state = NEEDS_UPDATE
                    status.target_version = latest
                    status.detail = f"Hay geckodriver {latest}"
                    _notify(status)

            if not status.needs_download:
                continue

            status.detail = "Buscando la version correcta..."
            _notify(status)
            target, url = _RESOLVERS[status.key](status.browser_version)
            status.target_version = target

            status.detail = f"Descargando {status.driver_name} {target}..."
            _notify(status)
            _download_and_install(status, url, on_progress=on_progress)

            # El .exe cambio, asi que esto vuelve a correr `--version` y deja el cache al dia
            # para los proximos arranques.
            fresh, _changed = _local_driver_version_cached(
                status.driver_name, os.path.join(DRIVERS_DIR, status.driver_name), state
            )
            status.local_version = fresh or target
            _save_check_state(state)
            status.state = OK
            status.updated = True
            status.detail = f"Actualizado a {status.local_version}"
            log_runtime(
                f"{status.driver_name} actualizado a {status.local_version} "
                f"({status.label} {status.browser_version})",
                level="INFO",
            )
            if status.key == "firefox":
                state[status.driver_name] = {
                    "checked_at": datetime.now().isoformat(timespec="seconds"),
                    "latest": status.local_version,
                }
                _save_check_state(state)
            _notify(status)

        except Exception as exc:
            # Sin internet, proxy que bloquea, driver en uso... nada de esto puede impedir que la
            # app abra: queda registrado y se sigue con el driver que ya estaba.
            status.error = str(exc)
            if status.state == MISSING:
                status.detail = "Falta el driver y no se pudo descargar"
            else:
                status.detail = "No se pudo actualizar (se usa el driver actual)"
            log_runtime(f"No se pudo actualizar {status.driver_name}: {exc}", level="ERROR")
            _notify(status)

    return statuses
