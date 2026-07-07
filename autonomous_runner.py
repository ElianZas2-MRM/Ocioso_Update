"""
autonomous_runner.py — Ejecutor programado en background (sin interfaz gráfica).
Revisa cada 60 segundos si hay tests programados y los lanza automáticamente.
Usa un mutex de Windows para evitar que se corran dos instancias al mismo tiempo.
"""
import glob
import os
import subprocess
import sys
import time
import atexit
import ctypes
import threading
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

from interface.helpers_interface import enviar_email_resultados_consolidados
from utils.scheduling import cargar_programacion, limpiar_programacion


if getattr(sys, 'frozen', False):
    PROJECT_ROOT = os.path.dirname(sys.executable)
else:
    PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))

JSON_DIR = os.path.join(PROJECT_ROOT, "json")
RESULTS_DIR = os.path.join(PROJECT_ROOT, "resultados")
LOG_FILE = os.path.join(JSON_DIR, "ejecutor_autonomo.log")
AUTONOMOUS_MUTEX_NAME = "Global\\OsocioFormAutomationAutonomous"
ERROR_ALREADY_EXISTS = 183
DETACHED_PROCESS = 0x00000008
CREATE_NEW_PROCESS_GROUP = 0x00000200

_mutex_handle = None


def _safe_console_print(message):
    """Imprime sin fallar cuando la consola no soporta algunos caracteres."""
    stream = sys.stdout
    encoding = getattr(stream, "encoding", None) or "utf-8"

    try:
        print(message)
        return
    except UnicodeEncodeError:
        pass

    safe_message = message.encode(encoding, errors="replace").decode(encoding, errors="replace")
    stream.write(safe_message + "\n")


def _close_mutex_handle():
    global _mutex_handle
    if _mutex_handle and os.name == "nt":
        ctypes.windll.kernel32.CloseHandle(_mutex_handle)
    _mutex_handle = None


def acquire_autonomous_lock():
    """Adquiere un lock global para evitar múltiples ejecutores autónomos."""
    global _mutex_handle

    if os.name != "nt":
        return True

    if _mutex_handle:
        return True

    handle = ctypes.windll.kernel32.CreateMutexW(None, False, AUTONOMOUS_MUTEX_NAME)
    if not handle:
        return True

    if ctypes.windll.kernel32.GetLastError() == ERROR_ALREADY_EXISTS:
        ctypes.windll.kernel32.CloseHandle(handle)
        return False

    _mutex_handle = handle
    atexit.register(_close_mutex_handle)
    return True


def is_autonomous_running():
    """Devuelve True si ya existe un ejecutor autónomo activo."""
    if os.name != "nt":
        return False

    handle = ctypes.windll.kernel32.OpenMutexW(0x00100000, False, AUTONOMOUS_MUTEX_NAME)
    if not handle:
        return False

    ctypes.windll.kernel32.CloseHandle(handle)
    return True


def build_autonomous_command():
    """Construye el comando para lanzar el ejecutor autónomo."""
    if getattr(sys, 'frozen', False):
        return [sys.executable, "--autonomous"]
    return [sys.executable, os.path.join(PROJECT_ROOT, "run.py"), "--autonomous"]


def _build_subprocess_env():
    """Aísla relanzamientos del .exe empaquetado para que no reutilicen el _MEI del padre."""
    env = os.environ.copy()
    if getattr(sys, 'frozen', False):
        env["PYINSTALLER_RESET_ENVIRONMENT"] = "1"
    return env


def start_autonomous_background():
    """Lanza el ejecutor autónomo en segundo plano si no hay uno activo."""
    if is_autonomous_running():
        return False

    popen_kwargs = {
        "cwd": PROJECT_ROOT,
        "stdin": subprocess.DEVNULL,
        "stdout": subprocess.DEVNULL,
        "stderr": subprocess.DEVNULL,
        "close_fds": True,
        "env": _build_subprocess_env(),
    }

    if os.name == "nt":
        popen_kwargs["creationflags"] = DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP
    else:
        popen_kwargs["start_new_session"] = True

    subprocess.Popen(build_autonomous_command(), **popen_kwargs)
    return True


def log_mensaje(mensaje):
    """Registra un mensaje en archivo de log y en consola."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    msg_completo = f"[{timestamp}] {mensaje}"
    _safe_console_print(msg_completo)
    try:
        os.makedirs(JSON_DIR, exist_ok=True)
        with open(LOG_FILE, "a", encoding="utf-8") as file_handle:
            file_handle.write(msg_completo + "\n")
    except Exception:
        pass


def obtener_numero_mayor_existente(pais, tipo="excel"):
    """Obtiene el número mayor para archivos existentes de un país."""
    try:
        if tipo == "excel":
            pattern = os.path.join(RESULTS_DIR, f"resultados_{pais}*.xlsx")
        else:
            pattern = os.path.join(RESULTS_DIR, f"screenshots_{pais}*/")

        matches = glob.glob(pattern)
        max_num = 0
        for match in matches:
            try:
                if tipo == "excel":
                    base = os.path.basename(match).replace(f"resultados_{pais}", "").replace(".xlsx", "")
                else:
                    base = os.path.basename(match.rstrip("/\\")).replace(f"screenshots_{pais}", "")
                if base.isdigit():
                    max_num = max(max_num, int(base))
            except Exception:
                pass
        return max_num
    except Exception as exc:
        log_mensaje(f"❌ Error obteniendo número máximo para {pais} ({tipo}): {exc}")
        return 0


def _build_country_command(pais_nombre, env_param):
    if getattr(sys, 'frozen', False):
        return [sys.executable, "--run-country", pais_nombre, "--environment", env_param, "--scheduled"]
    return [sys.executable, os.path.join(PROJECT_ROOT, "run.py"), "--run-country", pais_nombre, "--environment", env_param, "--scheduled"]


def _build_lambdatest_command(lt_type, pais_nombre):
    _lt_label = "MAC" if lt_type == "mac" else "ANDROID"
    build_name = f"Osocio Automatizado LT {_lt_label} - {pais_nombre}"
    if getattr(sys, 'frozen', False):
        return [sys.executable, "--run-lambdatest", lt_type, "--pais", pais_nombre, "--build-name", build_name]
    return [sys.executable, os.path.join(PROJECT_ROOT, "run.py"), "--run-lambdatest", lt_type, "--pais", pais_nombre, "--build-name", build_name]


def _get_lt_results_dir(lt_type):
    """Devuelve el directorio de resultados de LambdaTest según el tipo."""
    if lt_type == "mac":
        return os.path.join(PROJECT_ROOT, "resultados_lambdatestmac")
    return os.path.join(PROJECT_ROOT, "resultados_lambdatest_android")


def ejecutar_tests(programacion):
    """Ejecuta los tests según la programación y recopila resultados."""
    try:
        log_mensaje(" INICIANDO EJECUCIÓN PROGRAMADA AUTÓNOMA")
        log_mensaje(f" Países: {programacion['paises']}")
        log_mensaje(f" Navegadores: {programacion['navegadores']}")
        log_mensaje(f" Viewports: {programacion['viewports']}")

        resultados_ejecucion = []
        total_ejecutados = 0
        total_scripts = 0

        _LT_TYPES = ("lambdatest_mac", "lambdatest_android")

        for pais_nombre in programacion["paises"]:
            for navegador in programacion["navegadores"]:
                if navegador in _LT_TYPES:
                    total_scripts += 1  # LambdaTest: una ejecución por país (sin loop de viewports)
                else:
                    for _ in programacion["viewports"]:
                        total_scripts += 1

        viewport_nombres = {
            "fullscreen": "desktop",
            "600x738": "mobile",
        }

        modo_mercados = programacion.get("modo_mercados", "consecutivo")
        _lock = threading.Lock()

        def _run_lt(pais_nombre, navegador):
            # ── LambdaTest: no usa viewports. SIEMPRE secuencial ──────────────
            nonlocal total_ejecutados
            lt_type = "mac" if navegador == "lambdatest_mac" else "android"
            with _lock:
                total_ejecutados += 1
                _idx = total_ejecutados
            log_mensaje(f" Ejecutando {_idx}/{total_scripts}: LambdaTest {lt_type} — {pais_nombre}")

            lt_results_dir = _get_lt_results_dir(lt_type)
            archivos_antes = set(os.listdir(lt_results_dir)) if os.path.isdir(lt_results_dir) else set()

            try:
                result = subprocess.run(
                    _build_lambdatest_command(lt_type, pais_nombre),
                    check=False, capture_output=True, text=True, timeout=600,
                    cwd=PROJECT_ROOT, env=_build_subprocess_env(),
                )
                if result.returncode == 0:
                    log_mensaje(f" Completado: LambdaTest {lt_type} — {pais_nombre}")
                else:
                    log_mensaje(f" Con errores: LambdaTest {lt_type} — {pais_nombre}")
                    if result.stderr:
                        log_mensaje(f"   Error: {result.stderr.strip()}")

                if os.path.isdir(lt_results_dir):
                    archivos_despues = set(os.listdir(lt_results_dir))
                    nuevos_excels = [f for f in (archivos_despues - archivos_antes) if f.endswith(".xlsx")]
                    if nuevos_excels:
                        excel_file = os.path.join(lt_results_dir, sorted(nuevos_excels)[-1])
                        resultado = {
                            "pais": pais_nombre, "navegador": navegador, "viewport": lt_type,
                            "estado": "completado" if result.returncode == 0 else "con_errores",
                            "excel_path": excel_file, "screenshots_dir": None,
                        }
                        with _lock:
                            resultados_ejecucion.append(resultado)
                        log_mensaje(f"    ✅ Resultado LambdaTest registrado: {os.path.basename(excel_file)}")
            except subprocess.TimeoutExpired:
                log_mensaje(f"⏰ Timeout: LambdaTest {lt_type} — {pais_nombre}")
            except Exception as exc:
                log_mensaje(f" Error en LambdaTest {lt_type} — {pais_nombre}: {exc}")
            time.sleep(2)

        def _run_browser_pais(pais_nombre):
            # Corre todos los browsers/viewports de UN país, secuencialmente.
            # La detección de Excel es por país (glob resultados_{pais}*), por lo
            # que distintos países pueden correr en paralelo sin colisionar.
            nonlocal total_ejecutados
            for navegador in programacion["navegadores"]:
                if navegador in _LT_TYPES:
                    continue
                for viewport in programacion["viewports"]:
                    with _lock:
                        total_ejecutados += 1
                        _idx = total_ejecutados
                    env_param = f"{navegador}_{'desktop' if viewport == 'fullscreen' else 'mobile'}"
                    runner_name = f"Formulario_{pais_nombre}_Main"
                    log_mensaje(f" Ejecutando {_idx}/{total_scripts}: {runner_name} ({env_param})")

                    num_anterior_excel = obtener_numero_mayor_existente(pais_nombre, "excel")
                    num_anterior_screenshots = obtener_numero_mayor_existente(pais_nombre, "screenshots")
                    log_mensaje(f"    Números previos - Excel: {num_anterior_excel}, Screenshots: {num_anterior_screenshots}")

                    try:
                        result = subprocess.run(
                            _build_country_command(pais_nombre, env_param),
                            check=False, capture_output=True, text=True, timeout=300,
                            cwd=PROJECT_ROOT, env=_build_subprocess_env(),
                        )
                        if result.returncode == 0:
                            log_mensaje(f" Completado: {runner_name}")
                        else:
                            log_mensaje(f" Script con errores: {runner_name}")
                            if result.stderr:
                                log_mensaje(f"   Error: {result.stderr.strip()}")

                        num_nuevo_excel = obtener_numero_mayor_existente(pais_nombre, "excel")
                        log_mensaje(f"    📊 Números después - Excel: {num_nuevo_excel}")

                        if num_nuevo_excel > num_anterior_excel:
                            excel_file = os.path.join(RESULTS_DIR, f"resultados_{pais_nombre}{num_nuevo_excel}.xlsx")
                            screenshots_dir = os.path.join(RESULTS_DIR, f"screenshots_{pais_nombre}{num_nuevo_excel}")
                            resultado = {
                                "pais": pais_nombre, "navegador": navegador,
                                "viewport": viewport_nombres.get(viewport, viewport),
                                "estado": "completado" if result.returncode == 0 else "con_errores",
                                "excel_path": excel_file, "screenshots_dir": screenshots_dir,
                            }
                            with _lock:
                                resultados_ejecucion.append(resultado)
                            log_mensaje(f"    ✅ Resultado registrado: {os.path.basename(excel_file)}")
                        elif result.returncode == 0:
                            log_mensaje(f"    ⚠️ No se detectó nuevo Excel (anterior: {num_anterior_excel}, actual: {num_nuevo_excel})")
                    except subprocess.TimeoutExpired:
                        log_mensaje(f"⏰ Timeout: {runner_name}")
                    except Exception as exc:
                        log_mensaje(f" Error ejecutando {runner_name}: {exc}")
                    time.sleep(2)

        def _run_mercado_completo(pais_nombre):
            """Corre LambdaTest + browsers locales para un país.
            Si modo paralelo, todos los dispositivos corren a la vez
            (LT y local no compiten — máquinas distintas)."""
            _threads = []
            for navegador in programacion["navegadores"]:
                if navegador in _LT_TYPES:
                    _threads.append(threading.Thread(
                        target=_run_lt, args=(pais_nombre, navegador), daemon=True))
            _has_local = any(n not in _LT_TYPES for n in programacion["navegadores"])
            if _has_local:
                _threads.append(threading.Thread(
                    target=_run_browser_pais, args=(pais_nombre,), daemon=True))

            if modo_mercados == "paralelo" and len(_threads) > 1:
                for t in _threads:
                    t.start()
                for t in _threads:
                    t.join()
            else:
                # Secuencial: LT primero, después locales
                for navegador in programacion["navegadores"]:
                    if navegador in _LT_TYPES:
                        _run_lt(pais_nombre, navegador)
                _run_browser_pais(pais_nombre)

        # Ejecución por mercado (país): consecutivo o paralelo según configuración.
        if modo_mercados == "paralelo" and len(programacion["paises"]) > 1:
            log_mensaje(" ▶ Modo PARALELO: ejecutando mercados (países) en simultáneo")
            with ThreadPoolExecutor(max_workers=min(len(programacion["paises"]), 9)) as _ex:
                _futs = [_ex.submit(_run_mercado_completo, p) for p in programacion["paises"]]
                for _f in _futs:
                    try:
                        _f.result()
                    except Exception as _e:
                        log_mensaje(f" Error en mercado paralelo: {_e}")
        else:
            for pais_nombre in programacion["paises"]:
                _run_mercado_completo(pais_nombre)

        log_mensaje(f" EJECUCIÓN AUTÓNOMA COMPLETADA: {total_ejecutados}/{total_scripts} scripts")
        log_mensaje(f" Resultados recopilados: {len(resultados_ejecucion)} ejecuciones")
        for index, resultado in enumerate(resultados_ejecucion, 1):
            log_mensaje(f"   [{index}] {resultado['pais']} - {resultado['excel_path']}")
        return resultados_ejecucion

    except Exception as exc:
        log_mensaje(f" ERROR CRÍTICO en ejecución autónoma: {exc}")
        return []


def main():
    """Función principal del ejecutor autónomo."""
    if not acquire_autonomous_lock():
        log_mensaje("⚠️ Ya existe un ejecutor autónomo en ejecución. Se cancela este inicio duplicado.")
        return

    _safe_console_print("=" * 60)
    _safe_console_print(" EJECUTOR AUTÓNOMO GDCP - INICIADO")
    _safe_console_print("=" * 60)
    _safe_console_print(f" Ubicación: {JSON_DIR}")
    _safe_console_print(" Monitoreando: programacion_test.json")
    _safe_console_print("⏰ Verificando cada minuto tests programados")
    _safe_console_print(" Mantén esta ventana abierta para ejecuciones automáticas")
    _safe_console_print("⏹  Presiona Ctrl+C para salir")
    _safe_console_print("=" * 60)

    DIAS_ES = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    last_triggered = {}  # {(dia, hora): date} — evita doble disparo en el mismo slot

    try:
        while True:
            programacion = cargar_programacion()

            if programacion and programacion.get("tipo") == "semanal":
                ahora      = datetime.now()
                dia_actual  = DIAS_ES[ahora.weekday()]
                slot_min    = (ahora.minute // 15) * 15
                hora_actual = f"{ahora.hour:02d}:{slot_min:02d}"
                horarios_hoy = programacion.get("horarios", {}).get(dia_actual, [])

                if hora_actual in horarios_hoy:
                    key = (dia_actual, hora_actual)
                    if last_triggered.get(key) != ahora.date():
                        last_triggered[key] = ahora.date()
                        log_mensaje(f" ¡SLOT DETECTADO! {dia_actual} {hora_actual} — ejecutando test...")
                        resultados = ejecutar_tests(programacion)

                        log_mensaje(f"\n📊 RESULTADOS RECOPILADOS: {len(resultados)} ejecuciones")
                        for index, resultado in enumerate(resultados, 1):
                            log_mensaje(
                                f"   [{index}] {resultado['pais']} ({resultado['navegador']}/{resultado['viewport']}) - {resultado['estado']}"
                            )

                        if resultados:
                            log_mensaje(f" ✅ {len(resultados)} ejecuciones completadas")
                            log_mensaje(" 📧 INICIANDO ENVÍO DE EMAIL CONSOLIDADO...")
                            try:
                                envio_exitoso = enviar_email_resultados_consolidados(resultados)
                                if envio_exitoso:
                                    log_mensaje(" ✅ Email enviado exitosamente")
                                    time.sleep(10)
                                else:
                                    log_mensaje(" ⚠️ Proceso de email devolvió False")
                            except Exception as exc:
                                import traceback
                                log_mensaje(f" ❌ Excepción al enviar email: {exc}")
                                log_mensaje(traceback.format_exc())
                            # NO limpiar — programación semanal persiste para la próxima semana
                            log_mensaje(" Slot completado. Próxima ejecución: semana siguiente.")
                        else:
                            log_mensaje(" ⚠️ No se recopilaron resultados")

            time.sleep(60)

    except KeyboardInterrupt:
        log_mensaje(" Ejecutor autónomo finalizado por el usuario")
    except Exception as exc:
        log_mensaje(f" Error inesperado: {exc}")


if __name__ == "__main__":
    main()