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

from interface.helpers_interface import (
    enviar_email_resultados_consolidados,
    esperar_envios_pendientes,
)
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
    # Los runners imprimen emojis. Con capture_output el hijo escribe a un pipe con
    # el encoding local (cp1252 en Windows) y muere con UnicodeEncodeError antes de
    # enviar el lead. Forzar UTF-8 evita ese fallo en toda la ruta headless.
    env["PYTHONIOENCODING"] = "utf-8"
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


def _basenames_resultados(pais, navegador):
    """Prefijos reales de los archivos de una corrida PROGRAMADA.

    base_form_filler usa el prefijo "Automatizacion_" (no "resultados_") cuando
    is_scheduled=True, y agrega el dispositivo: Automatizacion_<Pais>_<Dev><N>.xlsx.
    Buscar con el prefijo equivocado hacía que el ejecutor nunca detectara los
    resultados y, por lo tanto, nunca enviara el email consolidado.
    """
    dev = {"chrome": "Chrome", "firefox": "Firefox", "edge": "Edge"}.get(navegador, "Chrome")
    return f"Automatizacion_{pais}_{dev}", f"Automatizacion_screenshots_{pais}_{dev}"


def obtener_numero_mayor_existente(pais, tipo="excel", basename=None):
    """Obtiene el número de corrida más alto ya existente para ese prefijo."""
    try:
        base = basename or (f"resultados_{pais}" if tipo == "excel" else f"screenshots_{pais}")
        pattern = os.path.join(RESULTS_DIR, f"{base}*.xlsx" if tipo == "excel" else f"{base}*/")

        matches = glob.glob(pattern)
        max_num = 0
        for match in matches:
            try:
                if tipo == "excel":
                    sufijo = os.path.basename(match).replace(base, "").replace(".xlsx", "")
                else:
                    sufijo = os.path.basename(match.rstrip("/\\")).replace(base, "")
                if sufijo.isdigit():
                    max_num = max(max_num, int(sufijo))
            except Exception:
                pass
        return max_num
    except Exception as exc:
        log_mensaje(f"Error obteniendo número máximo para {pais} ({tipo}): {exc}")
        return 0


def _build_country_command(pais_nombre, env_param, excel_suffix=""):
    extra = ["--excel-suffix", excel_suffix] if excel_suffix else []
    if getattr(sys, 'frozen', False):
        return [sys.executable, "--run-country", pais_nombre, "--environment", env_param, "--scheduled"] + extra
    return [sys.executable, os.path.join(PROJECT_ROOT, "run.py"), "--run-country", pais_nombre,
            "--environment", env_param, "--scheduled"] + extra


def _t3_excel_existe(pais_nombre, navegador):
    """True si el mercado tiene Excel de formulario T3 2.0 (AEM) para ese browser."""
    dev = {"chrome": "Chrome", "firefox": "Firefox", "edge": "Edge"}.get(navegador, "Chrome")
    return os.path.isfile(os.path.join(
        PROJECT_ROOT, "data", f"Lead_information_Formulario_{pais_nombre}_{dev}_T3.xlsx"))


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
        _corre_t3 = bool(programacion.get("t3_also"))

        for pais_nombre in programacion["paises"]:
            for navegador in programacion["navegadores"]:
                if navegador in _LT_TYPES:
                    total_scripts += 1  # LambdaTest: una ejecución por país (sin loop de viewports)
                else:
                    for _ in programacion["viewports"]:
                        total_scripts += 1
                        if _corre_t3 and _t3_excel_existe(pais_nombre, navegador):
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
                    check=False, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=600,
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

            def _run_uno(navegador, viewport, excel_suffix=""):
                nonlocal total_ejecutados
                with _lock:
                    total_ejecutados += 1
                    _idx = total_ejecutados
                env_param = f"{navegador}_{'desktop' if viewport == 'fullscreen' else 'mobile'}"
                runner_name = f"Formulario_{pais_nombre}_Main{excel_suffix}"
                log_mensaje(f" Ejecutando {_idx}/{total_scripts}: {runner_name} ({env_param})")

                base_excel, base_ss = _basenames_resultados(pais_nombre, navegador)
                num_anterior_excel = obtener_numero_mayor_existente(pais_nombre, "excel", base_excel)
                num_anterior_screenshots = obtener_numero_mayor_existente(pais_nombre, "screenshots", base_ss)
                log_mensaje(f"    Números previos - Excel: {num_anterior_excel}, Screenshots: {num_anterior_screenshots}")

                try:
                    result = subprocess.run(
                        _build_country_command(pais_nombre, env_param, excel_suffix),
                        check=False, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=300,
                        cwd=PROJECT_ROOT, env=_build_subprocess_env(),
                    )
                    if result.returncode == 0:
                        log_mensaje(f" Completado: {runner_name}")
                    else:
                        log_mensaje(f" Script con errores: {runner_name}")
                        if result.stderr:
                            log_mensaje(f"   Error: {result.stderr.strip()}")

                    num_nuevo_excel = obtener_numero_mayor_existente(pais_nombre, "excel", base_excel)
                    log_mensaje(f"    📊 Números después - Excel: {num_nuevo_excel}")

                    if num_nuevo_excel > num_anterior_excel:
                        excel_file = os.path.join(RESULTS_DIR, f"{base_excel}{num_nuevo_excel}.xlsx")
                        screenshots_dir = os.path.join(RESULTS_DIR, f"{base_ss}{num_nuevo_excel}")
                        resultado = {
                            "pais": pais_nombre + (" (T3)" if excel_suffix else ""), "navegador": navegador,
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

            for navegador in programacion["navegadores"]:
                if navegador in _LT_TYPES:
                    continue
                for viewport in programacion["viewports"]:
                    _run_uno(navegador, viewport)
                    # "Correr también T3": segunda pasada con el Excel …_T3.xlsx del mercado.
                    if _corre_t3 and _t3_excel_existe(pais_nombre, navegador):
                        _run_uno(navegador, viewport, "_T3")

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


def ejecutar_masivo_autonomo(programacion):
    """Ejecuta la revisión masiva de forms de forma autónoma."""
    try:
        from core.massive_check_runner import run_massive_check
        log_mensaje(" INICIANDO REVISIÓN MASIVA AUTÓNOMA")
        excel_path = programacion.get("var_excel_masivo") or os.path.join(PROJECT_ROOT, "data", "GM Forms - 2026.xlsx")
        custom_cols = {
            "segmento": programacion.get("var_col_segmento", "SEGMENTO"),
            "estado": programacion.get("var_col_estado", "ESTADO"),
            "url_live": programacion.get("var_col_url_live", "URL LIVE"),
            "url_secure": programacion.get("var_col_url_secure", "URL SECURE"),
        }
        paises = programacion.get("paises", [])
        log_mensaje(f" Mercados a revisar: {paises}")
        log_mensaje(f" Excel: {excel_path}")
        
        success, info = run_massive_check(
            excel_path, custom_cols,
            selected_markets=paises,
            borrar_comentarios=programacion.get("var_borrar_comentarios", False),
            tomar_capturas=programacion.get("var_tomar_capturas", True),
            solo_fails=programacion.get("var_solo_fails", False),
            browser="chrome", headless=True
        )
        msg = info.get("msg", "") if isinstance(info, dict) else str(info)
        log_mensaje(f" Revisión masiva completada ({'OK' if success else 'FAIL'}): {msg}")
        return success
    except Exception as exc:
        log_mensaje(f" Error en revisión masiva autónoma: {exc}")
        return False


def _cargar_programacion_leads():
    """Programación de Envío de Leads: la que escribe la app (programacion_leads.json).
    Cae al archivo legado programacion_test.json para instalaciones viejas."""
    prog = cargar_programacion("programacion_leads.json")
    if prog and prog.get("tipo") == "semanal":
        return prog
    prog = cargar_programacion("programacion_test.json")
    if prog and prog.get("tipo") == "semanal" and prog.get("modo_tarea", "leads") != "masivo":
        return prog
    return None


def run_once():
    """Ejecuta la programación guardada UNA sola vez y termina.

    Pensado para el Programador de tareas de Windows: el disparo horario lo pone
    Windows, así que acá no hay loop de espera ni control de slots — si la tarea
    corrió, se ejecuta. Devuelve 0 si ejecutó algo, 1 si no había nada que hacer.
    """
    if not acquire_autonomous_lock():
        log_mensaje("⚠️ Ya hay una ejecución autónoma en curso. Se cancela este disparo.")
        return 1

    programacion = _cargar_programacion_leads()
    if not programacion:
        log_mensaje("⚠️ No hay programación de leads guardada (json/programacion_leads.json). Nada que ejecutar.")
        return 1
    if not programacion.get("paises"):
        log_mensaje("⚠️ La programación no tiene mercados seleccionados. Nada que ejecutar.")
        return 1

    log_mensaje("▶ Disparo del Programador de Windows — ejecución única de Envío de Leads")
    resultados = ejecutar_tests(programacion)

    if resultados:
        log_mensaje(f" ✅ {len(resultados)} ejecuciones completadas")
        log_mensaje(" 📧 INICIANDO ENVÍO DE EMAIL CONSOLIDADO...")
        try:
            if enviar_email_resultados_consolidados(resultados):
                # OJO: la función de arriba devuelve True al ENCOLAR el mail, no al
                # entregarlo. Como el worker de Outlook es daemon, si el proceso termina
                # acá el mail se pierde. Hay que esperar a que la cola se vacíe.
                if esperar_envios_pendientes(180):
                    log_mensaje(" ✅ Email entregado a Outlook")
                else:
                    log_mensaje(" ⚠️ Timeout esperando que Outlook envíe el email consolidado")
            else:
                log_mensaje(" ⚠️ Proceso de email devolvió False")
        except Exception as exc:
            import traceback
            log_mensaje(f" ❌ Excepción al enviar email: {exc}")
            log_mensaje(traceback.format_exc())
    else:
        log_mensaje(" ⚠️ La ejecución no produjo resultados.")

    log_mensaje(" Ejecución única finalizada.")
    return 0


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
                        log_mensaje(f" ¡SLOT DETECTADO! {dia_actual} {hora_actual} — ejecutando...")
                        modo_t = programacion.get("modo_tarea", "leads")
                        if modo_t == "masivo":
                            log_mensaje(" ▶ Ejecutando Revisión Masiva programada...")
                            ejecutar_masivo_autonomo(programacion)
                        else:
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
                                        # Esperar la entrega real en vez de dormir 10s a ciegas.
                                        if esperar_envios_pendientes(180):
                                            log_mensaje(" ✅ Email entregado a Outlook")
                                        else:
                                            log_mensaje(" ⚠️ Timeout esperando el envío del email")
                                    else:
                                        log_mensaje(" ⚠️ Proceso de email devolvió False")
                                except Exception as exc:
                                    import traceback
                                    log_mensaje(f" ❌ Excepción al enviar email: {exc}")
                                    log_mensaje(traceback.format_exc())
                        log_mensaje(" Slot completado. Próxima ejecución: semana siguiente.")

            time.sleep(60)

    except KeyboardInterrupt:
        log_mensaje(" Ejecutor autónomo finalizado por el usuario")
    except Exception as exc:
        log_mensaje(f" Error inesperado: {exc}")


if __name__ == "__main__":
    main()