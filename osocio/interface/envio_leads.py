"""Envío de Leads: la corrida completa que dispara el botón "Enviar".

Vivía adentro de `iniciar_interfaz()`, en `main_interface.py`, como una función anidada de
1.054 líneas — casi el 20% de ese archivo. Se movió acá **sin tocar una sola línea del
cuerpo**: lo único que cambió es que las 31 cosas que tomaba del closure ahora llegan en
un objeto de contexto y se desempaquetan al principio.

Sobre por qué llegan los objetos y no sus valores: tres de esas capturas
(`var_pausar_autenticacion`, `var_preview_navegador`, `selected_disp`) se leen adentro de
funciones anidadas que corren **después**, en hilos de trabajo. Si el usuario destilda
"pausar autenticación" con la corrida ya empezada, las sesiones siguientes tienen que ver
el valor nuevo. Congelarlas acá rompería eso en silencio.
"""

import os

from osocio.core.variantes import T3 as VAR_T3
import threading
import tkinter as tk
from tkinter import messagebox

from osocio.paths import DATA_DIR
from osocio.interface.helpers_interface import (
    cargar_config_global,
    guardar_config_global,
)
from dataclasses import dataclass

from osocio.utils.fixed_field_mapping_store import build_excel_columns_for_country


@dataclass
class ContextoEnvioLeads:
    """Todo lo que el envío de leads necesita de la pantalla, declarado y agrupado.

    Antes esto era el closure de `iniciar_interfaz()`: 31 nombres sueltos que la función
    tomaba del aire. Declararlos acá cumple tres cosas — dice exactamente qué necesita la
    función, permite construirla en un test sin levantar la UI, y evita que alguien agregue
    una dependencia nueva sin que se note.

    Ojo con los tipos: casi todos son **objetos de Tkinter, no valores**. Eso es a
    propósito. `var_pausar_autenticacion`, `var_preview_navegador` y `selected_disp` se
    leen adentro de funciones que corren después, en hilos: si el usuario cambia un
    checkbox con la corrida ya empezada, las sesiones siguientes tienen que ver el valor
    nuevo. Guardar el valor en vez del objeto rompería eso sin que nada avise.
    """

    # --- la ventana y sus controles -----------------------------------------------
    root: object                      # la ventana principal: binds, protocolos, after()
    btn_enviar: object                # se deshabilita mientras corre
    btn_retry_leads: object           # se habilita si quedaron fallidos
    email_entry: object               # destinatario del mail final

    # --- qué eligió el usuario ------------------------------------------------------
    var_t3: object                    # formularios 2.0 de AEM
    var_t3_also: object               # además de los normales
    var_sched_t3: object              # el equivalente para corridas programadas
    var_enviar_email: object
    var_modo_email: object            # "por_pais" | "consolidado"
    var_adjuntar_res: object          # adjuntar el Excel de resultados
    var_adjuntar_ss: object           # adjuntar las capturas
    var_ver_navegador: object         # mostrar el navegador en vez de correr oculto
    var_url_parallel: object          # varias URLs a la vez
    url_max_var: object               # cuántas en paralelo
    var_pausar_autenticacion: object  # se lee EN VIVO, por sesión
    var_preview_navegador: object     # se lee EN VIVO, por sesión

    # --- qué mercados y dispositivos ------------------------------------------------
    paises_list: object               # todos los países disponibles
    selected_countries: object        # los tildados
    selected_disp: object             # dispositivos tildados; se lee EN VIVO
    active_p_tab: object              # la pestaña de país abierta
    mercados_mode: object             # secuencial o paralelo
    excels_mode: object
    excel_mode_holder: object
    scheduler_cfg_leads: object       # la configuración de la corrida programada

    # --- estado de la corrida --------------------------------------------------------
    _exec_state: object               # qué falló, para "reintentar fallidos"
    _RETRY_DEVICE_SUFFIX: object

    # --- cosas que la función le pide a la pantalla -----------------------------------
    log_message: object               # escribir en la consola de la app
    refresh_execute_state: object     # recalcular si el botón va habilitado
    _forzar_foreground: object        # traer la ventana al frente
    _restore_from_tray: object        # sacarla de la bandeja
    _build_retry_excel: object        # armar el Excel con solo las filas que fallaron




@dataclass(frozen=True)
class OpcionesEnvioLeads:
    """Lo que el usuario dejó elegido cuando apretó "Enviar".

    Son las once opciones que la corrida lee **una sola vez, al principio**, y que a partir
    de ahí no vuelve a mirar. Congelarlas acá tiene dos ventajas: la corrida deja de
    depender de widgets de Tkinter para saber qué hacer, y queda claro de un vistazo qué
    configura al envío.

    Ojo con lo que NO está acá. `var_pausar_autenticacion`, `var_preview_navegador` y
    `selected_disp` quedaron afuera a propósito: se leen adentro de funciones que corren
    después, en hilos, así que tienen que seguir viendo el valor actual. Si el usuario
    destilda "pausar autenticación" con la corrida ya empezada, las sesiones siguientes
    respetan el cambio. Congelarlas acá rompería eso en silencio.

    Por eso es `frozen=True`: si alguien intenta mutarla, es señal de que lo que necesita
    es una lectura en vivo, no una opción.
    """

    t3: bool                 # correr los formularios 2.0 de AEM
    t3_also: bool            # ademas de los normales
    url_paralelo: bool       # varias URLs a la vez
    url_max: int             # cuantas en paralelo (1 a 20)
    enviar_mail: bool
    modo_email: str          # "por_pais" | "consolidado"
    destinatario: str
    adjuntar_resultados: bool
    adjuntar_capturas: bool
    ver_navegador: bool      # mostrar el navegador en vez de correr oculto


def _leer_opciones(ctx, scheduled):
    """Toma una foto de lo que el usuario eligio en la pantalla."""
    t3 = bool(ctx.var_t3.get())
    return OpcionesEnvioLeads(
        t3=t3,
        # En una corrida programada manda la casilla del scheduler, no la de la pantalla.
        t3_also=(not t3) and bool(
            ctx.var_sched_t3.get() if scheduled else ctx.var_t3_also.get()
        ),
        url_paralelo=bool(ctx.var_url_parallel.get()),
        url_max=max(1, min(20, int(ctx.url_max_var.get()))),
        enviar_mail=bool(ctx.var_enviar_email.get()),
        modo_email=ctx.var_modo_email.get(),
        destinatario=ctx.email_entry.get().strip(),
        adjuntar_resultados=bool(ctx.var_adjuntar_res.get()),
        adjuntar_capturas=bool(ctx.var_adjuntar_ss.get()),
        ver_navegador=bool(ctx.var_ver_navegador.get()),
    )


def ejecutar_envio_leads(ctx, scheduled=False, retry_only=None):
    """Corre el envío de leads. `ctx` trae los widgets, variables y helpers de la UI."""
    # Constantes de estilo y helpers que siguen viviendo en main_interface. El import va
    # adentro de la funcion a proposito: main_interface importa este modulo al llamar, asi
    # que a nivel de modulo seria un ciclo esperando a pasar.
    from osocio.interface.main_interface import (
        BACKEND_OK,
        BORDER_COLOR,
        BUTTON_INACTIVE,
        EXECUTE_BG,
        EXECUTE_FG,
        TEXT_SECONDARY,
        _APP_BASE,
        _BACKEND_IMPORT_ERROR,
        _ensure_serialized_setup,
        _etiqueta_t3,
        _generic_excel_path_for,
        _lead_excel_name,
        get_button_icon,
    )

    _RETRY_DEVICE_SUFFIX = ctx._RETRY_DEVICE_SUFFIX

    opciones = _leer_opciones(ctx, scheduled)
    _build_retry_excel = ctx._build_retry_excel
    _exec_state = ctx._exec_state
    _forzar_foreground = ctx._forzar_foreground
    _restore_from_tray = ctx._restore_from_tray
    active_p_tab = ctx.active_p_tab
    btn_enviar = ctx.btn_enviar
    btn_retry_leads = ctx.btn_retry_leads
    excel_mode_holder = ctx.excel_mode_holder
    excels_mode = ctx.excels_mode
    log_message = ctx.log_message
    mercados_mode = ctx.mercados_mode
    paises_list = ctx.paises_list
    refresh_execute_state = ctx.refresh_execute_state
    root = ctx.root
    scheduler_cfg_leads = ctx.scheduler_cfg_leads
    selected_countries = ctx.selected_countries
    selected_disp = ctx.selected_disp
    var_pausar_autenticacion = ctx.var_pausar_autenticacion
    var_preview_navegador = ctx.var_preview_navegador

    if not BACKEND_OK:
        messagebox.showerror("Ejecutar", f"El backend no está disponible.\n{_BACKEND_IMPORT_ERROR}")
        return
    if scheduled:
        # El reintento es una acción manual (la elige la persona tildando países/
        # dispositivos en pantalla) — no tiene sentido en una corrida desatendida.
        retry_only = None

    def _build_retry_market_jobs(retry_only):
        """market_jobs para un reintento: una sesión por (país, dispositivo) presente
        en retry_only, apuntando a un Excel temporal armado SOLO con las filas que
        fallaron la última vez. No pasa por _sessions_for/T3/"sesión por URL": cada
        entrada de retry_only ya identifica exactamente qué (país, dispositivo —
        incluye el sufijo ·T3 si corresponde) hay que re-correr."""
        tmpdir = os.path.join(_APP_BASE, "temporales")
        by_pais = {}
        for (r_pais, device_label), fail_rows in retry_only.items():
            info = _exec_state.get("last_fails", {}).get((r_pais, device_label))
            if not info:
                continue
            original_excel = info.get("excel")
            base_device = device_label[:-len("·T3")] if device_label.endswith("·T3") else device_label
            match = next((row for row in _RETRY_DEVICE_SUFFIX if row[0] == base_device), None)
            if not match:
                continue
            _suffix, r_dtype, r_browser, _key = match
            tmp_path, row_map = _build_retry_excel(original_excel, fail_rows, tmpdir, r_pais, device_label)
            if not tmp_path:
                continue
            by_pais.setdefault(r_pais, []).append({
                "pais": r_pais, "dtype": r_dtype, "browser": r_browser, "device": device_label,
                "excel": tmp_path,
                # Mismo criterio que _t3_extra_sessions: sin esto, el T3 se fusiona
                # con el mercado normal en el email y su resultado desaparece.
                "email_label": _etiqueta_t3(r_pais) if device_label.endswith("·T3") else None,
                "_retry_source_excel": original_excel,
                "_retry_row_map": row_map,
            })
        return list(by_pais.items())

    # Sufijos de dispositivo ↔ tipo de ejecución (para localizar los Excels generados)
    # (suffix Excel, dtype, browser, key en selected_disp)
    _DEVICE_SUFFIX = [
        ("Chrome", "desktop", "chrome", "chrome"),
        ("Firefox", "desktop", "firefox", "firefox"),
        ("Edge", "desktop", "edge", "edge"),
        ("Mac", "mac", None, "mac lt"),
        ("Android", "android", None, "android lt"),
    ]

    # Formularios T3 2.0 (Adobe AEM): usar los Excels con nombre …_T3.xlsx
    t3 = opciones.t3
    # "Correr también T3": el mercado corre normal y, además, con su Excel …_T3.
    # Cada pestaña tiene su propio check de "correr también los T3": la corrida
    # programada debe mirar el de Envío de Leads Programados, no el de la pestaña
    # manual (si no, el T3 nunca entraba en las corridas automáticas).
    t3_also = opciones.t3_also

    def _t3_extra_sessions(sessions):
        """Duplica cada sesión apuntando a su Excel …_T3.xlsx. Sólo devuelve las
        que tienen Excel T3 real: un mercado sin form T3 no suma sesiones (y así
        no dispara la validación de 'Excel faltante' más abajo)."""
        if not t3_also:
            return []
        extra = []
        for s in sessions:
            t3_path = os.path.join(DATA_DIR, _lead_excel_name(s["pais"], s["device"], True))
            if not os.path.exists(t3_path):
                continue
            e = dict(s)
            e["excel"] = t3_path
            e["device"] = f"{s['device']}·T3"
            e["variante"] = VAR_T3
            # Nombre propio en el email: si va como "Brasil" se fusiona con el
            # mercado normal y el resultado del T3 desaparece del reporte.
            e["email_label"] = _etiqueta_t3(s["pais"])
            extra.append(e)
        return extra

    def _sessions_for(pais):
        """Una sesión por dispositivo tildado. Cada dispositivo (desktop y LT)
        usa su propio Excel (…_Chrome/_Firefox/_Edge/_Mac/_Android[_T3].xlsx), que es
        además donde LambdaTest guarda los resultados. No hay fallback: si falta
        el Excel del dispositivo, se detecta luego y no ejecuta."""
        shared = excel_mode_holder[0] == "compartido"
        gpath = _generic_excel_path_for(pais, t3)
        out = []
        for suffix, dtype, browser, key in _DEVICE_SUFFIX:
            if not selected_disp.get(key):
                continue
            path = gpath if shared else os.path.join(DATA_DIR, _lead_excel_name(pais, suffix, t3))
            out.append({"pais": pais, "dtype": dtype, "browser": browser, "device": suffix,
                        "excel": path, "variante": VAR_T3 if t3 else ""})
        return out + _t3_extra_sessions(out)

    if scheduled:
        disp_sched = scheduler_cfg_leads.get("dispositivo", "local")
        p_mode = scheduler_cfg_leads.get("modo_excel", "consecutivo")
        paises_run = list(scheduler_cfg_leads.get("paises", [])) or [active_p_tab[0]]

        market_jobs = []
        for p in paises_run:
            sessions = []
            if disp_sched == "lambdatest_android":
                sessions.append({
                    "pais": p, "dtype": "android", "browser": None, "device": "Android",
                    "excel": os.path.join(DATA_DIR, _lead_excel_name(p, "Android", t3))
                })
            elif disp_sched == "lambdatest_mac":
                sessions.append({
                    "pais": p, "dtype": "mac", "browser": None, "device": "Mac",
                    "excel": os.path.join(DATA_DIR, _lead_excel_name(p, "Mac", t3))
                })
            else: # local
                navs = scheduler_cfg_leads.get("navegadores", []) or ["chrome"]
                for nav in navs:
                    suffix = "Chrome" if nav == "chrome" else "Firefox" if nav == "firefox" else "Edge"
                    sessions.append({
                        "pais": p, "dtype": "desktop", "browser": nav, "device": suffix,
                        "excel": os.path.join(DATA_DIR, _lead_excel_name(p, suffix, t3))
                    })
            sessions += _t3_extra_sessions(sessions)
            market_jobs.append((p, sessions))
            
        mercados_par = (p_mode == "paralelo") and (len(market_jobs) > 1)
        excels_par = (p_mode == "paralelo")
    elif retry_only:
        market_jobs = _build_retry_market_jobs(retry_only)
        mercados_par = (mercados_mode[0] == "paralelo") and (len(market_jobs) > 1)
        excels_par = (excels_mode[0] == "paralelo")
    else:
        if not any(selected_disp.values()):
            messagebox.showwarning("Ejecutar", "Seleccioná al menos un dispositivo / navegador antes de ejecutar.")
            return
        paises_run = [p for p in paises_list if selected_countries.get(p)] or [active_p_tab[0]]
        market_jobs = [(p, _sessions_for(p)) for p in paises_run]

        mercados_par = (mercados_mode[0] == "paralelo") and (len(market_jobs) > 1)
        excels_par = (excels_mode[0] == "paralelo")

    total_sessions = sum(len(js) for _, js in market_jobs)
    if total_sessions == 0:
        if scheduled:
            log_message("[WARN] Ejecución programada sin dispositivos/Excels configurados. Se cancela.")
        elif retry_only:
            messagebox.showwarning("Reintentar fallidos", "No quedan filas válidas para reintentar "
                                    "(puede que el Excel original haya cambiado desde la última corrida).")
        else:
            messagebox.showwarning("Ejecutar", "No hay Excels para ejecutar. Generá datos o seleccioná un dispositivo.")
        return

    # Modo "una sesión por URL": expande cada fila de cada Excel en su propia sesión
    # (un reintento ya corre sólo sobre las filas que fallaron — no tiene sentido
    # volver a fragmentarlas una por URL).
    url_par = (not scheduled) and (not retry_only) and opciones.url_paralelo
    try:
        url_max = opciones.url_max
    except Exception:
        url_max = 6

    def _url_sessions():
        import pandas as pd
        tmpdir = os.path.join(_APP_BASE, "temporales")
        os.makedirs(tmpdir, exist_ok=True)
        out = []
        for _pais, sessions in market_jobs:
            for s in sessions:
                if s["dtype"] != "desktop" or not s["excel"] or not os.path.exists(s["excel"]):
                    out.append(s)  # LT o sin Excel → sesión entera
                    continue
                try:
                    df = pd.read_excel(s["excel"], dtype=str).fillna("")
                except Exception:
                    out.append(s)
                    continue
                if len(df) == 0:
                    continue
                for i in range(len(df)):
                    tmp = os.path.join(tmpdir, f"_url_{_pais}_{s['device']}_{i + 1}.xlsx")
                    try:
                        df.iloc[[i]].to_excel(tmp, index=False)
                    except Exception:
                        continue
                    out.append({"pais": _pais, "dtype": "desktop", "browser": s["browser"],
                                "device": f"{s['device']}·URL{i + 1}", "excel": tmp})
        return out

    if url_par:
        flat_sessions = _url_sessions()
        total_sessions = len(flat_sessions)
        if total_sessions == 0:
            messagebox.showwarning("Ejecutar", "No hay URLs para ejecutar.")
            return
        active_sessions_list = flat_sessions
    else:
        active_sessions_list = []
        for _pais, js in market_jobs:
            active_sessions_list.extend(js)

    # Validar Excels ANTES de arrancar: cada dispositivo seleccionado (desktop y LT)
    # debe tener su propio Excel. Si falta alguno, avisar y no ejecutar (sin modal).
    faltantes = []
    faltantes_sessions = []  # sesiones con Excel inexistente (para poder crearlos vacíos)
    invalidas_ids = set()    # id() de TODAS las sesiones con Excel faltante o vacío
    for s in active_sessions_list:
        _ex = s["excel"]
        _nombre = os.path.basename(_ex) if _ex else \
            f"Lead_information_Formulario_{s['pais']}_{s['device']}.xlsx"
        if not _ex or not os.path.exists(_ex):
            faltantes.append(f"• {s['pais']} · {s['device']}: {_nombre} (no existe)")
            faltantes_sessions.append(s)
            invalidas_ids.add(id(s))
            continue
        # El Excel existe: validar que tenga al menos una fila con datos (no arrancar vacío).
        try:
            import pandas as _pd
            _df = _pd.read_excel(_ex, dtype=str, keep_default_na=False)
            _tiene = (not _df.empty) and any(
                any(str(v).strip() for v in _r.values) for _, _r in _df.iterrows()
            )
        except Exception:
            _tiene = True  # si no se pudo leer (abierto/corrupto), no bloquear por esto
        if not _tiene:
            faltantes.append(f"• {s['pais']} · {s['device']}: {_nombre} (vacío / sin leads)")
            invalidas_ids.add(id(s))
    if faltantes:
        _msg = ("No se encontró un Excel válido (inexistente o vacío) para el/los dispositivo(s) seleccionado(s):\n\n"
                + "\n".join(faltantes))
        if scheduled:
            # Corrida desatendida: nadie va a estar para responder un messagebox.
            # Un askyesno/showerror acá bloquea el mainloop de Tk para SIEMPRE
            # (la app queda "colgada" sin ejecutar nada ni avisar por mail). En vez
            # de preguntar, se descartan las sesiones con Excel faltante/vacío y se
            # sigue con las que sí tienen datos.
            log_message("[WARN] Ejecución programada con Excel(s) faltante(s)/vacío(s), se omiten:\n" + _msg)
            active_sessions_list = [s for s in active_sessions_list if id(s) not in invalidas_ids]
            if not active_sessions_list:
                log_message("[WARN] Ninguna sesión tiene Excel válido. Se cancela la corrida programada.")
                return
        else:
            # Dos opciones: crear los Excel vacíos con las columnas del país (respeta el path
            # _T3 si corresponde), o cerrar. Solo se pueden crear los que NO existen.
            _creables = []
            _seen_paths = set()
            for s in faltantes_sessions:
                _p = s.get("excel")
                if _p and _p not in _seen_paths and "temporales" not in _p:
                    _seen_paths.add(_p)
                    _creables.append(s)
            if _creables:
                _msg += ("\n\n¿Querés CREAR el/los Excel vacío(s) con las columnas del país (para completarlos a mano)?\n\n"
                         "Sí = crear los Excel vacíos    ·    No = cerrar")
                if messagebox.askyesno("Error Excel", _msg):
                    _creados = []
                    for s in _creables:
                        try:
                            _cols = build_excel_columns_for_country(s["pais"])
                            import pandas as _pd
                            os.makedirs(os.path.dirname(s["excel"]) or DATA_DIR, exist_ok=True)
                            _pd.DataFrame(columns=_cols).to_excel(s["excel"], index=False)
                            _creados.append(os.path.basename(s["excel"]))
                        except Exception as _ce:
                            log_message(f"[ERROR] No se pudo crear {s.get('excel')}: {_ce}")
                    if _creados:
                        messagebox.showinfo(
                            "Excel creados",
                            "Se crearon vacíos (solo encabezados). Completá al menos un lead "
                            "(o generá datos) y volvé a ejecutar:\n\n"
                            + "\n".join("• " + n for n in _creados))
            else:
                messagebox.showerror("Error Excel", _msg
                                     + "\n\nGenerá los datos (con al menos un lead) antes de ejecutar.")
            return

    import pandas as pd
    total_leads = 0
    for idx, s in enumerate(active_sessions_list):
        s["sess_id"] = idx
        rc = 0
        if s["dtype"] == "desktop" and s["excel"] and os.path.exists(s["excel"]):
            try:
                rc = len(pd.read_excel(s["excel"]))
            except Exception:
                rc = 1
        else:
            rc = 1
        s["rows_count"] = max(1, rc)
        total_leads += s["rows_count"]

    # Persistir email + enviar_mail en config_global (lo lee el backend de email)
    enviar_mail = opciones.enviar_mail
    _email_modo = opciones.modo_email  # "por_pais" | "consolidado"
    dest = opciones.destinatario
    try:
        _cfg = cargar_config_global()
        _cfg["email_destinatario"] = dest
        _cfg["enviar_mail"] = enviar_mail
        # Opt-in real: sólo adjunta lo que el usuario tildó (no forzar True por default).
        _cfg["adjuntar_resultados"] = opciones.adjuntar_resultados
        _cfg["adjuntar_screenshots"] = opciones.adjuntar_capturas
        guardar_config_global(_cfg)
    except Exception:
        pass

    background = not opciones.ver_navegador  # ver navegador → visible
    stop_event = threading.Event()

    if not scheduled:
        btn_enviar.config(state="disabled", text=" EN CURSO...", bg=BUTTON_INACTIVE, fg=TEXT_SECONDARY)
        try:
            btn_retry_leads.config(state="disabled", bg=BUTTON_INACTIVE, fg=TEXT_SECONDARY, cursor="arrow")
        except Exception:
            pass
    log_message(f"[INFO] Iniciando ejecución {'programada' if scheduled else 'manual'}: "
                f"{len(market_jobs)} mercado(s) · {len(active_sessions_list)} sesión(es) · "
                f"mercados={'paralelo' if mercados_par else 'consecutivo'} · "
                f"excels={'paralelo' if excels_par else 'consecutivo'}.")

    # Crear modal centrado. El alto se calcula según cuántos mercados hay que mostrar:
    # con alto fijo, a partir del 5º mercado las barras quedaban fuera de la ventana.
    modal = tk.Toplevel(root)
    modal.overrideredirect(True) # Quitar bordes de Windows
    modal_width = 520
    _n_mercados = max(len({_s["pais"] for _s in active_sessions_list}), 1)
    # 260 = encabezado + pastillas + aviso + espacio del resumen final; 46 = alto de cada barra
    modal_height = 260 + 46 * _n_mercados
    modal_height = max(370, min(modal_height, int(root.winfo_screenheight() * 0.85)))
    MODAL_BG = "#231830"
    MODAL_PILL_BG = "#38234D"
    
    # Si la corrida es automática y la app está minimizada (o en la bandeja), hay que
    # traerla de vuelta: si no, Windows reporta la posición de la ventana en -32000 y
    # el modal se crea FUERA DE PANTALLA — parecía que no aparecía nunca.
    if scheduled:
        try:
            if root.state() in ("iconic", "withdrawn") or not root.winfo_viewable():
                try:
                    # Si estaba en la bandeja del sistema hay que sacarla de ahí
                    # (además de restaurarla) o el ícono queda huérfano.
                    _restore_from_tray()
                except Exception:
                    root.deiconify()
                    root.state("normal")
                    _forzar_foreground(root)
                root.update_idletasks()
        except Exception:
            pass

    # Centrar relativo a la app; si sus coordenadas no son válidas (minimizada), se
    # centra sobre la pantalla para que el modal siempre quede visible.
    px = root.winfo_rootx() + (root.winfo_width() - modal_width) // 2
    py = root.winfo_rooty() + (root.winfo_height() - modal_height) // 2
    _sw, _sh = root.winfo_screenwidth(), root.winfo_screenheight()
    if not (0 <= px <= _sw - 50) or not (0 <= py <= _sh - 50):
        px = max(0, (_sw - modal_width) // 2)
        py = max(0, (_sh - modal_height) // 2)
    modal.geometry(f"{modal_width}x{modal_height}+{px}+{py}")
    modal.configure(bg=MODAL_BG, bd=1, highlightthickness=1, highlightbackground=BORDER_COLOR)

    # Al cerrar el modal por CUALQUIER motivo, el botón vuelve a su estado original.
    def _reset_exec_btn(_e=None):
        if scheduled:
            return
        try:
            btn_enviar.config(state="normal", text=" EJECUTAR ENVÍO",
                              bg=EXECUTE_BG, fg=EXECUTE_FG, cursor="hand2")
            refresh_execute_state()
            btn_retry_leads.config(state="normal", bg="#3D2E1A", fg="#F8C471", cursor="hand2")
        except Exception:
            pass
    modal.bind("<Destroy>", lambda e: _reset_exec_btn() if str(e.widget) == str(modal) else None, add="+")

    def on_cerrar():
        _reset_exec_btn()
        modal.destroy()  # el handler <Destroy> restaura el botón Ejecutar
        log_message("[INFO] Ventana de ejecución cerrada.")

    # Modal real: bloquea la interacción con la interfaz de atrás mientras se ejecuta
    modal.transient(root)
    if scheduled:
        # Corrida desatendida: nadie clickeó nada para "ganarse" el foreground de
        # Windows, así que un simple lift()/transient() puede dejar el modal
        # perfectamente creado pero tapado por cualquier otra ventana. Se fuerza.
        _forzar_foreground(modal)
    else:
        modal.attributes("-topmost", False)
        modal.lift()
    # No usamos grab_set() para permitir que el usuario minimice o cierre la ventana principal desde la barra de título de Windows.
    # En su lugar, deshabilitamos la interacción con el área cliente del main window agregando BlockTag a sus widgets.
    try:
        modal.focus_set()
    except Exception:
        pass

    def set_event_blocking(parent, block):
        for child in parent.winfo_children():
            if child == modal or str(child).startswith(str(modal)):
                continue
            tags = list(child.bindtags())
            if block:
                if "BlockTag" not in tags:
                    child.bindtags(("BlockTag",) + tuple(tags))
            else:
                if "BlockTag" in tags:
                    new_tags = tuple(t for t in tags if t != "BlockTag")
                    child.bindtags(new_tags)
            set_event_blocking(child, block)

    def block_evt(e):
        if modal.winfo_exists():
            modal.lift()
            modal.focus_set()
        return "break"

    root.bind_class("BlockTag", "<Button-1>", block_evt)
    root.bind_class("BlockTag", "<ButtonRelease-1>", lambda e: "break")
    root.bind_class("BlockTag", "<Double-Button-1>", lambda e: "break")
    root.bind_class("BlockTag", "<B1-Motion>", lambda e: "break")
    root.bind_class("BlockTag", "<Enter>", lambda e: "break")
    root.bind_class("BlockTag", "<Leave>", lambda e: "break")
    root.bind_class("BlockTag", "<Motion>", lambda e: "break")
    root.bind_class("BlockTag", "<Key>", lambda e: "break")
    root.bind_class("BlockTag", "<FocusIn>", lambda e: block_evt(e))
    set_event_blocking(root, True)

    def on_close_modal():
        if not modal.winfo_exists():
            return
        if "Ejecutando" in title_lbl.cget("text"):
            if messagebox.askyesno("Detener ejecución", "¿Querés detener la ejecución y cerrar la ventana?"):
                on_detener()
                _reset_exec_btn()
                modal.destroy()
        else:
            on_cerrar()

    def on_root_close_request():
        if modal.winfo_exists() and "Ejecutando" in title_lbl.cget("text"):
            if messagebox.askyesno("Salir", "Hay un test en ejecución. ¿Querés detenerlo y salir de la app?"):
                on_detener()
                modal.destroy()
                root.destroy()
        else:
            if modal.winfo_exists():
                modal.destroy()
            root.destroy()

    orig_close_protocol = root.protocol("WM_DELETE_WINDOW")
    root.protocol("WM_DELETE_WINDOW", on_root_close_request)



    unmap_id = root.bind("<Unmap>", lambda e: modal.withdraw() if (e.widget == root and modal.winfo_exists()) else None, add="+")
    map_id = root.bind("<Map>", lambda e: (modal.deiconify(), modal.lift()) if (e.widget == root and modal.winfo_exists()) else None, add="+")
    focus_id = root.bind("<FocusIn>", lambda e: modal.lift() if (modal.winfo_exists() and e.widget.winfo_toplevel() == root and e.widget != modal and not str(e.widget).startswith(str(modal))) else None, add="+")

    def cleanup_root_binds(e=None):
        if e and str(e.widget) != str(modal):
            return
        try:
            root.unbind("<Unmap>", unmap_id)
            root.unbind("<Map>", map_id)
            root.unbind("<FocusIn>", focus_id)
            root.protocol("WM_DELETE_WINDOW", orig_close_protocol)
            set_event_blocking(root, False)
        except Exception:
            pass
    modal.bind("<Destroy>", cleanup_root_binds)

    # Custom Title Bar for minimizing and closing
    title_bar = tk.Frame(modal, bg=MODAL_BG)
    title_bar.pack(fill="x", side="top", padx=15, pady=(5, 0))
    
    title_lbl_bar = tk.Label(title_bar, text="Ejecución de Test", font=("Segoe UI", 8, "bold"), bg=MODAL_BG, fg="#C5A9DF")
    title_lbl_bar.pack(side="left")

    # Hacer el modal arrastrable/movible
    def _start_drag(event):
        modal._drag_start_x = event.x
        modal._drag_start_y = event.y

    def _drag(event):
        x = modal.winfo_x() - modal._drag_start_x + event.x
        y = modal.winfo_y() - modal._drag_start_y + event.y
        modal.geometry(f"+{x}+{y}")

    title_bar.bind("<Button-1>", _start_drag)
    title_bar.bind("<B1-Motion>", _drag)
    title_lbl_bar.bind("<Button-1>", _start_drag)
    title_lbl_bar.bind("<B1-Motion>", _drag)
    
    btn_cls = tk.Button(title_bar, text="✕", font=("Segoe UI", 8, "bold"), bg=MODAL_BG, fg="#C5A9DF", relief="flat", bd=0, cursor="hand2", padx=6, pady=2, command=on_close_modal)
    btn_cls.pack(side="right")
    btn_cls.bind("<Enter>", lambda e: btn_cls.config(bg="#E74C3C", fg="white"))
    btn_cls.bind("<Leave>", lambda e: btn_cls.config(bg=MODAL_BG, fg="#C5A9DF"))

    btn_min = tk.Button(title_bar, text="—", font=("Segoe UI", 8, "bold"), bg=MODAL_BG, fg="#C5A9DF", relief="flat", bd=0, cursor="hand2", padx=6, pady=2, command=root.iconify)
    btn_min.pack(side="right", padx=2)
    btn_min.bind("<Enter>", lambda e: btn_min.config(bg="#38234D"))
    btn_min.bind("<Leave>", lambda e: btn_min.config(bg=MODAL_BG))

    # 1. Header (Ejecutando / Completo)
    header_frame = tk.Frame(modal, bg=MODAL_BG)
    header_frame.pack(fill="x", padx=20, pady=(5, 10))
    
    icon_lbl = tk.Label(header_frame, text="↻", font=("Segoe UI", 16, "bold"), bg=MODAL_BG, fg="#C5A9DF")
    icon_lbl.pack(side="left")
    
    # Animación de rotación del icono
    rotation_glyphs = ["↻", "➔", "↻", "➔"]
    def rotate_icon(idx=0):
        if modal.winfo_exists() and "Ejecutando" in title_lbl.cget("text"):
            icon_lbl.config(text=rotation_glyphs[idx % len(rotation_glyphs)])
            modal.after(250, lambda: rotate_icon(idx + 1))

    title_info = tk.Frame(header_frame, bg=MODAL_BG)
    title_info.pack(side="left", padx=10)

    title_lbl = tk.Label(title_info, text="Ejecutando...", font=("Segoe UI", 12, "bold"), bg=MODAL_BG, fg="white")
    title_lbl.pack(anchor="w")
    subtitle_lbl = tk.Label(title_info, text=f"Preparando… 0/{total_leads} forms", font=("Segoe UI", 9), bg=MODAL_BG, fg=TEXT_SECONDARY)
    subtitle_lbl.pack(anchor="w")
    rotate_icon()
    
    # Botón Detener: mismo comportamiento que la X de cerrar — pide confirmación antes
    # de cortar. on_close_modal ya maneja el popup de confirmación mientras se ejecuta.
    def on_detener():
        stop_event.set()
        btn_detener.config(state="disabled", text=" Deteniendo...")

    def on_detener_click():
        on_close_modal()
        try:
            run_note.config(text="Deteniendo… (termina el lead en curso)", fg="#F8C471")
        except Exception:
            pass
        log_message("[WARN] Detención solicitada por el usuario.")

    btn_detener = tk.Button(header_frame, text=" Detener", image=get_button_icon("stop_coral.png"), compound="left",
                            font=("Segoe UI", 9, "bold"),
                            bg="#3D1220", fg="#F1948A", relief="flat", bd=0, highlightthickness=1,
                            highlightbackground="#F1948A", cursor="hand2", command=on_detener_click, padx=12, pady=4)
    btn_detener.pack(side="right")
    btn_detener.bind("<Enter>", lambda e: btn_detener.config(bg="#5E1D31") if btn_detener["state"] == "normal" else None)
    btn_detener.bind("<Leave>", lambda e: btn_detener.config(bg="#3D1220") if btn_detener["state"] == "normal" else None)

    # 2. Badges Row
    badges_row = tk.Frame(modal, bg=MODAL_BG)
    badges_row.pack(fill="x", padx=20, pady=5)

    def make_pill(parent, text, icon=None):
        img = get_button_icon(icon) if icon else None
        lbl = tk.Label(parent, text=text, font=("Segoe UI", 8, "bold"), bg=MODAL_PILL_BG, fg="white", padx=8, pady=3,
                       bd=0, highlightthickness=1, highlightbackground=BORDER_COLOR)
        if img:
            lbl.config(image=img, compound="left")
        lbl.pack(side="left", padx=3)
        return lbl

    if url_par:
        make_pill(badges_row, f" Por URL: Paralelo (máx {url_max})", icon="link_lav.png")
        make_pill(badges_row, " 1 navegador por URL", icon="gear_lav.png")
    else:
        make_pill(badges_row, f" Mercados: {'Paralelo' if mercados_par else 'Secuencial'}", icon="link_lav.png")
        make_pill(badges_row, f" Excels: {'Paralelo' if excels_par else 'Secuencial'}", icon="gear_lav.png")
    _n_paises = len({_s["pais"] for _s in active_sessions_list})
    make_pill(badges_row, f" {_n_paises} mercado(s) · {total_leads} forms", icon="monitor_lav.png")

    # Aviso durante la ejecución (se quita al completar)
    run_note = tk.Label(modal, text="⚠ No podés cerrar esta ventana mientras se ejecuta. Para correr otro test ahora, abrí otra ventana de la app.",
                        font=("Segoe UI", 8, "italic"), bg=MODAL_BG, fg="#F8C471", wraplength=480, justify="left")
    run_note.pack(anchor="w", padx=20, pady=(4, 8))

    def _ui(fn):
        try:
            root.after(0, fn)
        except Exception:
            pass

    # 3. Una barra de progreso por MERCADO (país). El nombre aparece una sola vez
    #    arriba de su barra; la barra se llena a medida que avanza.
    from collections import OrderedDict as _OrderedDict
    _pais_totals = _OrderedDict()        # sesiones (dispositivos) por país
    _pais_forms_total = {}               # formularios (leads) por país = suma de rows_count
    _pais_devices = {}
    for _s in active_sessions_list:
        _pais_totals[_s["pais"]] = _pais_totals.get(_s["pais"], 0) + 1
        _pais_forms_total[_s["pais"]] = _pais_forms_total.get(_s["pais"], 0) + int(_s.get("rows_count", 1) or 1)

        # Formatear el nombre del dispositivo para mostrar
        dev_name = _s.get("device") or ""
        if "·URL" in dev_name:
            dev_name = dev_name.split("·URL")[0]
        if dev_name == "Mac":
            dev_name = "Mac LT"
        elif dev_name == "Android":
            dev_name = "Android LT"
            
        if _s["pais"] not in _pais_devices:
            _pais_devices[_s["pais"]] = []
        if dev_name and dev_name not in _pais_devices[_s["pais"]]:
            _pais_devices[_s["pais"]].append(dev_name)

    # Área de mercados con scroll: si no entran todas las barras (muchos mercados en
    # una pantalla chica), aparece la barra lateral en vez de recortarse.
    markets_wrap = tk.Frame(modal, bg=MODAL_BG)
    markets_wrap.pack(fill="both", expand=True, padx=20, pady=(2, 8))
    markets_canvas = tk.Canvas(markets_wrap, bg=MODAL_BG, highlightthickness=0)
    markets_sb = tk.Scrollbar(markets_wrap, orient="vertical", command=markets_canvas.yview)
    markets_canvas.configure(yscrollcommand=markets_sb.set)
    markets_canvas.pack(side="left", fill="both", expand=True)
    markets_frame = tk.Frame(markets_canvas, bg=MODAL_BG)
    _mk_win = markets_canvas.create_window((0, 0), window=markets_frame, anchor="nw")

    def _sync_markets_scroll(_e=None):
        try:
            markets_canvas.configure(scrollregion=markets_canvas.bbox("all"))
            markets_canvas.itemconfig(_mk_win, width=markets_canvas.winfo_width())
            hace_falta = markets_frame.winfo_reqheight() > markets_canvas.winfo_height()
            if hace_falta and not markets_sb.winfo_ismapped():
                markets_sb.pack(side="right", fill="y")
            elif not hace_falta and markets_sb.winfo_ismapped():
                markets_sb.pack_forget()
        except Exception:
            pass

    markets_frame.bind("<Configure>", _sync_markets_scroll)
    markets_canvas.bind("<Configure>", _sync_markets_scroll)
    markets_canvas.bind("<MouseWheel>",
                        lambda e: markets_canvas.yview_scroll(int(-e.delta / 120), "units"))

    _pais_bars = {}
    for _p, _tot in _pais_totals.items():
        _row = tk.Frame(markets_frame, bg=MODAL_BG)
        _row.pack(fill="x", pady=(0, 9))
        _hdr = tk.Frame(_row, bg=MODAL_BG)
        _hdr.pack(fill="x")
        
        # Formatear el texto de dispositivos
        devs_list = _pais_devices.get(_p, [])
        devs_str = " — " + " / ".join(devs_list) if devs_list else ""
        display_name = f"{_p}{devs_str}"
        
        _forms_tot = int(_pais_forms_total.get(_p, _tot) or _tot)
        tk.Label(_hdr, text=display_name, font=("Segoe UI", 10, "bold"), bg=MODAL_BG, fg="white").pack(side="left")
        _stx = tk.Label(_hdr, text=f"0/{_forms_tot} forms", font=("Segoe UI", 8), bg=MODAL_BG, fg=TEXT_SECONDARY)
        _stx.pack(side="right")
        _cb = tk.Canvas(_row, height=8, bg="#35164D", highlightthickness=0)
        _cb.pack(fill="x", pady=(3, 0))
        _fl = _cb.create_rectangle(0, 0, 0, 8, fill="#F8C471", width=0)
        # total     = sesiones (para saber cuándo el país terminó todas sus sesiones)
        # forms_tot = formularios reales (lo que se muestra y lo que va al email)
        _pais_bars[_p] = {"canvas": _cb, "fill": _fl, "status": _stx,
                          "total": _tot, "done": 0,
                          "forms_total": _forms_tot, "forms_ok": 0, "forms_fail": 0,
                          "fail_rows": []}

    # Progreso real por formulario: cada sesión reporta cuántos leads lleva hechos.
    _sess_frac = {}          # fracción de la sesión (para animar la barra en vivo)
    _sess_forms_done = {}    # nº de formularios completados por sesión (entero)
    _pais_of_sess = {_s["sess_id"]: _s["pais"] for _s in active_sessions_list}

    def _forms_done_pais(pais):
        ids = [sid for sid, pp in _pais_of_sess.items() if pp == pais]
        return sum(int(_sess_forms_done.get(sid, 0)) for sid in ids)

    def _paint_pais(pais):
        b = _pais_bars.get(pais)
        if not b or not b["canvas"].winfo_exists():
            return
        with _lock:
            ftot = max(1, b["forms_total"])
            fdone = min(_forms_done_pais(pais), ftot)
            frac = fdone / ftot
            sdone, stot = b["done"], b["total"]
            okc, failc = b["forms_ok"], b["forms_fail"]
        w = int(max(1, b["canvas"].winfo_width()) * max(0.0, min(1.0, frac)))
        b["canvas"].coords(b["fill"], 0, 0, w, 8)
        if sdone >= stot:
            col = "#F1948A" if failc else "#82E0AA"
            b["canvas"].itemconfig(b["fill"], fill=col)
            b["status"].config(text=f"✓ {okc} OK · {failc} error(es)", fg=col)
        else:
            b["status"].config(text=f"{fdone}/{b['forms_total']} forms", fg=TEXT_SECONDARY)

    _dev_label = {"desktop": None, "mac": "Mac LT (Safari)", "android": "Android LT"}

    def show_completed(ok_total, fail_total, detenido, err_msg):
        _exec_state["running"] = False
        if not modal.winfo_exists():
            return
        if detenido:
            icon_lbl.config(text="■", fg="#F1948A", font=("Segoe UI", 16, "bold"))
            title_lbl.config(text="Ejecución detenida")
        elif err_msg:
            icon_lbl.config(text="✕", fg="#F1948A", font=("Segoe UI", 16, "bold"))
            title_lbl.config(text="Ejecución con error")
        else:
            icon_lbl.config(text="✓", fg="#82E0AA", font=("Segoe UI", 18, "bold"))
            title_lbl.config(text="Ejecución completada")
        try:
            btn_detener.pack_forget()
        except Exception:
            pass
        try:
            run_note.destroy()
        except Exception:
            pass

        # Veredicto global + subtítulo con el total de formularios OK / con error.
        _global_ok = (fail_total == 0 and not err_msg and not detenido)
        try:
            subtitle_lbl.config(text=f"{ok_total + fail_total}/{total_leads} forms  ·  "
                                     f"{ok_total} OK · {fail_total} con error")
        except Exception:
            pass
        try:
            if detenido:
                _v_txt, _v_bg = "⏹  DETENIDO POR EL USUARIO", "#B9770E"
            elif _global_ok:
                _v_txt, _v_bg = "✓  RESULTADO GLOBAL: PASS", "#1E8449"
            else:
                _v_txt, _v_bg = f"✕  RESULTADO GLOBAL: FAIL  ({fail_total} form/s con error)", "#943126"
            tk.Label(modal, text=_v_txt, font=("Segoe UI", 10, "bold"),
                     bg=_v_bg, fg="white", padx=12, pady=6).pack(fill="x", padx=20, pady=(6, 4))
        except Exception:
            pass

        # Completar todas las barras por mercado (contando FORMULARIOS)
        _fail_detail_lines = []  # ("País", "fila 3, fila 7")
        for _p, b in _pais_bars.items():
            if not b["canvas"].winfo_exists():
                continue
            _col = "#F1948A" if (b["forms_fail"] or err_msg or detenido) else "#82E0AA"
            _wfull = max(1, b["canvas"].winfo_width())
            b["canvas"].itemconfig(b["fill"], fill=_col)
            b["canvas"].coords(b["fill"], 0, 0, _wfull, 8)
            # PASS/FAIL explicito por mercado: antes decia "✓ 7 OK · 1 error(es)" y el
            # tilde daba a entender que habia salido todo bien aunque hubiera fallas.
            _es_pass = not (b["forms_fail"] or err_msg or detenido)
            _pf = "PASS" if _es_pass else "FAIL"
            _mark = "✓" if _es_pass else "✕"
            b["status"].config(
                text=f"{_mark} {_pf} — {b['forms_ok']} OK · {b['forms_fail']} error(es)", fg=_col)
            if b.get("fail_rows"):
                for _fr in b["fail_rows"]:
                    if isinstance(_fr, dict):
                        _rn = _fr.get("row", "?")
                        _rs = _fr.get("reason", "")
                        # URL suelta: la landing viene vacia y la URL util es la del form
                        _u = (_fr.get("url") or _fr.get("url_form") or "").strip()
                    else:
                        _rn, _rs, _u = _fr, "", ""
                    _fail_detail_lines.append((_p, _rn, _rs, _u))

        # Detalle de los forms con error: fila, motivo y URL, para ubicarlos sin
        # tener que abrir el Excel.
        if _fail_detail_lines:
            _det = tk.Frame(modal, bg=MODAL_BG)
            _det.pack(fill="x", padx=20, pady=(2, 0))
            tk.Label(_det, text="Forms con error (FAIL):", font=("Segoe UI", 8, "bold"),
                     bg=MODAL_BG, fg="#F1948A").pack(anchor="w")
            for _p, _rn, _rs, _u in _fail_detail_lines:
                _txt = f"   • {_p} — fila {_rn}"
                if _rs:
                    _txt += f": {_rs}"
                tk.Label(_det, text=_txt, font=("Segoe UI", 8),
                         bg=MODAL_BG, fg="#F1948A", wraplength=470, justify="left").pack(anchor="w")
                if _u:
                    tk.Label(_det, text=f"       {_u}", font=("Segoe UI", 7),
                             bg=MODAL_BG, fg="#D98880", wraplength=460, justify="left").pack(anchor="w")

        if err_msg:
            tk.Label(modal, text=f"✕ {err_msg}", font=("Segoe UI", 8), bg=MODAL_BG, fg="#F1948A",
                     wraplength=480, justify="left").pack(anchor="w", padx=20, pady=(2, 0))

        if scheduled:
            tk.Label(modal, text="✓ Esta ventana se cierra sola en unos segundos. Los tests programados posteriores se ejecutarán igual.",
                     font=("Segoe UI", 8, "italic"), bg=MODAL_BG, fg="#82E0AA", wraplength=480, justify="left").pack(anchor="w", padx=20, pady=(2, 0))
            # El modal bloquea toda la interfaz (BlockTag) y sólo se libera al
            # destruirse. En una corrida automática no hay nadie para cerrarlo, así
            # que la app quedaba inutilizable hasta que alguien la mirara: se cierra
            # solo y devuelve el control. El resultado ya quedó en el Excel y el mail.
            def _autocerrar_modal():
                try:
                    if modal.winfo_exists():
                        modal.destroy()  # el handler <Destroy> saca el BlockTag
                except Exception:
                    pass
            root.after(20000, _autocerrar_modal)

        # Banner de email (el backend encola el envío si "Enviar mail" está activo)
        if enviar_mail and not detenido:
            if dest:
                _bg, _fg, _tx = "#1F3A30", "#82E0AA", f"✉  Email de resultados encolado a: {dest}"
            else:
                _bg, _fg, _tx = "#3A1F22", "#F1948A", "⚠  Falta el destinatario: no se envió email."
        else:
            _bg, _fg, _tx = BUTTON_INACTIVE, TEXT_SECONDARY, "✉  Envío de email desactivado."
        _eb = tk.Frame(modal, bg=_bg, bd=0, highlightthickness=1, highlightbackground=_fg)
        _eb.pack(fill="x", padx=20, pady=5)
        tk.Label(_eb, text=_tx, font=("Segoe UI", 9, "bold"), bg=_bg, fg=_fg, pady=4, wraplength=475, justify="center").pack(anchor="center")

        summary_row = tk.Frame(modal, bg=MODAL_BG)
        summary_row.pack(fill="x", padx=20, pady=5)
        tk.Label(summary_row, text=f"🟢 {ok_total} OK      🔴 {fail_total} con error", font=("Segoe UI", 9, "bold"),
                 bg=MODAL_BG, fg="white").pack(side="left")



        btn_close = tk.Button(modal, text="Cerrar resultados", font=("Segoe UI", 10, "bold"), bg="#AED6F1", fg="#110518",
                              relief="flat", bd=0, cursor="hand2", command=on_cerrar, pady=6)
        btn_close.pack(fill="x", padx=20, pady=(15, 10))
        btn_close.bind("<Enter>", lambda e: btn_close.config(bg="#D4E6F1"))
        btn_close.bind("<Leave>", lambda e: btn_close.config(bg="#AED6F1"))

    # Estado compartido entre sesiones (thread-safe)
    _lock = threading.Lock()
    # ok/fail acá cuentan FORMULARIOS (no sesiones); done cuenta sesiones terminadas.
    _st = {"ok": 0, "fail": 0, "done": 0, "err": ""}
    _running_markets = set()   # mercados que están ejecutando ahora mismo
    _email_results = []  # entradas para el email consolidado / por país
    _email_lock = threading.Lock()

    def _subtitle_text():
        with _lock:
            gdone = min(sum(int(v) for v in _sess_forms_done.values()), total_leads)
            run = [p for p in _pais_totals if p in _running_markets]
        run_str = ", ".join(run) if run else "—"
        return f"Ejecutando: {run_str}  ·  {gdone}/{total_leads} forms"
    # LambdaTest Android suele permitir 1 sesión concurrente (device real): serializamos
    # SOLO las sesiones Android para que TODOS los países se ejecuten (en cola), sin que
    # una quede afuera. El resto (desktop / Mac) sigue en paralelo.
    _android_sem = threading.Semaphore(1)

    def _bump(pais, sess_id=None, ok=0, fail=0, err="", fail_rows=None):
        """Cierra una sesión. ok/fail son FORMULARIOS de esa sesión; fail_rows son los
        números de fila (del Excel) que fallaron, para poder ubicarlos."""
        with _lock:
            _st["ok"] += ok
            _st["fail"] += fail
            _st["done"] += 1
            if err:
                _st["err"] = err
            if sess_id is not None:
                _sess_frac[sess_id] = 1.0  # sesión terminada = barra de esa sesión llena
                _sess_forms_done[sess_id] = ok + fail  # forms procesados por esta sesión
            b = _pais_bars.get(pais)
            if b:
                b["done"] += 1
                b["forms_ok"] += ok
                b["forms_fail"] += fail
                if fail_rows:
                    b["fail_rows"].extend(fail_rows)
                if b["done"] >= b["total"]:
                    _running_markets.discard(pais)
        def _u():
            if modal.winfo_exists():
                subtitle_lbl.config(text=_subtitle_text())
                _paint_pais(pais)
        _ui(_u)

    def _record_retry_fails(pais, device, sess, excel, scheduled, frows):
        """Guarda/actualiza el estado para 'REINTENTAR FALLIDOS' de esta sesión —
        mismo criterio que ya usa la rama desktop: sólo corridas manuales, no las
        sub-sesiones de 'una sesión por URL', traduce fila-temporal -> fila-original
        vía _retry_row_map cuando el reintento vino de un reintento anterior."""
        if scheduled or "·URL" in device:
            return
        _orig_excel = sess.get("_retry_source_excel") or excel
        _row_map = sess.get("_retry_row_map") or {}
        _orig_rows = set()
        for _fr in frows:
            _rn = _fr.get("row") if isinstance(_fr, dict) else _fr
            try:
                _rn = int(_rn)
            except (TypeError, ValueError):
                continue
            _orig_rows.add(_row_map.get(_rn, _rn))
        with _lock:
            _lf = _exec_state.setdefault("last_fails", {})
            if _orig_rows:
                _lf[(pais, device)] = {"excel": _orig_excel, "rows": sorted(_orig_rows)}
            else:
                _lf.pop((pais, device), None)

    def _collect_lt_email(pais, navegador, viewport, summary):
        """Registra el resultado LT para email y, si es modo por país, lo envía ya."""
        if not (enviar_mail and not stop_event.is_set()):
            return
        _rp = summary.get("results_excel") if summary else None
        if not _rp:
            return
        _entry = {"pais": pais, "navegador": navegador, "viewport": viewport,
                  "estado": "completado", "excel_path": _rp, "screenshots_dir": None}
        with _email_lock:
            _email_results.append(_entry)
        if _email_modo == "por_pais":
            try:
                from osocio.interface.helpers_interface import enviar_email_resultados_consolidados
                enviar_email_resultados_consolidados([_entry])
            except Exception as _e:
                log_message(f"[ERROR] email LT {pais}: {_e}")

    def _run_session(sess):
        """Corre una sesión (un Excel de un mercado en un dispositivo)."""
        if stop_event.is_set():
            return
        pais, dtype, browser, device, excel = sess["pais"], sess["dtype"], sess["browser"], sess["device"], sess["excel"]

        _sid = sess["sess_id"]

        with _lock:
            _running_markets.add(pais)

        def _set_cur():
            if modal.winfo_exists():
                title_lbl.config(text="Ejecutando...")
                subtitle_lbl.config(text=_subtitle_text())
        _ui(_set_cur)

        try:
            if dtype == "desktop":
                from osocio.core.generic_country_base import GenericCountryBase
                _pausar = var_pausar_autenticacion.get()
                _preview = bool(var_preview_navegador.get())
                # La variante tiene que viajar como tal, no solo como ruta de Excel:
                # de ella dependen en que carpeta caen los resultados (t1/ o t3/) y
                # como se llama el archivo. Pisar EXCEL_PATH a secas alcanzaba para
                # LEER el Excel correcto, pero dejaba al motor creyendo que era una
                # corrida normal: los T3 lanzados desde la app escribian en t1/.
                form = GenericCountryBase(pais, browser=browser, viewport="fullscreen",
                                          headless=False, background=background, is_scheduled=scheduled,
                                          pausar_autenticacion=_pausar,
                                          preview_visible_browser=_preview,
                                          excel_suffix=sess.get("variante", ""))

                if excel:
                    form.EXCEL_PATH = excel  # ← una sesión por Excel generado
                def _pcb(done, total):
                    with _lock:
                        _sess_frac[_sid] = (done / total) if total else 0.0
                        _sess_forms_done[_sid] = int(done)   # forms reales completados
                    def _u():
                        if modal.winfo_exists():
                            _paint_pais(pais)
                            subtitle_lbl.config(text=_subtitle_text())
                    _ui(_u)
                form.run(progress_callback=_pcb)
                if enviar_mail and not stop_event.is_set():
                    rp = getattr(form, "RESULTADOS_PATH", None)
                    sd = getattr(form, "SCREENSHOT_DIR", None)
                    if rp:
                        _entry = {"pais": sess.get("email_label") or pais,
                                  "navegador": browser, "viewport": "fullscreen",
                                  "estado": "completado", "excel_path": rp, "screenshots_dir": sd}
                        with _email_lock:
                            _email_results.append(_entry)
                        if _email_modo == "por_pais":
                            from osocio.interface.helpers_interface import enviar_email_resultados
                            enviar_email_resultados(pais, rp, sd, browser=browser, viewport="fullscreen")
                # Resumen autoritativo por formulario (mismo criterio que colorea el Excel).
                _summ = getattr(form, "run_summary", None) or {}
                _ok = int(_summ.get("ok", 0))
                _fail = int(_summ.get("fail", 0))
                _frows = list(_summ.get("fail_rows", []) or [])
                if _ok == 0 and _fail == 0:
                    # Excel sin filas procesadas: contar la sesión como 1 form OK para no romper totales.
                    _ok = int(sess.get("rows_count", 1) or 1)
                _bump(pais, _sid, ok=_ok, fail=_fail, fail_rows=_frows)
                _record_retry_fails(pais, device, sess, excel, scheduled, _frows)
            elif dtype == "mac":
                from osocio.providers.lambdatest_mac import lt_controller
                b_name = f"Osocio Automatizado LT MAC - {pais}" if scheduled else f"Osocio LT Mac - Envío Manual - {pais}"
                summary = lt_controller.run(pais=pais, build_name=b_name, excel_path=excel, log_fn=log_message) or {}
                _collect_lt_email(pais, "lambdatest_mac", "mac", summary)
                _err = summary.get("error")
                _ok, _fail = int(summary.get("ok", 0)), int(summary.get("failed", 0))
                _lt_frows = list(summary.get("fail_rows", []) or [])
                if _err and _ok == 0 and _fail == 0:
                    _bump(pais, _sid, fail=1, err=str(_err)[:200])
                else:
                    _bump(pais, _sid, ok=_ok, fail=_fail, fail_rows=_lt_frows)
                _record_retry_fails(pais, device, sess, excel, scheduled, _lt_frows)
            elif dtype == "android":
                from osocio.providers.lambdatest_android import lt_android_controller
                with _android_sem:  # 1 sesión Android a la vez → todos los países corren
                    if stop_event.is_set():
                        return
                    b_name = f"Osocio Automatizado LT ANDROID - {pais}" if scheduled else f"Osocio LT Android - Envío Manual - {pais}"
                    summary = lt_android_controller.run(pais=pais, build_name=b_name, excel_path=excel, log_fn=log_message) or {}
                _collect_lt_email(pais, "lambdatest_android", "android", summary)
                _err = summary.get("error")
                _ok, _fail = int(summary.get("ok", 0)), int(summary.get("failed", 0))
                _lt_frows = list(summary.get("fail_rows", []) or [])
                if _err and _ok == 0 and _fail == 0:
                    _bump(pais, _sid, fail=1, err=str(_err)[:200])
                else:
                    _bump(pais, _sid, ok=_ok, fail=_fail, fail_rows=_lt_frows)
                _record_retry_fails(pais, device, sess, excel, scheduled, _lt_frows)
        except Exception as e:
            log_message(f"[ERROR] {pais}/{device}: {e}")
            # Crash de la sesión: contar como fallidos todos sus formularios.
            _bump(pais, _sid, fail=int(sess.get("rows_count", 1) or 1), err=str(e)[:200])

    def _run_market(sessions):
        """Corre las sesiones de un mercado. Si mercados_par o excels_par están
        activos, lanza todos los dispositivos/Excels del mercado en paralelo
        (Chrome local y LambdaTest no compiten, corren en máquinas distintas)."""
        _parallel = (excels_par or mercados_par) and len(sessions) > 1
        if _parallel:
            ts = [threading.Thread(target=_run_session, args=(s,), daemon=True) for s in sessions]
            for t in ts:
                t.start()
            for t in ts:
                t.join()
        else:
            for s in sessions:
                if stop_event.is_set():
                    break
                _run_session(s)

    def _worker():
        _ensure_serialized_setup()  # evita choque de resultados en paralelo
        if url_par:
            # Una sesión por URL, todas en paralelo con tope de concurrencia (url_max)
            sem = threading.Semaphore(url_max)

            def _guarded(s):
                if stop_event.is_set():
                    return
                with sem:
                    _run_session(s)

            ts = [threading.Thread(target=_guarded, args=(s,), daemon=True) for s in flat_sessions]
            for t in ts:
                t.start()
            for t in ts:
                t.join()
        elif mercados_par:
            ts = [threading.Thread(target=_run_market, args=(js,), daemon=True) for _, js in market_jobs]
            for t in ts:
                t.start()
            for t in ts:
                t.join()
        else:
            for _pais, js in market_jobs:
                if stop_event.is_set():
                    break
                _run_market(js)
        _detenido = stop_event.is_set()
        # Modo consolidado: un único email al terminar TODOS los mercados/dispositivos.
        if enviar_mail and _email_modo == "consolidado" and _email_results and not _detenido:
            try:
                from osocio.interface.helpers_interface import enviar_email_resultados_consolidados
                enviar_email_resultados_consolidados(list(_email_results))
            except Exception as _e:
                log_message(f"[ERROR] email consolidado: {_e}")
        _ui(lambda: show_completed(_st["ok"], _st["fail"], _detenido, _st["err"]))

    _exec_state["running"] = True
    threading.Thread(target=_worker, daemon=True).start()
