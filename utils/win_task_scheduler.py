"""
win_task_scheduler.py — Registro del envío de leads en el Programador de tareas de Windows.

La app sólo dispara los horarios mientras está abierta. Para que el envío corra con la
app cerrada se registra una tarea diaria de Windows por cada horario, que lanza
`run.py --autonomous --once` (ejecución única, sin loop de espera).

schtasks no admite varios horarios fijos en una sola tarea diaria, así que se crea una
tarea por hora: "Osocio - Envio de Leads 09h", "... 12h", etc.
"""
import os
import re
import subprocess
import sys

TASK_PREFIX = "Osocio - Envio de Leads"

# No abrir ventana de consola al correr schtasks desde la GUI.
_NO_WINDOW = 0x08000000 if os.name == "nt" else 0


def _project_root():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _launch_command(abrir_app=False):
    """Comando que ejecutará Windows. Entre comillas: las rutas llevan espacios.

    abrir_app=False → ejecución silenciosa (`--autonomous --once`): no se abre nada,
                      el envío corre en segundo plano.
    abrir_app=True  → se abre la app y ella misma dispara el envío programado
                      (`--autostart-leads`), aunque estuviera cerrada.
    """
    modo = "--autostart-leads" if abrir_app else "--autonomous --once"
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}" {modo}'
    # pythonw.exe: sin ventana negra de consola. La ventana de la app (Tk) sí se ve.
    exe = sys.executable
    pythonw = os.path.join(os.path.dirname(exe), "pythonw.exe")
    if os.path.isfile(pythonw):
        exe = pythonw
    return f'"{exe}" "{os.path.join(_project_root(), "run.py")}" {modo}'


def _launch_parts(abrir_app=False):
    """Igual que _launch_command pero separado en (ejecutable, argumentos), que es
    lo que necesita Register-ScheduledTask de PowerShell."""
    modo = "--autostart-leads" if abrir_app else "--autonomous --once"
    if getattr(sys, "frozen", False):
        return sys.executable, modo
    exe = sys.executable
    pythonw = os.path.join(os.path.dirname(exe), "pythonw.exe")
    if os.path.isfile(pythonw):
        exe = pythonw
    return exe, f'"{os.path.join(_project_root(), "run.py")}" {modo}'


def _ps_quote(valor):
    """Cita un literal para PowerShell (comilla simple: se escapa duplicándola)."""
    return "'" + str(valor).replace("'", "''") + "'"


def _registrar_ps(nombre, hora, exe, args):
    """Registra la tarea con PowerShell en vez de schtasks.

    schtasks no permite configurar lo que hace falta para que esto sirva de verdad
    en una notebook: despertar la máquina si está suspendida (-WakeToRun) y correr
    igual si el horario se perdió porque estaba apagada/dormida (-StartWhenAvailable).
    Register-ScheduledTask sí.
    """
    script = (
        f"$a = New-ScheduledTaskAction -Execute {_ps_quote(exe)} "
        f"-Argument {_ps_quote(args)} -WorkingDirectory {_ps_quote(_project_root())}; "
        f"$t = New-ScheduledTaskTrigger -Daily -At {_ps_quote(hora)}; "
        "$s = New-ScheduledTaskSettingsSet -WakeToRun -StartWhenAvailable "
        "-AllowStartIfOnBatteries -DontStopIfGoingOnBatteries "
        "-ExecutionTimeLimit (New-TimeSpan -Hours 3) "
        "-MultipleInstances IgnoreNew; "
        f"Register-ScheduledTask -TaskName {_ps_quote(nombre)} -Action $a -Trigger $t "
        "-Settings $s -Force | Out-Null"
    )
    return _run(["powershell", "-NoProfile", "-NonInteractive", "-Command", script])


def task_name_for(hora):
    """"09:00" → "Osocio - Envio de Leads 09h"."""
    return f"{TASK_PREFIX} {hora[:2]}h"


def _run(args):
    return subprocess.run(args, capture_output=True, text=True,
                          creationflags=_NO_WINDOW, cwd=_project_root())


def _horas_validas(horarios):
    """Normaliza y deduplica una lista de "HH:MM"."""
    out = []
    for h in horarios or []:
        h = str(h).strip()
        if re.fullmatch(r"\d{2}:\d{2}", h) and h not in out:
            out.append(h)
    return sorted(out)


def registrar(horarios, abrir_app=False):
    """Crea (o reemplaza) una tarea diaria por horario.

    Devuelve (ok, mensaje). No pide permisos de administrador: las tareas se crean
    para el usuario actual y sólo corren cuando la sesión está iniciada.
    """
    if os.name != "nt":
        return False, "El Programador de tareas sólo está disponible en Windows."

    horas = _horas_validas(horarios)
    if not horas:
        return False, "No hay horarios válidos para registrar."

    exe, args = _launch_parts(abrir_app)
    creadas, errores = [], []
    for hora in horas:
        nombre = task_name_for(hora)
        res = _registrar_ps(nombre, hora, exe, args)
        if res.returncode == 0:
            creadas.append(f"{nombre} ({hora})")
        else:
            errores.append(f"{nombre}: {(res.stderr or res.stdout or '').strip()[:200]}")

    if errores and not creadas:
        return False, "No se pudo registrar ninguna tarea:\n" + "\n".join(errores)
    msg = "Tareas registradas en Windows:\n" + "\n".join("• " + c for c in creadas)
    if errores:
        msg += "\n\nCon errores:\n" + "\n".join("• " + e for e in errores)
    return True, msg


def desregistrar():
    """Elimina todas las tareas del prefijo. Devuelve (ok, mensaje)."""
    if os.name != "nt":
        return False, "El Programador de tareas sólo está disponible en Windows."

    nombres = listar()
    if not nombres:
        return True, "No había tareas registradas."
    borradas, errores = [], []
    for nombre in nombres:
        res = _run(["schtasks", "/Delete", "/TN", nombre, "/F"])
        (borradas if res.returncode == 0 else errores).append(nombre)
    msg = "Tareas eliminadas:\n" + "\n".join("• " + b for b in borradas) if borradas else ""
    if errores:
        msg += "\n\nNo se pudieron eliminar:\n" + "\n".join("• " + e for e in errores)
    return not errores, msg or "Sin cambios."


def listar():
    """Nombres de las tareas registradas por la app (lista vacía si no hay ninguna)."""
    if os.name != "nt":
        return []
    res = _run(["schtasks", "/Query", "/FO", "CSV", "/NH"])
    if res.returncode != 0:
        return []
    nombres = []
    for linea in (res.stdout or "").splitlines():
        campos = linea.split('","')
        if not campos:
            continue
        nombre = campos[0].strip('" ').lstrip("\\")
        if nombre.startswith(TASK_PREFIX) and nombre not in nombres:
            nombres.append(nombre)
    return sorted(nombres)
