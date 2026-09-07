Resumen de cambios — Ocioso_Update
¿Por qué se hicieron estos cambios?
Se compararon los dos proyectos (Form_Automation_Project-pasaje-git-main = original, Ocioso_Update = tu versión mejorada) y se encontró que:

El original tenía algunas cosas que tu versión había perdido en el proceso de mejora
Tu versión tenía dependencias innecesarias que pesaban y descargaban cosas de internet sin necesidad
Había riesgos de seguridad (contraseñas guardadas en git)
Brasil tenía un pequeño error en el generador de documentos de respaldo
1. Email — Se restauró el indicador visual de estado (interface/helpers_interface.py + interface/main_interface.py)
Antes: El email se enviaba por detrás y no se veía nada en la pantalla. Si fallaba, tampoco había forma de saberlo desde la interfaz.

Ahora: En la parte superior derecha de la pantalla, donde está el campo de email, aparece un indicador en tiempo real:

⏳ Enviando email... (mientras está en proceso)
✅ Email enviado correctamente (cuando llegó bien)
❌ Error al enviar: [motivo] (si algo falló)
Esto existía en el proyecto original y se había perdido. Además se mejoró el manejo interno de errores: ahora guarda el error completo (no solo un resumen) en el log para facilitar la solución de problemas.

2. Drivers de Selenium — Eliminado el descargador automático (core/browser_manager.py + requirements.txt)
Antes: La app usaba webdriver-manager, una librería que se conecta a internet para descargar automáticamente el driver de Chrome/Firefox/Edge cuando detecta que el que tenés es viejo.

Ahora: La app solo usa los drivers que están en la carpeta /drivers/ del proyecto. Si el driver está desactualizado, el mensaje de error te indica que lo reemplaces manualmente descargándolo del sitio oficial.

¿Por qué? Descargar ejecutables automáticamente desde internet sin control es un riesgo de seguridad: si el servidor fuente fuera comprometido, el sistema descargaría algo malicioso sin que te des cuenta. Vos ya tenés los drivers ahí, así que no hace falta.

3. Dependencias limpias (requirements.txt)
Antes: Tenía webdriver-manager (ver punto 2) y cpf-and-cnpj-generator, un paquete para generar documentos brasileños que nunca se importaba en ningún archivo del proyecto. Era letra muerta.

Ahora: Solo quedan las dependencias que realmente se usan: selenium, openpyxl, pandas, pillow, pytz, pywin32.

4. Seguridad — Credenciales protegidas de git (.gitignore)
Antes: Los archivos lambdatest_credentials.txt (que tienen tu usuario y clave de LambdaTest) y config_global.json (que también guarda la clave de LambdaTest) podían terminar en git si hacías un commit.

Ahora: Ambos archivos están en .gitignore. Siguen en tu disco y la app los usa normalmente, pero git los ignora y nunca los va a subir al repositorio. Las claves se quedan solo en tu máquina.

5. Brasil — Corregido el generador de respaldo (utils/data_generator.py)
Antes: Si la app intentaba generar un CNPJ o un CEP para Brasil y la API de 4devs fallaba (sin internet, API caída, etc.), el respaldo local devolvía un CPF en lugar de un CNPJ o CEP. Eso habría hecho que el Excel se generara con datos en formato incorrecto para esos campos.

Ahora:

Si la API falla para CPF → genera un CPF válido localmente (algoritmo oficial)
Si la API falla para CNPJ → genera un CNPJ válido localmente (nuevo, generar_cnpj_brasil())
Si la API falla para CEP → usa uno de los CEPs reales conocidos que están guardados en una lista (ej. Av. Paulista, São Paulo)
En todos los casos, guarda un aviso en el log para que sepas que se usó el respaldo
Nota: El motor de llenado de formularios (base_form_filler.py) ya tenía este comportamiento correcto para cuando llena el formulario en tiempo real. Esta corrección es para la generación de Excels desde la interfaz.

6. Comentarios explicativos en todos los archivos Python
Cada archivo .py principal ahora tiene un encabezado en español al inicio que explica en 2-3 líneas qué hace ese archivo. Los archivos de LambdaTest ya los tenían; se agregaron a todos los demás (run.py, autonomous_runner.py, base_form_filler.py, browser_manager.py, helpers_interface.py, main_interface.py, y todos los de utils/ y validation/).

¿Qué NO se tocó?
Todo el llenado de formularios (base_form_filler.py) — tu versión ya era mejor que el original
LambdaTest (lambdatest_mac, lambdatest_android) — funcionan perfecto, no se modificó nada
Las configuraciones de países (country_configs.py) — igual
La lógica de campos no mapeados (ya auto-selecciona aleatoriamente cuando el Excel viene vacío)


Listo. El cambio clave es que antes se ponía --start-maximized como argumento de Chrome (que hace que el browser se abra maximizado en primer plano y roba el foco), y después se intentaba moverlo con set_window_position — pero para ese momento ya había robado el foco.

Ahora se usan --window-position=10000,0 y --window-size=1366,768 directamente como argumentos de inicio, entonces Chrome abre ya posicionado fuera de la pantalla y nunca aparece en primer plano ni molesta mientras escribís.








Resumen completo de cambios
Contexto
Tomamos el proyecto Ocioso_Update (tu versión mejorada) y lo actualizamos para restaurar features del original, eliminar peso innecesario y agregar nuevas funcionalidades.

1. Seguridad
requirements.txt — Eliminados webdriver-manager (descargaba drivers de internet automáticamente, riesgo de supply chain) y cpf-and-cnpj-generator (importado en cero lugares).

.gitignore — Agregados lambdatest_credentials.txt y config_global.json para que tus credenciales de LambdaTest y contraseñas SMTP nunca sean versionadas por accidente.

2. Drivers Selenium — solo locales
core/browser_manager.py — Se eliminó completamente la lógica de webdriver-manager. Ahora si el driver no está en /drivers/, el mensaje de error dice exactamente qué hacer ("descargá el driver manualmente y colocálo en /drivers/").

3. Chrome, Edge y Firefox no roban el foco
core/browser_manager.py — Chrome y Edge ahora arrancan con --window-position=10000,0 como argumento de inicio, lo que evita que el browser aparezca en pantalla antes de ser movido. Firefox usa set_window_position(10000, 0) post-creación (Firefox no soporta el flag de posición en CLI).

Antes: --start-maximized hacía que el browser abriera en primer plano, interrumpiendo lo que estabas haciendo.
Ahora: El browser arranca directamente fuera de pantalla.

4. Sistema de email restaurado con feedback visual
interface/helpers_interface.py — Restaurado el callback de estado (registrar_callback_ui_email) que existía en el original pero se había perdido.

interface/main_interface.py — Agregado un label de estado que muestra:

⏳ Enviando email... (mientras procesa)
✅ Email enviado correctamente (éxito)
❌ Error al enviar: [motivo] (fallo)
5. Fallback de datos Brasil mejorado
utils/data_generator.py — El fallback cuando falla la API de 4devs ahora genera correctamente el tipo pedido (CPF → algoritmo CPF, CNPJ → algoritmo CNPJ, CEP → lista de CEPs reales). Antes siempre devolvía un CPF sin importar qué tipo se pedía.

6. LambdaTest Mac y Android en Programación
Qué es: Ahora podés programar tests automáticos no solo en Chrome/Firefox/Edge sino también en LambdaTest Mac (Safari en Mac) y LambdaTest Android (Chrome en Samsung Galaxy).

interface/main_interface.py

Dos nuevos checkboxes en la fila de navegadores: LambdaTest Mac y LambdaTest Android (en celeste para distinguirlos visualmente).
programar_test() incluye "lambdatest_mac" / "lambdatest_android" en la lista de navegadores guardada.
monitor_programacion() detecta estos valores y llama directamente a los controllers de LambdaTest en lugar de lanzar un browser local.
run.py — Nuevo flag --run-lambdatest mac|android --pais X --build-name "..." para lanzar LambdaTest desde línea de comandos o como subprocess.

autonomous_runner.py — El ejecutor autónomo (background) detecta LambdaTest en el JSON de programación, salta el loop de viewports (LambdaTest no lo necesita) y lanza el subprocess correspondiente.

7. Email con título de sesión "Automatizado" en LambdaTest
Cuando un test se lanza desde la programación automática, el build en LambdaTest aparece como "Automatizado — [País]", diferenciándolo de los runs manuales que ves en el dashboard de LambdaTest.

8. Email mejorado: PASS/FAILED, browser info y todos los errores
interface/helpers_interface.py — Reescrita la función de email consolidado:

Asunto antes: Resultados Osocio Programación 26/06/2026
Asunto ahora: [PASS] Osocio Programación 26/06/2026 — 15 OK / 0 errores o [FAILED] ... — 12 OK / 3 errores

Cuerpo antes: Bloques sin distinguir el browser, mostraba solo los primeros 10 errores.
Cuerpo ahora:

Resumen global al inicio: 🟢 RESULTADO GLOBAL: PASS o 🔴 RESULTADO GLOBAL: FAILED
Cada bloque muestra el browser legible: Argentina — Chrome / Desktop o Brasil — LambdaTest Mac / Mac (Safari)
🔴 si hubo errores (= lead no enviado), 🟢 si todo OK
Todos los errores listados con número de línea del Excel y URL


Email fix: enviar_email_resultados (runs normales vía botón) ahora tiene el mismo formato que la versión consolidada — asunto con [PASS] o [FAILED], muestra el browser y viewport, todos los errores con 🔴 y número de línea, 🟢 si todo OK.

Browser visibility toggle: Nueva fila "Modo:" en Configuración Global con checkbox "Ver navegador mientras corre". Por defecto desmarcado = browser corre fuera de pantalla sin molestar. Al marcarlo, el browser abre normalmente en pantalla. La programación automática siempre corre en background sin importar el checkbox. El cambio recorre toda la cadena: UI → _runner_common.py → GenericCountryBase → BaseFormFiller → BrowserManager.


Resumen de cambios completados — Ocioso_Update
Contexto
Dos cambios pendientes:

Email format bug: enviar_email_resultados() (usada en runs normales via botón "Enviar Leads") nunca fue actualizada — sigue con el formato viejo sin PASS/FAILED en el asunto, sin indicar el browser, y truncando errores a 10. Solo se actualizó la versión consolidada (para programación automática).

Browser visibility toggle: Por defecto el browser corre fuera de pantalla (--window-position=10000,0) para no interrumpir al usuario. Se pide hacer esto opcional: checkbox en la UI para elegir entre "segundo plano silencioso" (default) o "visible para poder seguirlo".

Cambio 1: Fix email enviar_email_resultados
interface/helpers_interface.py (línea 994)
Actualizar la firma:

def enviar_email_resultados(pais, excel_path, screenshots_dir, browser=None, viewport=None):
Cambios dentro:

Asunto: [PASS] Osocio {pais} {fecha} — {ok} OK / {err} errores (o [FAILED])
Usar _label_navegador(browser, viewport) (ya existe en línea 1103) para mostrar nombre legible en el cuerpo
Mostrar todos los errores (quitar [:10] y el mensaje de truncación)
Ícono 🔴 por cada error detallado, 🟢 si todo OK (mismos íconos que la versión consolidada)
forms/_runner_common.py (línea ~45)
Donde se llama a enviar_email_resultados, pasar browser y viewport que ya están en scope:

enviar_email_resultados(country_name, resultados_path, screenshot_dir, browser=browser, viewport=viewport)
Cambio 2: Browser visibility toggle
Estrategia
Nuevo parámetro background: bool = True en toda la cadena. Cuando True (default), el browser arranca off-screen. Cuando False, abre normalmente. La cadena es:

UI → _ejecutar_envio() → run_func(background=...) → _runner_common.py → GenericCountryBase.__init__() → BaseFormFiller.initialize_browser() → BrowserManager.create_browser(..., background=True)

A. interface/main_interface.py
Nueva BooleanVar — agregar en los 3 bloques donde se inicializan chrome_var etc:

visible_browser_var = BooleanVar(value=False)
También al global statement (línea 2354).

Nuevo checkbox + label informativo — debajo de la fila de viewports en frame_izquierda:

frame_vis = Frame(frame_izquierda, bg=APP_BG_COLOR)
frame_vis.pack(anchor="w", pady=2)
Label(frame_vis, text="Browser:", ..., width=10).pack(side=LEFT)
Checkbutton(frame_vis, text="Ver navegador mientras corre", variable=visible_browser_var, ...).pack(side=LEFT, padx=5)
Label(frame_vis, text="ℹ️ Por defecto corre en segundo plano sin interrumpirte", fg="#888", ...).pack(side=LEFT, padx=4)
_ejecutar_envio() (línea ~3910): pasar background:

run_func(browser=nav, viewport=vp, headless=False, background=not visible_browser_var.get())
monitor_programacion() (~línea 4602): programación siempre silenciosa, pasar background=True explícito.

B. forms/_runner_common.py
Agregar background=True a run_country_form() y al closure _runner() / _DynamicCountry.__init__():

def run_country_form(form_class, country_name, browser, viewport, headless, enviar_email, background=True):
    formulario = form_class(browser=browser, viewport=viewport, headless=headless, background=background)
    ...

def _runner(browser="chrome", viewport="fullscreen", headless=False, enviar_email=True, background=True):
    class _DynamicCountry(GenericCountryBase):
        def __init__(self, ...):
            super().__init__(country_name, browser=browser, viewport=viewport, headless=headless, background=background)
    return run_country_form(..., background=background)
C. core/generic_country_base.py
def __init__(self, country_name, browser="chrome", viewport="fullscreen", headless=False, background=True):
    ...
    config['background'] = background
    super().__init__(config)
D. core/base_form_filler.py (línea 573)
self.driver = BrowserManager.create_browser(
    browser_type=self.config['browser'],
    viewport=self.config['viewport'],
    headless=self.config.get('headless', False),
    background=self.config.get('background', True),
)
E. core/browser_manager.py
Agregar background=True a create_browser() y a los 3 métodos _create_chrome/firefox/edge().

Solo aplicar la posición off-screen cuando background=True:

# Chrome y Edge: solo si background
if background:
    options.add_argument("--window-position=10000,0")

# Firefox: solo si background
if not headless:
    if background:
        driver.set_window_position(10000, 0)
Archivos modificados
Archivo	Cambio
interface/helpers_interface.py	Fix enviar_email_resultados: PASS/FAILED, browser legible, todos los errores
forms/_runner_common.py	Pasar browser/viewport al email; agregar background param en toda la cadena
interface/main_interface.py	visible_browser_var, checkbox, label info, pasar background en ejecuciones
core/generic_country_base.py	Agregar background al __init__ y config dict
core/base_form_filler.py	Pasar background desde config a create_browser()
core/browser_manager.py	Agregar background param, condicionar --window-position con él
Verificación
Email runs normales: Correr un test con el botón → email llega con [PASS] o [FAILED] en asunto, muestra "Chrome / Desktop", lista todos los errores con 🔴.
Browser background (default): Correr con checkbox desmarcado → browser no aparece en pantalla.
Browser visible: Marcar checkbox "Ver navegador mientras corre" → browser abre normalmente en pantalla del usuario.
Programación automática: Tests programados siempre corren en background sin importar el checkbox.


================================================================================
BLOQUE 3 — Programación Semanal Recurrente (Weekly Scheduler)
Fusión y reemplazo completo del sistema de programación automática
================================================================================

CONTEXTO: La fusión entre dos versiones del proyecto
El proyecto original (Form_Automation_Project-pasaje-git-main) tenía un programador de test simple: el usuario elegía una fecha y hora específica, y la app enviaba el formulario una sola vez en ese momento. Era un disparo único.

Tu versión mejorada (Ocioso_Update) fue creciendo con features nuevas, pero el programador seguía siendo ese sistema de fecha única. Se decidió reemplazarlo completamente por un calendario semanal recurrente — el test se programa para ciertos días y horarios y se repite semana a semana automáticamente, sin que el usuario tenga que volver a configurarlo.

──────────────────────────────────────────────────────────────────
A. ANTES: Programador de fecha única (eliminado)
──────────────────────────────────────────────────────────────────

Cómo funcionaba antes:
- El usuario elegía una fecha (ej: 26/06/2026) y una hora (ej: 10:00)
- La app guardaba eso en json/programacion_test.json con formato {"fecha_hora": "2026-06-26 10:00:00", ...}
- Un hilo de fondo revisaba cada minuto si había llegado ese momento
- Al disparar, enviaba el formulario y BORRABA el JSON — el ciclo terminaba
- Para programar de nuevo, el usuario tenía que volver a elegir fecha y hora

Problemas:
- Solo corría una vez. Había que reprogramar manualmente cada semana.
- No se podía elegir días específicos de la semana ni múltiples horarios.
- La UI era básica: dos campos de texto (fecha / hora) y un botón.

──────────────────────────────────────────────────────────────────
B. AHORA: Calendario Semanal Recurrente (WeeklySchedulerPanel + WeeklySchedulerDialog)
──────────────────────────────────────────────────────────────────

Cómo funciona ahora:
- El usuario abre el modal de configuración, selecciona días de la semana (Lun a Dom) y para cada día elige uno o varios horarios en cuartos de hora (09:00, 09:15, 09:30...)
- También elige los países a testear
- Al activar, se guarda en json/programacion_test.json con el nuevo formato:
  {"tipo": "semanal", "horarios": {"Lunes": ["09:00", "09:15"], "Miércoles": ["14:00"]}, "paises": ["Argentina", "Chile"], ...}
- El JSON NUNCA se borra después de correr — persiste para la próxima semana
- Tanto la UI (WeeklySchedulerPanel) como el ejecutor autónomo (autonomous_runner.py) revisan ese JSON cada 60 segundos y disparan cuando el día y la franja horaria coinciden con el reloj de la máquina
- Un diccionario last_triggered {(dia, hora): fecha} evita que se dispare dos veces en el mismo slot del mismo día

Archivos nuevos/modificados:
- interface/weekly_scheduler.py (NUEVO, ~900 líneas): contiene WeeklySchedulerPanel (panel de estado embebido) y WeeklySchedulerDialog (modal de configuración)
- interface/main_interface.py: reemplazó el viejo bloque del programador (~500 líneas) por el nuevo panel
- utils/scheduling.py: actualizado para leer/escribir tanto el esquema semanal como el legado (compatibilidad)
- autonomous_runner.py: el loop principal ahora detecta tipo="semanal" y usa el horario semanal en vez de fecha única; nunca llama limpiar_programacion()

──────────────────────────────────────────────────────────────────
C. FEATURES DEL NUEVO MODAL DE CONFIGURACIÓN
──────────────────────────────────────────────────────────────────

Días de la semana:
- 7 botones (Lun-Dom), cada uno muestra cuántos slots tiene configurados
- Click en un día abre el panel de horas correspondiente

Panel de horas (novedad: cuartos de hora):
- Antes: solo horas exactas (00:00, 01:00, ... 23:00) = 24 slots
- Ahora: cuartos de hora (00:00, 00:15, 00:30, 00:45, 01:00...) = 96 slots por día
- El monitor de disparo calcula el slot actual con slot_min = (minuto // 15) * 15 para tolerar desvíos de hasta 14 minutos

Modo de edición (nuevo):
- "Solo este día": los cambios aplican solo al día seleccionado (comportamiento por defecto)
- "Todos los días": cada click en un slot agrega/quita ese horario en TODOS los días de la semana simultáneamente

Copiar horarios a otros días:
- Botón "Aplicar a otros días": abre selector para elegir a cuáles copiar
- Botón "⚡ TODOS": copia instantáneamente los horarios del día activo a todos los demás días

Países:
- 9 checkboxes (Argentina, Bolivia, Brasil, Chile, Colombia, Ecuador, Paraguay, Peru, Uruguay)
- Toggle "Seleccionar todos / Desmarcar todos"

Footer fijo (fix de UX):
- El botón "Guardar configuración" estaba dentro del scroll y requería bajar para verlo
- Ahora está fijo en la parte inferior del modal, siempre visible sin importar cuánto contenido haya arriba

──────────────────────────────────────────────────────────────────
D. BUGS CORREGIDOS EN EL SCHEDULER
──────────────────────────────────────────────────────────────────

1. Color inválido en Tkinter (#F59E0B22):
   - Antes: bg="#F59E0B22" (color RGBA de 8 dígitos, válido en CSS pero NO en Tkinter) → crash al abrir el modal
   - Ahora: bg=SCH_BG (color sólido de 6 dígitos)

2. AttributeError en _val_lbl:
   - Antes: el trace de BooleanVar disparaba _on_country_toggle durante la construcción del modal, antes de que _val_lbl existiera → crash
   - Ahora: self._val_lbl = None inicializado en __init__, _update_footer() tiene guard "if self._val_lbl is None: return"

3. LambdaTest no detectaba resultado en test programado:
   - Antes: lt_controller.run() era llamado pero el valor de retorno (summary) se descartaba. Se buscaban archivos xlsx en el directorio de resultados, que podía no coincidir con dónde guarda el controller.
   - Ahora: se captura summary = lt_controller.run(...), se usa summary["results_excel"] directamente. Si no hay path en el summary, se hace el fallback de escaneo de directorio. Si hay error, se loguea con el mensaje del controller.

4. Email se enviaba aunque el checkbox estuviera desmarcado:
   - Antes: send_email_cb=_send_consolidated se llamaba siempre al completar el test programado, sin consultar el checkbox "Enviar mail"
   - Ahora: send_email_cb=lambda r: _send_consolidated(r) if enviar_mail_var.get() else None

5. "Enviar mail" aparecía chequeado al arrancar la app:
   - Causa: field_validation_ui.py lee "enviar_mail" del archivo config_global.json y lo setea en el var de la UI. Si el archivo tenía "enviar_mail": true de una sesión anterior, pisaba el valor inicial False del checkbox.
   - Ahora: al inicio de la app se fuerza cfg_global["enviar_mail"] = False (sin condición) y se guarda al archivo. Además, cada cambio del checkbox guarda su valor actual al archivo, manteniendo siempre archivo y UI sincronizados.

6. Radio buttons de modo email no se deshabilitaban:
   - Antes: cuando "Enviar mail" estaba desmarcado, se deshabilitaban entry_email, chk_adj_resultados y chk_adj_screens, pero los radio buttons "1 por país" / "Consolidado al final" quedaban activos
   - Ahora: rb_por_pais y rb_consolidado también se deshabilitan/habilitan junto con los demás controles

7. Botón "Detener" no detenía la ejecución real:
   - Antes: _stop() solo cambiaba el estado visual del panel pero el hilo de ejecución seguía corriendo país por país hasta completar todos
   - Ahora: se usa un threading.Event (self._stop_event). Al hacer click en "Detener", se llama stop_event.set(). La función _ejecutar_programacion() chequea _stopped() entre cada país y cada navegador. LambdaTest recibe el mismo stop_event para que pueda cancelar internamente. El monitor loop también limpia el evento antes de cada disparo automático.

8. Horario no se detectaba (timing):
   - Antes: el monitor comparaba hora = ahora.strftime("%H:00") — solo comparaba contra slots de hora exacta (:00). Si el monitor dormía 60 segundos y despertaba en 10:01 en vez de 10:00, perdía el slot del todo.
   - Ahora: slot_min = (ahora.minute // 15) * 15 → cualquier momento entre 10:00 y 10:14 mapea al slot "10:00", entre 10:15 y 10:29 mapea a "10:15", etc. El diccionario last_triggered sigue evitando doble disparo. Esto aplica tanto al monitor de la UI (WeeklySchedulerPanel) como al ejecutor autónomo (autonomous_runner.py).

================================================================================
BLOQUE 4 — Ajustes de UX en el Panel de Horas del Weekly Scheduler
================================================================================

Los siguientes cambios son pulidos de usabilidad sobre el modal de configuración semanal,
hechos después de pruebas reales con la interfaz.

──────────────────────────────────────────────────────────────────
A. TAMAÑO DE LOS BOTONES DE HORARIO
──────────────────────────────────────────────────────────────────

Iteración 1 — demasiado pequeños:
  - Grid de 8 columnas, font 7, padding mínimo. Los botones eran ilegibles en pantallas normales.

Iteración 2 — demasiado grandes:
  - Se cambió a 4 columnas + font 9 bold. El grid ocupaba demasiado espacio vertical y se necesitaba scrollear mucho.

Iteración 3 — tamaño final (actual):
  - 6 columnas, font 9 (sin bold), padx=4, pady=6, width=5.
  - Equilibrio entre legibilidad y espacio: 96 slots caben en 16 filas cómodas.

──────────────────────────────────────────────────────────────────
B. SCROLL CON RUEDA DEL MOUSE EN EL PANEL DE HORAS
──────────────────────────────────────────────────────────────────

Antes: el canvas del modal tenía canvas.bind_all("<MouseWheel>") para capturar el scroll de la rueda.
El problema es que en Windows, cuando el cursor está sobre un widget Button, ese widget captura el evento
MouseWheel antes de que llegue al canvas, y el scroll no funcionaba — solo se podía scrollear
arrastrando la barra lateral con el mouse.

Ahora: se almacena la referencia al canvas en self._scroll_canvas al construir la UI.
Dentro de _build_hours_panel(), cada botón de hora recibe un bind individual:
  btn.bind("<MouseWheel>", lambda e: self._scroll_canvas.yview_scroll(...))
Así el scroll de la rueda funciona sin importar sobre qué botón esté el cursor.

──────────────────────────────────────────────────────────────────
C. BOTÓN "✓ LISTO" MÁS VISIBLE
──────────────────────────────────────────────────────────────────

Antes: bg=SCH_CARD (#7230A0) con fg=SCH_PRIMARY (#C084FC) — se mezclaba con el fondo del panel y era difícil de ver.
Ahora: bg=SCH_GREEN (#10B981) con fg=SCH_WHITE (#FFFFFF) — botón verde bien visible, consistente con el estado "acción completada".

──────────────────────────────────────────────────────────────────
D. MODO "TODOS LOS DÍAS" — REEMPLAZO EN VEZ DE SUMA
──────────────────────────────────────────────────────────────────

Antes: al activar el modo "Todos los días" y clickear un horario, se añadía o quitaba ese horario específico
en todos los días. Si el lunes ya tenía 09:00, 10:00, 11:00 y se activaba "Todos los días" y se agregaba
12:00, los demás días solo recibían el 12:00 además de lo que ya tenían — no quedaban iguales al lunes.

Ahora: cada vez que se modifica un horario en modo "Todos los días", la lista completa del día seleccionado
se copia íntegra a todos los demás días (reemplazando lo que tenían). El resultado es que todos los días
quedan con exactamente los mismos horarios que el día activo.

  Ejemplo: Lunes tiene [09:00, 10:00]. Modo "Todos los días" activo.
  Se clickea 11:00 → Lunes queda [09:00, 10:00, 11:00] → todos los demás días quedan [09:00, 10:00, 11:00].
  Se clickea 09:00 para quitarlo → Lunes queda [10:00, 11:00] → todos los demás días quedan [10:00, 11:00].

──────────────────────────────────────────────────────────────────
E. FIX: DESELECCIONAR UN HORARIO A VECES REQUERÍA DOS CLICKS
──────────────────────────────────────────────────────────────────

Problema:
  Hacer click en un horario ya seleccionado a veces no lo deseleccionaba, o requería un segundo click.
  Era intermitente: unas veces funcionaba al primer click, otras no.

Causa raíz:
  _toggle_hour() llamaba a _build_copy_ui() al final de cada toggle. Esta función destruye y recrea
  los widgets de la sección "Copiar a otros días" dentro del canvas scrollable. Cuando los widgets
  se destruyen y recrean, Tkinter encola una tarea idle para recalcular el scrollregion del canvas
  (tamaño total del contenido). Esta tarea idle puede ejecutarse durante el evento ButtonPress del
  siguiente click, desplazando los botones de posición antes de que llegue el ButtonRelease. En Tkinter,
  el comando de un Button solo se dispara si el cursor está sobre el widget en el momento del ButtonRelease.
  Si el canvas movió el botón entre Press y Release, el click no se registra.

Fix:
  Se eliminó el doble-delete redundante en el cleanup de schedule (que causaba KeyError silencioso).
  Se condicionó el llamado a _build_copy_ui() para que solo se ejecute cuando la visibilidad de
  la sección cambia (es decir, cuando se pasa de 0 horarios a 1, o de N horarios a 0). En los demás
  toggles (agregar/quitar horarios cuando ya había otros), no se reconstruye la sección de copia y
  no hay reflow del canvas. Así el ButtonRelease siempre encuentra el botón en el mismo lugar.