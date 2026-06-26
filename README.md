# Form Automation Project

Automatización de formularios para varios países con Selenium, generación de resultados y UI de administración.

## Requisitos

- Python 3.10+ o compatible
- Activar el entorno virtual provisto con:
  ```powershell
  .\venv\Scripts\activate
  ```
- Instalar dependencias:
  ```powershell
  pip install -r requirements.txt
  ```

## Ejecución

1. Asegúrate de ejecutar desde la carpeta raíz del proyecto.
2. Inicia la interfaz con:
   ```powershell
   python run.py
   ```
3. Desde la interfaz puedes:
   - cargar o crear el Excel de datos
   - editar filas y guardar cambios
   - lanzar la ejecución de formularios por país
   - abrir resultados y capturas

## Estructura de carpetas

- `core/`: lógica base de automatización y formularios por país.
- `forms/`: scripts de entrada para ejecutar los formularios de cada país.
- `interface/`: UI de administración y utilidades de soporte.
- `data/`: archivos Excel de entrada.
- `drivers/`: drivers locales del navegador (`chromedriver.exe`, `geckodriver.exe`, `msedgedriver.exe`).
- `resultados/`: resultados generados y capturas de pantalla.
- `json/`: configuración persistente y programación.

## Drivers locales (sin descargas automáticas)

La app usa exclusivamente drivers locales desde la carpeta `drivers/` ubicada junto al proyecto (o junto al `.exe` al compilar con PyInstaller).

Estructura esperada:

```text
Form_Automation_Project/
|-- run.py
|-- drivers/
|   |-- chromedriver.exe
|   |-- geckodriver.exe
|   `-- msedgedriver.exe
```

Si el navegador se actualiza y el driver queda desfasado, reemplaza el `.exe` correspondiente dentro de `drivers/`.

## Programación automática

El ejecutor automático usa `json/programacion_test.json` para planificar ejecuciones.
La interfaz mantiene su propio monitor mientras la ventana está abierta.
Si necesitas que la programación siga activa con la interfaz cerrada, inicia el ejecutor autónomo manualmente.

Si necesitas iniciarlo manualmente para pruebas o diagnóstico, puedes usar:

```powershell
python run.py --autonomous
```

También se mantiene compatibilidad con:

```powershell
python json\ejecutor_autonomo.py
```

## Build portable

El script `build_exe.bat` genera:

- `dist/OsocioFormAutomation.exe`
- `dist/OsocioFormAutomation_portable/`

Dentro de `dist/OsocioFormAutomation_portable/` se crean:

- `OsocioFormAutomation.exe`
- `Abrir_Osocio_Form_Automation.bat`
- carpetas externas `data/`, `drivers/`, `json/`, `resultados/` y `temporales/`

Los drivers deben seguir distribuyéndose manualmente dentro de `drivers/`.

## Notas importantes

- El proyecto ya usa `sys.executable` para ejecutar los scripts desde el entorno activo.
- Se normalizó la carpeta `json` para evitar problemas de mayúsculas/minúsculas.
- `requirements.txt` contiene las dependencias necesarias.

---

## Envío de email

El envío se configura en la sección **Configuración Global** (parte superior de la app).

**Destinatario(s)**
- Campo "Email": ingresá una o varias direcciones separadas por coma.
  ```
  usuario@empresa.com, otro@empresa.com
  ```
- El valor se guarda automáticamente en `json/config_global.json`.

**Activar el envío**
- Checkbox **"Enviar mail"** — debe estar marcado antes de ejecutar. Arranca siempre desmarcado por seguridad.
- **"Adjuntar resultados"**: incluye el Excel de resultados como adjunto.
- **"Adjuntar screenshots"**: incluye las capturas de pantalla como adjunto.

**Modo de envío**
- **1 por país**: envía un email independiente al terminar cada país.
- **Consolidado**: espera a que terminen todos los países y manda un único email con el resumen completo.

**Método de envío (automático)**
- Si Outlook Desktop está instalado y tiene una cuenta activa → usa Outlook COM (envía desde tu cuenta Outlook configurada).
- Si no → usa SMTP (requiere credenciales configuradas en `json/config_global.json`).

**Feedback visual**
Junto al campo de email aparece el estado en tiempo real:
- ⏳ Enviando email...
- ✅ Email enviado correctamente
- ❌ Error al enviar: [motivo]

El asunto incluye resultado global, fecha y países ejecutados:
```
[PASS] Osocio — 26/06/2026 — AR CO BO ✓
[FAILED] Osocio — 26/06/2026 — AR ✓ | CL ✗
```

---

## Test programado

El panel **"Test Automático"** (pestaña Testing) permite configurar ejecuciones recurrentes
sin intervención manual. El schedule persiste entre sesiones y se repite cada semana.

> **La app debe estar abierta** para que el monitor detecte el horario y dispare la ejecución.

### Orden de configuración

**Paso 1 — Navegador** (Configuración Global, obligatorio)

Seleccioná al menos uno: **Chrome**, **Firefox** o **Edge** en modo **Desktop**.
El test programado siempre corre en background (ventana fuera de pantalla), independientemente
del checkbox "Ver navegador mientras corre". LambdaTest no es compatible con el modo programado.

**Paso 2 — Email** (Configuración Global, opcional)

Si querés recibir el resultado por email: ingresá el destinatario y marcá el checkbox "Enviar mail".
El email del test programado es siempre consolidado (un solo email al terminar todos los países).

**Paso 3 — Configurar el scheduler**

1. Click en **"⚙ Configurar automatización"** → se abre el modal.
2. Hacer click en un día (Lun–Dom) para abrirlo y seleccionar horarios:
   - **Botones de cuartos de hora**: click en `09:00`, `09:15`, `09:30`, etc. para activar/desactivar.
   - **Horario personalizado**: escribir cualquier hora (ej: `09:33`) en el campo "Personalizado" y presionar **Enter** o **"+ Agregar"**.
   - Los horarios activos aparecen como badges `✕ HH:MM` — hacer click en uno para quitarlo.
   - Click en **"✓ Listo"** para cerrar el panel de horas.
3. **Modo de edición**:
   - *Solo este día*: los cambios aplican únicamente al día seleccionado.
   - *Todos los días*: replica íntegramente el schedule del día activo a todos los demás.
4. **"Copiar a otros días"**: copia los horarios del día actual a días específicos que elijas.
5. En la sección **"🌎 PAÍSES A TESTEAR"**: marcar los países a ejecutar.
6. Click en **"💾 Guardar configuración"** → el badge del panel pasa a "⚙ Configurado".
7. Click en **"▶ Programar test automático"** → badge pasa a "📅 Programado".

Para detener: click en **"■ Desactivar"**.

### Cómo funciona el monitor

- Un hilo en background verifica cada **60 segundos** si hay un horario que coincida con la hora actual.
- El test se dispara si el minuto actual cae en la ventana `[hora configurada, hora + 15 min)`.
  Ejemplo: horario `09:33` → se dispara entre las `09:33` y las `09:47`.
- Cada slot se ejecuta **como máximo una vez por día** (se persiste en `json/scheduler_triggered.json`).
- Si la app estaba cerrada cuando llegó el horario, al reabrirla ejecuta el slot pendiente del mismo día.

### Resetear el scheduler

- Borrar `json/programacion_test.json` → el panel vuelve al estado "Sin configurar".
- El portable siempre arranca sin schedule (ese archivo se excluye del build).

---

## UPDATE — Historial de cambios (Ocioso_Update)

### Bloque 1 — Restauración, seguridad y correcciones base

**1. Seguridad — Credenciales protegidas de git**
`lambdatest_credentials.txt` y `config_global.json` ahora están en `.gitignore`. Nunca se subirán al repositorio por accidente.

**2. Drivers Selenium — solo locales**
Eliminado `webdriver-manager`. La app usa exclusivamente los drivers de la carpeta `/drivers/`. Si el driver está desactualizado, el error indica exactamente qué reemplazar.

**3. Dependencias limpias**
Eliminados `webdriver-manager` y `cpf-and-cnpj-generator` (no se importaba en ningún archivo) de `requirements.txt`.

**4. Browsers no roban el foco**
Chrome y Edge arrancan con `--window-position=10000,0` como argumento de inicio → abren directamente fuera de pantalla sin interrumpir al usuario. Antes usaban `--start-maximized` que abría en primer plano.

**5. Feedback visual de email restaurado**
El label de estado al lado del campo de email muestra en tiempo real:
- ⏳ Enviando email...
- ✅ Email enviado correctamente
- ❌ Error al enviar: [motivo]

**6. Fallback de datos Brasil corregido**
`utils/data_generator.py`: cuando falla la API de 4devs, el fallback ahora genera el tipo correcto (CPF → algoritmo CPF, CNPJ → algoritmo CNPJ, CEP → lista de CEPs reales). Antes siempre devolvía un CPF.

---

### Bloque 2 — Email mejorado y visibilidad del browser

**7. Email con PASS/FAILED en asunto**
Asunto: `[PASS] Osocio Programación 26/06/2026 — 15 OK / 0 errores` o `[FAILED] ... — 12 OK / 3 errores`. Aplica tanto al email de runs normales (botón) como al consolidado de programación automática.

**10. Checkbox "Ver navegador mientras corre"**
Por defecto el browser corre fuera de pantalla (silencioso). Al marcar el checkbox, abre normalmente en pantalla. La programación automática siempre corre en background sin importar este checkbox.

---

### Bloque 3 — Programación semanal recurrente (Weekly Scheduler)

Reemplazo completo del programador de fecha única por un calendario semanal recurrente.

**Antes:** el usuario elegía una fecha y hora específica → la app disparaba una vez y borraba el JSON. Había que reprogramar cada semana manualmente.

**Ahora:** calendario semanal persistente con días y horarios en cuartos de hora (96 slots/día). El JSON nunca se borra. El test se repite automáticamente cada semana.

**Features del modal de configuración:**
- 7 botones de día (Lun–Dom), cada uno muestra cuántos slots tiene
- 96 slots por día en cuartos de hora: 00:00, 00:15, 00:30, 00:45...
- Modo "Solo este día" / "Todos los días" para editar uno o todos a la vez
- En modo "Todos los días": reemplaza íntegramente la config de todos los días con la del día activo
- Botón "Aplicar a otros días" y "⚡ TODOS" para copiar horarios
- Footer fijo siempre visible (botón Guardar no queda dentro del scroll)

**Bugs corregidos en el scheduler:**
- Color RGBA inválido en Tkinter (`#F59E0B22`) → crash al abrir el modal
- `AttributeError` en `_val_lbl` por trace de `BooleanVar` antes de que el widget existiera
- LambdaTest no detectaba resultado: ahora captura `summary = lt_controller.run(...)` directamente
- Email se enviaba aunque el checkbox estuviera desmarcado
- "Enviar mail" aparecía chequeado al arrancar (causa: JSON con valor stale de sesión anterior)
- Radio buttons de modo email no se deshabilitaban junto con el resto de controles
- Botón "Detener" solo cambiaba el estado visual; ahora usa `threading.Event` para cancelar realmente
- Horario no se detectaba si el monitor despertaba 1 minuto tarde → ahora usa `slot_min = (minuto // 15) * 15` para tolerar desvíos de hasta 14 minutos

---

### Bloque 4 — Ajustes de UX en el panel de horas

- **Tamaño de botones de horario:** ajustado a 6 columnas, font 9, padding cómodo (96 slots en 16 filas)
- **Scroll con rueda del mouse:** en Windows los botones interceptaban el evento MouseWheel. Fix: cada botón recibe su propio bind que redirige al canvas
- **Botón "✓ Listo":** ahora en verde (`#10B981`) con texto blanco, más visible
- **Fix doble click para deseleccionar:** `_build_copy_ui()` reconstruía widgets en cada toggle causando reflow del canvas entre `ButtonPress` y `ButtonRelease`. Fix: solo reconstruye cuando la sección cambia de visible a oculta (0↔N horarios)

---

### Bloque 5 — Email, múltiples destinatarios y fix re-ejecución

**11. Email HTML con emojis visibles en Outlook clásico**
El body del email se enviaba como texto plano — Outlook clásico no mostraba ✅ ❌ 🟢 🔴. Ahora se envía como HTML con `font-family: 'Segoe UI Emoji'`. Aplica tanto al path de Outlook COM (`mail.HTMLBody`) como al SMTP (`MIMEText html`).

**12. Múltiples destinatarios de email**
El campo de email ahora acepta varios emails separados por coma: `email1@dominio.com, email2@dominio.com`. El hint `(varios emails separados por coma)` aparece junto al campo. `obtener_email_destinatario()` parsea la lista y la pasa al enviador (ambos paths ya soportaban lista).

**13. Fix: re-ejecución al reabrir la app**
Al cerrar y reabrir la app dentro del mismo slot de 15 minutos, el monitor volvía a disparar el test porque `last_triggered` era en memoria y se reseteaba. Fix en dos capas:
- Se persiste en `json/scheduler_triggered.json` para sobrevivir reinicios.
- En la primera iteración del monitor (startup), cualquier slot que coincida con la hora actual se registra como "ya visto" sin ejecutar. Esto evita el disparo falso si se abre la app justo en un horario programado. Los slots nuevos que llegan mientras la app está corriendo se detectan y ejecutan normalmente.

**14. Build: spec actualizado**
`FormAutomation.spec` ahora incluye `interface.weekly_scheduler` y todos los submodules de `validation/` en `hiddenimports`, necesarios para que el build portable funcione correctamente.
