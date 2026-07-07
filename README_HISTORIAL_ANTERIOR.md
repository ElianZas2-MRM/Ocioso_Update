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
- `resultados/`: resultados generados y capturas de pantalla (Envío de Leads / LambdaTest).
- `Dealerscheck_resultados/`: reportes Excel y capturas del Comparador Dealers.
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

El script `build.bat` genera:

- `dist/OsocioFormAutomation.exe`
- `dist/OsocioFormAutomation_portable/`
- `dist/OsocioFormAutomation_portable.zip`

Dentro de `dist/OsocioFormAutomation_portable/` se crean:

- `OsocioFormAutomation.exe`
- `Abrir_Osocio_Form_Automation.bat`
- carpetas externas `data/`, `drivers/`, `json/`, `resultados/`, `temporales/`, `Dealerscheck_resultados/`, `lambdatest_mac/` y `lambdatest_android/`

El portable arranca siempre sin schedule activo ni configuración personal del Comparador Dealers (esos JSON se excluyen del build). Los drivers deben seguir distribuyéndose manualmente dentro de `drivers/`.

## Migración a formularios visid

Los nuevos formularios estandarizados (contenedor `.visid-form-content` / `.visid-fields-grid`) son soportados automáticamente — no se necesita cambiar el Excel ni la configuración por país cuando un mercado migra.

**Cómo funciona la coexistencia:**
El engine tiene un mapa de aliases: si el ID del campo configurado no aparece en el DOM del nuevo form, automáticamente prueba el ID equivalente del estándar visid. Funciona en Desktop, LambdaTest Mac y Android.

| Campo | ID actual (forms viejos) | ID visid |
|-------|--------------------------|----------|
| Nombre | `firstname`, variantes AEM | `name` |
| Modelo | `models`, `model_1` | `model` |
| Fecha estimada | `estimated-date-purchase` | `estimated-day` |
| Resto de campos | `lastname`, `document`, `phone`, `email`, `city`, `dealer` | iguales ✓ |

**Submit button:** `.btn-visid-submit` — ya soportado.
**Form de 3 pasos:** `.btn-steps-submit` — ya soportado.
**Checkbox términos:** `#terms-and-conditions` — detectado automáticamente.

**Si un campo no se llena en un form visid nuevo:** agregar el mapping viejo → nuevo en `_VISID_ID_ALIASES` dentro de `core/base_form_filler.py` y `lambdatest_mac/lt_runner.py`.

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

**Qué se reporta en el cuerpo del email**

El email diferencia tres tipos de situaciones:

**1. Todo OK**
```
✅ OK — todos los formularios insertados y leads enviados correctamente

=== RESUMEN ===
  Exitosos:    5
  Con errores: 0
  Total filas: 5

✅ Todos los formularios se completaron exitosamente.
```

**2. Formulario no inserto correctamente** — el Excel tiene una URL en columna B (form URL) pero esa URL no se encontró como iframe en la landing page. El test falla aunque el lead haya podido enviarse por un iframe alternativo.
```
❌ FORMULARIO NO INSERTO en 1 fila(s)

=== RESUMEN ===
  Exitosos:    4
  Con errores: 1
  Total filas: 5

=== FORMULARIOS NO INSERTADOS CORRECTAMENTE ===

  ⚠️  Línea 3 | Landing: https://www.marca.com/ar/modelo
      URL Form esperada:   https://forms.hubspot.com/form/abc123
      URL Form encontrada: ninguna
```
> La columna "Form coincide" en el Excel de resultados muestra **NO** en esa fila, con la URL real del iframe encontrado (si hubo alguno).

**3. Error de envío** — el formulario estaba correctamente inserto pero el lead no se pudo enviar (timeout, error del servidor, etc.).
```
❌ LEAD NO ENVIADO en 1 formulario(s)

=== RESUMEN ===
  Exitosos:    4
  Con errores: 1
  Total filas: 5

=== ERRORES DE ENVÍO ===

  ❌ Línea 2 | URL: https://www.marca.com/ar/modelo
     Error: Timeout esperando página de confirmación (div#thank-you no apareció)
```

**4. Ambos tipos de error** — si hay filas con form no inserto Y filas con fallo de envío, el email muestra ambas secciones y el estado indica `ERRORES MÚLTIPLES`.

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

## Guía de uso — desde cero

### Paso previo: preparar el Excel de datos

La app necesita un archivo Excel con las URLs y datos para rellenar los formularios.

1. Abrí la app: `python run.py`
2. En la sección de datos (botón "Abrir/Crear Excel"):
   - Primera vez: hacé click en **"Crear Excel"** → se genera un archivo con las columnas correctas en `data/`.
   - Ya tenés un Excel: hacé click en **"Cargar Excel"** y seleccioná tu archivo.
3. Completá las columnas requeridas:
   - **URL Landing**: la página donde está inserto el formulario.
   - **URL Form**: la URL del iframe del formulario (ej. HubSpot, Marketo, etc.).
   - Datos del lead: nombre, apellido, documento, teléfono, email, ciudad, modelo, dealer, fecha estimada, etc.
4. Hacé click en **"Guardar"** para persistir los cambios.

---

### Pestaña: Desktop

Automatiza el rellenado de formularios directamente desde tu PC usando Chrome, Firefox o Edge.

**Cómo usarla:**

1. **Configuración Global** (parte superior):
   - Seleccioná el navegador: Chrome, Firefox o Edge.
   - Viewport: Desktop (1366×768) o Mobile (600×738).
   - Checkbox **"Ver navegador mientras corre"**: si lo marcás, el browser abre en pantalla. Por defecto corre en background sin robar el foco.
   - Campo **"Email"**: ingresá el/los destinatarios separados por coma.
   - Checkbox **"Enviar mail"**: debe estar marcado si querés recibir el resultado por email. Arranca siempre desmarcado.
   - **"Adjuntar resultados"** / **"Adjuntar screenshots"**: incluyen archivos al email.
   - Modo de envío: **"1 por país"** (email al terminar cada país) o **"Consolidado"** (un solo email al final de todos).

2. **Selección de países**: marcá los países que querés ejecutar.

3. **Botón "Ejecutar"**: inicia la automatización. Aparece un overlay semitransparente con el progreso por país:
   - `En curso → 1/5 → 2/5 → ... → 5/5`
   - Al terminar, si hay email: `📧 Enviando email...` → `✉ Email enviado` / `✉ Email no enviado`

4. **Botón "Detener"**: cancela la ejecución en curso. Si "Enviar mail" estaba activo, no envía email al detener manualmente (no es un fallo real).

5. Al terminar: los resultados quedan en `resultados/<País>/` con el Excel de resultados y capturas de pantalla. Si configuraste email, llega un resumen con tabla de URLs (fails primero, pasados después).

---

### Pestaña: LambdaTest Mac

Ejecuta los formularios en Safari/Chrome/Firefox sobre macOS real via LambdaTest (requiere credenciales en `lambdatest_credentials.txt`).

**Cómo usarla:**

1. Asegurate de tener `lambdatest_credentials.txt` con tu username y access key de LambdaTest.
2. Seleccioná el navegador y la versión de macOS.
3. Seleccioná los países y hacé click en **"Ejecutar"**.
4. El progreso aparece en el overlay igual que en Desktop.
5. Al terminar se envía el email consolidado con los resultados si configuraste email.

> Las credenciales de LambdaTest nunca se suben al repositorio — están en `.gitignore`.

---

### Pestaña: LambdaTest Android

Igual que Mac pero ejecuta en dispositivos Android reales via LambdaTest.

- Los campos se rellenan por JS puro (sin teclado virtual, más rápido y sin animaciones).
- Útil para verificar que el formulario funciona en móvil antes de lanzar una campaña.

---

### Pestaña: Test Automático

Configura ejecuciones recurrentes semanales sin intervención manual.

**Cómo usarla desde cero:**

1. **Configurar email** (Configuración Global): ingresá destinatario y marcá "Enviar mail".
2. **Abrir el modal**: click en **"⚙ Configurar automatización"**.
3. **Seleccionar días y horarios**:
   - Click en un día (Lun–Dom) para expandirlo.
   - Click en los botones de cuarto de hora (ej: `09:00`, `09:15`...) para activar ese slot.
   - O escribí un horario personalizado (ej: `09:33`) y presioná Enter.
4. **Seleccionar países** en la sección "🌎 PAÍSES A TESTEAR".
5. Click en **"💾 Guardar configuración"**.
6. Click en **"▶ Programar test automático"** → el test se dispara automáticamente cada semana a la hora configurada.

> La app debe estar abierta para que el monitor detecte el horario y dispare la ejecución. Al dispararse aparece el mismo overlay con progreso y botón Detener.

---

## Guía completa de la app (todas las pestañas, versión actual)

Esta sección documenta en detalle, y en limpio, cómo funciona cada pestaña de la app tal como está hoy. No incluye datos sensibles (usuarios/claves de LambdaTest, direcciones de email reales, etc.) — esos siempre se cargan de forma local en `lambdatest_credentials.txt` y `json/config_global.json`, ambos excluidos de git.

### Configuración Global (barra superior, compartida por todas las pestañas)

- **Email destinatario**: campo de texto, deshabilitado hasta marcar "Enviar mail". Acepta una o varias direcciones separadas por coma.
- **Enviar mail**: al marcarlo aparecen "Adjuntar resultados" (adjunta el Excel de resultados), "Adjuntar screenshots" (adjunta las capturas) y el modo de envío: **"1 por país"** (un email por país al terminar) o **"Consolidado"** (un solo email al final con el resumen de todos).
- El checkbox **"Ver navegador mientras corre"** y **"Minimizar a la bandeja al cerrar"** ya no están en esta barra superior: viven dentro del modal **"⚙ Configurar" → "Configuración avanzada"** de la pestaña Envío de Leads (ver abajo).

### Pestaña: Envío de Leads

Es la pestaña principal para correr los formularios y enviar leads de verdad.

1. **"⚙ Configurar" → Configuración avanzada**: elegís entre "Un Excel por dispositivo (recomendado)" o "Un Excel compartido para todos los dispositivos", y marcás/desmarcás "Ver navegador mientras corre (solo Chrome / Firefox / Edge)" y "Minimizar a la bandeja al cerrar".
2. **Modo de ejecución** (pills): "Mercados" en Consecutivo o Paralelo, y "Excels por mercado" en Consecutivo o Paralelo.
3. **Dispositivos / Navegadores** (selección múltiple): Chrome, Firefox, Edge, Mac LT, Android LT. Al elegir Mac LT o Android LT aparece el panel de credenciales de LambdaTest (usuario/access key, se auto-cargan desde `lambdatest_credentials.txt` si existe).
4. **"⚡ Enviar en paralelo por URL"** (opcional): abre una sesión de navegador por URL en simultáneo, con un campo para limitar la concurrencia máxima.
5. **"🌎 PAÍSES A EJECUTAR"**: tarjetas de los 9 mercados (AR/BO/BR/CL/CO/EC/PY/PE/UY) — click para marcar/desmarcar, con contador "N seleccionados".
6. **"EJECUTAR ENVÍO"**: se habilita al marcar al menos un país. Abre un modal centrado que bloquea la ventana principal y muestra el progreso en vivo por país (`1/5 → 2/5...`); el botón pasa a "EN CURSO..." mientras corre.
7. **"Resultados"**: abre la carpeta `resultados/` (Excel de resultados + capturas por país/dispositivo).

### Pestaña: Programación de Tests

Programa ejecuciones recurrentes semanales sin intervención manual (ver también la sección "Test programado" más arriba, con el detalle del monitor de 60s).

1. Click en **"⚙ Configurar automatización"** → abre el calendario semanal (días Lun a Dom, 96 franjas de 15 minutos por día).
2. Elegís horarios por día (botones de cuarto de hora o uno personalizado con Enter), modo "Solo este día" / "Todos los días", y podés copiar horarios entre días.
3. Elegís mercados y dispositivos/navegadores (mismos selectores que en Envío de Leads).
4. Al guardar, la app valida que existan los Excel necesarios en `data/` para cada combinación país+dispositivo elegida — si falta alguno, muestra "Archivos Excel Faltantes" (generalos antes desde "Generar Excels con Datos").
5. **"Programar test automático"** activa la programación; una vez activa, el botón se reemplaza por **"Iniciar ahora"** (dispara ya, sin esperar el horario) y **"Desactivar"**.
6. La app debe estar abierta para que el monitor detecte el horario y dispare la ejecución (ver "Cómo funciona el monitor" más arriba).

### Pestaña: Validación de Campos

Sirve para chequear que las reglas de validación (regex, largo, obligatoriedad) de cada campo del formulario real coincidan con lo configurado, sin enviar un lead real.

1. **Tabla de URLs** (País / URL / Formulario): se carga desde un Excel; botones **"Abrir Excel"** y **"Actualizar"** para recargar. Checkbox **"Ver navegador"** para ver el browser mientras corre.
2. **"▶ Configuración de ID"** (expandible): formulario para mapear cada campo — ID del elemento, descripción, si es dropdown, si es numérico, regex completo y por carácter, y reglas rápidas (letras minúsculas, email, campo obligatorio, etc.). Incluye sub-formularios para "Mensaje de error" y "Dependencia" (campos que dependen de otro, ej. ciudad depende de región). La tabla de reglas ya configuradas permite editar o **"Eliminar regla"**.
3. **"Ejecutar validación"**: corre la validación contra el/los form(s) reales de forma asíncrona.
4. **"Resultados"**: abre la carpeta de resultados de validación.

### Pestaña: Generar Excels con Datos

Genera el Excel de datos de prueba (nombre, documento, teléfono, email, modelo, etc.) que después usan "Envío de Leads" y "Programación de Tests".

1. **Modo de URLs**: pill "URL Landing + URL Form" o "Solo URL Form" — define el formato del texto que pegás en el cuadro de URLs (`url landing • url form • ...` o `url form • url form • ...`).
2. **Mercado**: elegís un país a la vez (se genera un Excel por mercado).
3. **Dispositivos**: selección múltiple (Chrome, Firefox, Edge, Mac LT, Android LT) — se genera un Excel por dispositivo elegido, o uno solo si tildás "Excel compartido" en la config avanzada.
4. **"🧩 Es formulario T3 2.0"**: si el form es la versión Adobe AEM nueva, tilda esto — el archivo se genera con sufijo `_T3.xlsx` para diferenciarlo.
5. Botones de la barra fija inferior: **"GENERAR EXCELS"** (crea los archivos), **"REGENERAR DATOS"** (recrea los datos aleatorios manteniendo las URLs ya cargadas), **"Borrar URLs"** (limpia el cuadro de texto).
6. Los archivos quedan en `data/`, con el patrón `Lead_information_Formulario_<País>_<Dispositivo>.xlsx` (o `_T3.xlsx` / `_Generico.xlsx` según el modo).

### Pestaña: Comparador Dealers

Chequea que los concesionarios (dealers) de una marca estén correctamente cargados en un formulario real, comparando contra un Excel de dealers esperados — reemplaza el bookmarklet manual que se pegaba antes en la consola del navegador.

**1 · Mercado a chequear** — tarjetas de los 9 países (AR/BO/BR/CL/CO/EC/PY/PE/UY). Cada país guarda su propia configuración de columnas/URLs automáticamente al cambiar de mercado.

**2 · URL del formulario a chequear** — mismo bloque visual que "Generar Excels": pill "URL Landing + URL Form" / "Solo URL Form", y un cuadro de texto donde podés pegar **una o varias URLs** (se procesan todas en la misma pasada):
- Modo Landing+Form: `url landing` / `url form` / `url landing` / `url form` / ... (de a 2 líneas por form).
- Modo Solo Form: una URL de form por línea.
- Si las URLs pegadas parecen ser de un país distinto al mercado seleccionado, aparece un aviso (mismo criterio que "Generar Excels con Datos").
- Selector de navegador (Chrome/Firefox/Edge) y checkbox **"Ver navegador mientras corre"** (apagado por defecto: corre atrás sin molestar, igual que "Envío de Leads").

**3 · Excel de dealers a chequear** — botón "Seleccionar Excel" (cualquier .xlsx/.xls), fila de encabezado configurable (el Excel real no siempre arranca en la fila 1), y toggle **"Este Excel: Tiene columna de filtro"** / **"No tiene filtro (usar todas las filas)"**. Con filtro, se define la columna, el valor a buscar, y la condición: **Incluir**, **Excluir**, o **Buscar extras** (esta última además dispara la búsqueda de dealers EXTRA/duplicados, ver punto 4).

**4 · Columnas del Excel** — pills **region / city / dealer** (ids reales del HTML del `<select>`, no cambian de país a país aunque el label visible sí — ej. en Argentina se ve "Provincia" pero el id sigue siendo `region`); `dealer` es siempre obligatorio, `region`/`city` se pueden apagar si el form no los tiene. Columnas del Excel a usar para región/ciudad/dealer/BAC (verificar BAC es opcional, muchos forms no lo exponen en el HTML). Checkbox para además **buscar dealers EXTRA** (presentes en el form pero no en el Excel filtrado, o en el Excel completo si "No tiene filtro" está activo) **y DUPLICADOS** (mismo dealer repetido dentro del `<select>` del form — no importa si el Excel tiene duplicados, lo que importa es lo que ve el usuario en el form real). Sección para agregar **columnas adicionales a comprobar en el form**: pares (columna del Excel, id del campo en el HTML) para validar cualquier otro campo del formulario contra el Excel.

**Modelos** (opcional, solo forms T1 — no soporta T2/T3 todavía) — si el form tiene selector de modelo (id `models`), podés correr la comparación para "Todos los modelos" (los detecta en vivo) o para modelos específicos (lista separada por coma). Repite la comparación completa por cada modelo.

**5 · Modo de salida** — "Solo Excel" o "Excel + Capturas (ZIP)". Con capturas, cada imagen lleva un banner con la URL del form arriba (mismo mecanismo que usa el resto de la app), y solo queda el ZIP final (no se dejan sueltas las capturas individuales).

**6 · Configuraciones guardadas** — guardá el mapeo completo de columnas con el nombre que quieras (ej. "GMUY Livianos") para reusarlo después sin volver a configurar todo; podés cargar o eliminar configuraciones guardadas.

**Ejecutar**: valida todo antes de arrancar (Excel seleccionado, columnas resueltas, al menos una URL de form) — si falta algo, muestra un aviso corto sin llegar a abrir nada, así nunca queda la pestaña bloqueada por error. Si todo está bien, abre un modal de progreso (igual estilo que "Envío de Leads", bloquea la ventana mientras corre) con botón Detener. Al terminar, muestra el resumen PASS/FAIL/EXTRA/DUPLICADO y un botón para cerrar.

**Resultados**: el reporte Excel (con colores por estado: verde PASS, rojo FAIL, amarillo EXTRA, naranja DUPLICADO) y el ZIP de capturas quedan en la carpeta `Dealerscheck_resultados/` en la raíz del proyecto.

**Tiempos de espera**: no son configurables — el motor reintenta automáticamente (primer intento rápido, hasta un máximo de 1.2s si el select tarda más en poblarse, por ejemplo en departamentos con muchas ciudades).

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

---

### Bloque 6 — LambdaTest email, Excel de resultados y optimización Android

**15. Email al terminar LambdaTest Mac y Android**
LambdaTest Mac y Android no enviaban email al terminar — solo lo hacía Desktop. Ahora, al completar todos los países seleccionados, se envía automáticamente el email consolidado con los resultados, igual que Desktop en modo consolidado. El email usa los labels "LambdaTest Mac" / "LambdaTest Android" en el asunto y cuerpo.

**16. Columna "Form URL esperada" en Excel de resultados**
Se agregó la columna `Form URL esperada` (antes de `Form URL encontrada`) en el Excel de resultados de Desktop, LambdaTest Mac y LambdaTest Android. Muestra la URL del formulario que debería estar inserto (columna B del Excel de origen), para poder comparar a simple vista qué se esperaba vs qué se encontró sin tener que cruzar con el Excel original.

**17. Fix: "Form coincide" ahora siempre escribe SI/NO**
Antes, si no se encontraba ningún iframe (`Form URL encontrada` vacía), la celda `Form coincide` quedaba en blanco aunque hubiera una URL esperada. Ahora escribe "NO" en esa fila siempre que exista una URL esperada, independientemente de si se encontró algo.

**18. Fix: browser en background nunca roba el foco**
`--window-position=10000,0` movía la ventana fuera de pantalla visualmente, pero Windows igual le daba foco brevemente al abrirla. Ahora, después de crear el driver en modo background, se llama `minimize_window()` → la ventana queda minimizada en la barra de tareas y nunca interrumpe al usuario mientras escribe o trabaja. Aplica a Chrome, Firefox y Edge.

**19. Android — llenado de campos por JS (sin teclado virtual)**
`_fill_text_android` usaba `click()` + `send_keys` carácter a carácter, lo que abría el teclado virtual y escribía lento. Ahora usa JS puro con `native setter` + eventos `touchstart/touchend/input/change/blur`, igual que la versión iOS. Sin teclado, sin esperas de animación, una sola llamada de red por campo.

### Bloque 7 — Overlay semi-transparente, ejecución programada con overlay y detalle de URLs en email

**20. Overlay semi-transparente con fondo visible**
El overlay que bloquea la UI durante la ejecución cambió de un Frame sólido a un `Toplevel` con `alpha=0.88`. El fondo de la ventana se ve levemente a través del overlay, lo que queda visualmente más limpio. El botón "Detener ejecución" sigue siendo completamente opaco y visible. El overlay se reposiciona automáticamente si se mueve o redimensiona la ventana.

**21. Overlay también aparece durante ejecución programada (Test Automático)**
Cuando el scheduler semanal dispara una ejecución, ahora aparece el mismo overlay con progreso por país y el botón "Detener ejecución". Al presionar Detener, se cancela la ejecución programada inmediatamente y el overlay desaparece.

**22. Tabla de URLs en el cuerpo del email (HTML)**
El email de resultados incluye al final una tabla HTML con columnas `URL Landing | URL Secure / Stage` y el ícono ✅/❌ por fila. Si hay error, aparece en rojo debajo de la fila correspondiente. Las filas se ordenan por número de línea del Excel. El campo "Formulario" (Desktop) y "Form URL esperada" (LambdaTest) se mapean automáticamente a la columna Secure/Stage. Aplica a email individual y consolidado.

**23. Consola eliminada de la interfaz**
El panel "Consola" en la parte inferior fue removido. Los logs internos se emiten al stdout nativo (visible en la terminal si se corre con Python, no en el exe). El overlay con progreso por país cubre la necesidad de feedback visual durante la ejecución.

---

### Bloque 8 — Progreso por lead, estado de email en overlay y correcciones de estabilidad

**24. Progreso por lead en el overlay**
El overlay ahora muestra cuántos leads se enviaron en tiempo real: `En curso → 1/5 → 2/5 → 3/5...`. Al completar todos los leads, el contador muestra el total correcto (`5/5`) en lugar de quedarse en `0/5`. El contador se actualiza después de cada fila procesada del Excel, sin importar si el lead fue exitoso o tuvo error.

**25. Estado del email en el overlay**
Al terminar la ejecución de un país, si "Enviar mail" está activo, el overlay muestra el estado del email en tiempo real:
- `📧 Enviando email...` — mientras se envía
- `✉ Email enviado` (verde) — si llegó correctamente
- `✉ Email no enviado` (rojo) — si hubo error al enviar

El estado del email es el estado final del overlay y no es sobreescrito por "Completado".

**26. No se envía email si el usuario detiene la ejecución**
Si se presiona "Detener", el email no se envía aunque "Enviar mail" estuviera activo. Una detención manual no es un fallo real y no merece reportarse.

**27. Rediseño de la tabla de URLs en el email**
- Headers renombrados: `URL LANDING | URL FORM` (antes `URL Landing | URL Secure / Stage`).
- Las filas fallidas aparecen primero, luego las exitosas.
- El cuerpo del email usa un resumen compacto: `FAILED (2): /ar/modelo | /cl/modelo` y `PASSED (4): /co/... | ...` en lugar de secciones verbosas de detalle de errores.

**28. Fix: formularios no se rellenaban (UnicodeEncodeError en stdout)**
Al eliminar el panel de consola, Windows restauró el stdout con encoding `cp1252`. Los `print()` con emojis (ej. `🔹 Procesando select...`) crasheaban con `UnicodeEncodeError` antes de que los campos se rellenaran, lo que dejaba el formulario vacío y disparaba los errores de validación del form. Fix: `sys.stdout.reconfigure(encoding="utf-8", errors="replace")` al inicio de la app.

El diagnóstico se hizo via `debug_run.log` — un archivo de log que el engine escribe en el directorio raíz del proyecto cuando ocurre una excepción durante el rellenado. Si un formulario no se llena y no hay error visible, revisar ese archivo.

---

### Bloque 9 — Radio buttons en Windows 11, sincronización del botón Detener y ajustes del scheduler

**29. Fix: radio buttons aparecían "marcados" sin selección real (Windows 11)**
En Windows 11, el `selectcolor` de `Radiobutton` es ignorado por el tema nativo, haciendo que todo el grupo se vea seleccionado aunque no haya ninguna opción elegida. Se reemplazaron los `Radiobutton` nativos por un indicador custom (○ no seleccionado / ● seleccionado) en: "Modo de envío" (Envío de Leads), selector de País y "Tipo de URLs" (Generar Excels con Datos).

**30. Renombrado "Mobile emulado" → "Mobile Emulado-Navegador"**

**31. Fix: el botón Detener del overlay no sincronizaba la tarjeta de Test Automático**
Detener desde el overlay global y detener desde la tarjeta del scheduler eran dos caminos independientes que no se avisaban entre sí. Ahora ambos comparten el mismo `request_stop()`, y al detener la tarjeta muestra: "Detenido. El navegador se cerrará solo al terminar el lead en curso (si no, cerralo manualmente)."

**32. Tarjeta de Test Automático: acciones tras detener**
Después de detener (o al completar), si sigue habiendo una configuración guardada, la tarjeta vuelve a mostrar "▶ Iniciar ahora" y "■ Desactivar" en vez de ofrecer solo "Desactivar".

**33. Se quitó el badge "Detenido" de la tarjeta**
Podía confundirse con "la programación no se va a ejecutar más". Ahora, mientras exista una configuración guardada, se muestra el badge normal "✓ Configurado".

**34. Fix: copiar horarios a otros días ahora reemplaza en vez de fusionar**
Al desmarcar/borrar horarios de un día y aplicar "para todos los días" o "para días seleccionados", los días destino ahora quedan exactamente iguales al día origen (incluyendo las horas borradas), en lugar de solo sumar horarios nuevos sin quitar los removidos.

---

### Bloque 10 — Editor de celdas, auto-refresh, compatibilidad gm_front y mejoras de overlay

**35. Editor de celdas inline en tabla Excel (Envío de Leads)**
Se reemplazó el comportamiento anterior (que borraba el valor al moverse de celda) por un editor inline basado en un `Entry` flotante:
- Click en cualquier celda abre el editor directamente sobre esa celda.
- El valor se preserva al moverse a otra celda, al scrollear con la rueda del mouse, o al presionar "Guardar Cambios".
- Fixes específicos: posición del `Entry` usaba `bbox(item, col_index)` con entero en vez de `"#N"` string (causaba que apareciera en la posición incorrecta); la celda guardada era la nueva en vez de la anterior (se invierte el orden: `finish_edit()` primero, luego actualizar celda activa).

**36. Auto-refresh de la pestaña "Envío de Leads" al cambiar de pestaña**
Al generar o regenerar un Excel en "Generar Excels con Datos" y luego cambiar a "Envío de Leads", la tabla se recarga automáticamente desde disco sin necesidad de presionar "Actualizar". Implementado mediante `<<NotebookTabChanged>>` en el notebook principal.

**37. Fix definitivo: browser en background no roba el foco**
Se eliminó `minimize_window()` de los tres drivers (Chrome, Firefox, Edge). En Windows, `minimize_window()` hacia que el sistema restaurara el foco brevemente cada vez que Selenium interactuaba con la ventana. El reemplazo es exclusivamente `--window-position=10000,0` (Chrome/Edge) y `driver.set_window_position(10000, 0)` (Firefox), que mantiene la ventana fuera de pantalla sin que Windows le asigne foco en ningún momento.

**38. Modo de envío desbloquea checkboxes de países y botón Enviar**
En la sección "Envío de Leads", los checkboxes de países y el botón "Enviar Leads" aparecen deshabilitados hasta que el usuario selecciona "Múltiples países consecutivos" o "Múltiples países paralelos". Evita errores de configuración incompleta.

**39. Disclaimer al detener ejecución desde el overlay**
Al presionar "Detener ejecución" en el overlay de envío de leads, el botón pasa a "Deteniendo..." (deshabilitado) y aparece el mensaje: _"El lead en curso terminará de enviarse y el navegador se cerrará solo al finalizar."_ Mismo comportamiento ya existente en el Test Automático, ahora también en el overlay de ejecución manual.

**40. Compatibilidad con formularios React/SPA (gm_front y similares)**
El motor de llenado ahora maneja correctamente formularios basados en React u otros frameworks modernos:
- **Aliases extendidos**: `telephone → phone`, `cellphone → phone`, `ci → document` — cubre los IDs que usan los nuevos formularios sin romper los existentes.
- **`_fill_and_dispatch` para todos los campos de texto**: antes los campos mapeados usaban `send_keys` (que no dispara los eventos que React necesita); ahora usan el setter nativo de `HTMLInputElement` + `input`/`change`/`blur` synthetic events.
- **Espera adaptativa para SPAs standalone**: cuando la URL del formulario va directo en columna A (sin iframe en columna B), el engine espera hasta que haya al menos un `input`/`select` visible en el DOM (hasta 8s) antes de iniciar el discovery y el llenado.
- **Extra wait para URLs `gm_front`**: 3 segundos adicionales tras el scroll inicial para que React monte los componentes antes de escanear el formulario.

---

### Bloque 11 — Reversión visual de la demo (`interface_demo.py`)

Rediseño visual estilo Figma de la app en `interface_demo.py` (preview UX/UI, sin backend real). Es una maqueta interactiva para iterar el diseño; no reemplaza a la app productiva (`run.py`).

**Layout general**
- Se quitó el panel de **Consola** del pie; los logs van a `print` (stdout).
- Barra de título en **modo oscuro** (DWM `DwmSetWindowAttribute`) en la ventana principal y el modal de programación.
- Footer de créditos: "Made by Ariel Melgratti" centrado + "Some Updates by Elian Zás" a la derecha.

**Pestaña Envío de Leads**
- CTAs **Ejecutar** / **Ver Resultados** movidos a la cabecera de "DATOS POR PAÍS" (antes al pie).
- **Ejecutar** queda deshabilitado hasta seleccionar al menos un país.
- Checkbox **"Ver navegador mientras corre"** movido al card CONFIGURACIÓN (bajo DISPOSITIVOS), desmarcado por defecto y habilitado solo si hay Chrome/Firefox/Edge seleccionado.
- Campo **Email** deshabilitado hasta activar "Enviar mail".
- Tabla de datos: altura fija a ~5 filas visibles + scroll; **edición inline** de celdas (doble clic) con aviso efímero "estás editando el Excel, recordá Guardar"; **multi-selección** para clonar/eliminar varias filas.
- **Modal de ejecución**: modal real (bloquea la interfaz de atrás), muestra info del lead en proceso, aviso "No podés cerrar mientras se ejecuta", mensajes de email condicionales (enviado / falta destinatario / desactivado) y variante para test programado (aviso de que los tests programados posteriores igual corren).

**Card Test Automático (scheduler)**
- Estados como el original: **Sin configurar → Configurado → Activado**, con acciones "Programar test automático", "Iniciar ahora" y "Desactivar". CTA "Configurar automatización" en la cabecera junto al badge.
- Fondo distinto (morado más azulado + borde más visible) para que la sección resalte.
- **Modal de configuración** portado del `weekly_scheduler` original y recoloreado: horario personalizado (HH:MM) arriba; "Aplicar a otros días" **dinámico** (se copia al instante al tocar cada día o "Todos", con resalto 3D del botón seleccionado, sin botón "Copiar"); botón "Guardar".

**Pestaña Validación de Campos**
- Layout del original (`field_validation_ui.py`): preview de Excel de URLs arriba, expander "Configuración de ID" con formulario + botonera de 4 + filtros + tabla de reglas, manteniendo la paleta nueva. Multi-select para eliminar varias reglas.

**Pestaña Generar Excels con Datos**
- Selector de **MERCADO A GENERAR** (un Excel por mercado a la vez, varios dispositivos del mismo mercado).
- **Detección de país por las URLs** con aviso si no coincide con el mercado elegido.
- Botón **Borrar URLs** (reemplaza "Guardar modificaciones").

**Iconos**
- Reemplazo de emojis crudos por iconos estilo tabler generados a juego (play, stop, gear, download, bolt, report, link, monitor, eye) en `Asset/tabler_icons/`, aplicados a EJECUTAR, Ver Resultados, GENERAR EXCELS, REGENERAR DATOS, Borrar URLs, Ejecutar validación, Resultados, Configurar automatización, Programar/Iniciar/Desactivar, Detener y las pills del modal.

---

### Bloque 12 — Nueva pestaña "Comparador Dealers"

**41. Pestaña nueva completa**: reemplaza el bookmarklet JS manual que se pegaba en la consola del navegador para chequear dealers. Lee un Excel de dealers esperados (fila de encabezado y columnas configurables porque el Excel varía de país a país), navega región→ciudad→dealer en el form real vía Selenium, y compara contra el Excel filtrado. Módulos nuevos: `core/dealer_comparator_runner.py` (lógica pura, sin Tkinter) e `interface/dealer_comparator_ui.py` (la pestaña).

**42. Excel con o sin filtro**: toggle explícito "Tiene columna de filtro" / "No tiene filtro (usar todas las filas)" — antes había que dejar el campo vacío a mano y no quedaba claro que eso era válido.

**43. Condición "Buscar extras"**: se sumó como tercera opción del dropdown de condición (renombrada desde "Incluir + buscar extras" para que sea más clara), dispara automáticamente la búsqueda de dealers EXTRA (en el form pero no en el Excel filtrado, o en todo el Excel si no hay filtro) y DUPLICADOS (mismo dealer repetido dentro del `<select>` del form real — el Excel puede tener duplicados sin que importe, lo que se chequea es el form) sin necesitar tildar un checkbox aparte.

**44. Múltiples URLs de form en la misma pasada**: el cuadro de URLs acepta uno o varios forms (pares landing+form, o varios forms sueltos), igual formato que "Generar Excels con Datos" — se corren todos y el reporte final indica de qué URL vino cada fila.

**45. Detección de país por URL**: mismo criterio que "Generar Excels con Datos" — si las URLs pegadas no coinciden con el mercado seleccionado, aparece un aviso antes de ejecutar.

**46. Selector de modelo (solo T1)**: si el form tiene selector de modelo (id `models`), permite correr la comparación completa para "Todos los modelos" (detectados en vivo) o una lista específica separada por coma.

**47. Columnas adicionales a comprobar en el form**: reemplaza el concepto más críptico de "validaciones por data-attr" por algo directo — columna del Excel + id del campo HTML, compara el valor real de cualquier input/select del form contra esa columna.

**48. Tiempos de espera fijos con reintento**: no configurables por el usuario; primer intento rápido (0.6s) y, si no alcanzó, un segundo intento hasta 1.2s — cubre selects que tardan más en poblarse (ej. ciudades de un departamento grande).

**49. Modal de ejecución bloqueante**: mismo lenguaje visual que el overlay de "Envío de Leads" — al ejecutar, bloquea la ventana principal (`grab_set`) para evitar que el usuario edite campos sin querer mientras corre, con botón Detener y resumen final PASS/FAIL/EXTRA/DUPLICADO.

**50. "Ver navegador mientras corre"**: apagado por defecto (corre atrás sin robar el foco), igual criterio que "Envío de Leads".

**51. Resultados en carpeta propia**: reportes Excel (con color por estado) y capturas empaquetadas en ZIP quedan en `Dealerscheck_resultados/`, separado de `resultados/`. Las capturas sueltas no se dejan en disco una vez armado el ZIP.

**52. Configuraciones guardadas con nombre propio**: además del autoguardado por país, se puede guardar el mapeo completo de columnas con cualquier nombre (ej. "GMUY Livianos") para reusarlo sin reconfigurar.

**53. Ícono de la app y barra de tareas**: `AppUserModelID` propio (`SetCurrentProcessExplicitAppUserModelID`) para que Windows no agrupe la app bajo el ícono genérico de python.exe en la barra de tareas. `Asset/tabler_icons/` sumado a los datos empaquetados por PyInstaller (antes no viajaba en el build).

**54. Fix: duplicados se chequean contra el form, no contra el Excel**: la búsqueda de "DUPLICADO" comparaba filas repetidas dentro del propio Excel, lo cual no era el chequeo que importaba. Ahora detecta dealers que aparecen más de una vez como `<option>` en el mismo `<select>` del form real (junto con la búsqueda de EXTRA, recorriendo las mismas combinaciones región/ciudad). El Excel puede tener duplicados sin que afecte el resultado.

**55. Modal de "Envío de Leads": "mercado(s) completados" en vez de "lead(s) completados"**: el texto del modal de ejecución confundía "sesión" con "lead" cuando en realidad cada barra de progreso representa un mercado. Corregido el subtítulo y el badge de resumen.
