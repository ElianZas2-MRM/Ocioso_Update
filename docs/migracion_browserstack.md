# Migración de LambdaTest a BrowserStack

> **En curso.** Este documento se va actualizando a medida que avanza la migración. Lo
> que dice "pendiente" todavía no está hecho.

## Objetivo

Reemplazar LambdaTest por **BrowserStack Automate (web)** en tres dispositivos:

| Dispositivo | Navegador | Estado |
|---|---|---|
| iPhone 14 Pro Max (iOS 16) | Safari | ✅ llena, envía y detecta la TY |
| Samsung Galaxy S26 Ultra (Android 16) | Chrome | ✅ llena, envía y detecta la TY |
| Mac (macOS Sequoia) | Safari 18.4 | ✅ llena, envía y detecta la TY |

**Estado:** BrowserStack ya **reemplazó a LambdaTest en la app** (Envío de Leads, Generar
Excels, Reintentar fallidos, mails, CLI y runner autónomo). El código de `lambdatest_*`
sigue en el repo, sin uso, hasta retirarlo en un paso aparte.

## Decisiones tomadas

- **Automate web, no App Automate.** Todo lo de abajo corre sobre el plan Automate.
- **Sin el SDK de BrowserStack** (`browserstack.yml` + pytest). La app arma el driver a
  mano, igual que con LambdaTest, y se conecta directo al hub. Por eso BrowserStack
  muestra el cartel *"non SDK builds"*: es informativo, lo que se pierde son reportes
  avanzados que la app no usa.
- **Dispositivos fijos**, no regex (BrowserStack permite `"iPhone.*"`): para QA cada
  corrida tiene que caer en el mismo equipo y ser comparable.
- **El form se llena SIEMPRE dentro de la landing.** No vale abrir la URL del iframe
  aparte, ni leer el form en otro navegador para mapearlo.
- **Detección automática de campos**, como en Chrome/LambdaTest: nada de mapear a mano
  qué es cada campo.
- **Evidencia: el video** de la sesión. En el Excel va solo la columna del video (sin
  columna de dashboard).
- Proyecto en el dashboard: **`Osocio Form Automation [Local]`**. BrowserStack descarta
  los paréntesis, por eso los corchetes.

## Cómo está armado

```
osocio/providers/browserstack/
  runner.py           punto de entrada: run(pais, plataforma, excel, variante...) -> resumen
  dispositivos.py     tabla única iPhone BS / Android BS / Mac BS + tope de paralelo
  bs_driver.py        conexión: credenciales, dispositivos, driver, estado y video
  sesion.py           ciclo de vida de una sesión (estado por API REST, video)
  fabrica.py          le da al motor de la app un navegador de BrowserStack
  iphone_runner.py    FormFillerIphone: el motor de la app con el llenado nativo
  ios_nativo.py       la capa nativa del iPhone (ver abajo)
  prueba_conexion.py  herramienta de desarrollo contra una landing real
osocio/core/campos_inferencia.py  reconocer campos/botones por texto (neutro)
```

**Todo corre con el motor de los navegadores** (`BaseFormFiller`). `runner.run` arma
`GenericCountryBase(..., browser_factory=FabricaNavegadorBrowserStack(...), capturas=False)`
y corre el Excel como si fuera Chrome: mapping, IDs dinámicos (sólo si la celda está
vacía), `-` para omitir, documentos de Brasil normalizados, variantes CDC/T3 y carpetas
t1/t3, reintentos (event_id incluido), columnas de resultado y `fail_rows` son los de la
app, sin copia. Al terminar escribe la columna **Video**. El iPhone usa la subclase
`FormFillerIphone` (ver abajo).

Por qué no el motor de LambdaTest (`lt_runner`): un análisis mostró que no respetaba varias
reglas de los navegadores (IDs dinámicos le ganaban al Excel, ignoraba el `-`, regeneraba
documentos de Brasil, marcaba checkboxes opcionales, no conocía variantes ni t1/t3, no
reintentaba event_id, entre otras).

**Costo:** el motor de los navegadores fue escrito para un Chrome local y hace ~850
comandos por lead; en la nube cada uno tarda ~0,35 s → **~3–6 min por lead** (iPhone ~3,
Mac ~5,5, Android ~6). Optimizarlo es un paso pendiente (beneficia también a Chrome).

### `bs_driver.py`

- **Credenciales**: igual que LambdaTest. Primero `BROWSERSTACK_USERNAME` /
  `BROWSERSTACK_ACCESS_KEY`, si no `browserstack_credentials.txt` en la raíz
  (`username=` / `access_key=`). El archivo está en el `.gitignore`, igual que
  `browserstack.yml`.
- **`DISPOSITIVOS`**: la tabla de los tres equipos. Cambiar de modelo o versión es tocar
  ahí.
- **`platformName`** se manda explícito: `SafariOptions` trae `"mac"` de fábrica, y sin
  pisarlo el iPhone quedaba pedido como Mac.
- **Estado de la sesión**: `browserstack_executor: setSessionStatus` (el equivalente a
  `lambda-status`).
- **Video**: `GET api.browserstack.com/automate/sessions/<id>.json` → `video_url`.

### Resultados

Mismas carpetas y nombres que los navegadores (`core/variantes.py`):
`resultados/t1/resultados_Brasil_BS_Mac1.xlsx`, `…_CDC_BS_iPhone1.xlsx`,
`resultados/t3/Automatizacion_Brasil_T3_BS_Android1.xlsx`. Sin capturas: en *Formulario
Inserto / Completado / TY Page* dice "✓ ver Video" y el link está en **Video**.

### iPhone: por qué hace falta un motor aparte

Safari en iPhone **bloquea que la automatización entre a un iframe de otro dominio**, y
los forms (`secure-developments.com`) están dentro de las landings
(`chevrolet.com.xx`):

```
Blocked a frame with origin "https://www.chevrolet.com.co" from accessing a cross-origin frame
```

Es una restricción de Apple, no de BrowserStack ni del código. Lo que se probó:

| Vía | Resultado |
|---|---|
| Selenium `switch_to.frame()` en iOS 16 y 17 | ❌ bloqueado (iOS 18 no existe para el 14 Pro Max en BrowserStack) |
| Capability `disableCorsRestrictions` | ❌ solo existe para Safari de Mac |
| Playwright sobre iPhone real | ❌ BrowserStack solo le expone la landing, el iframe no existe para él |
| JavaScript desde la landing | ❌ bloqueado por la misma regla |
| **Capa nativa del iPhone (`NATIVE_APP`)** | ✅ ve el form dentro de la landing, sin importar el dominio |

### iPhone: `FormFillerIphone` + `ios_nativo.py`

`FormFillerIphone` es el motor de la app con un solo cambio de fondo: en vez de entrar al
iframe (`_entrar_al_iframe`, un punto de extensión de `BaseFormFiller`) se queda en la
landing y llena por la capa nativa. Reemplaza sólo los métodos que tocan el DOM del iframe
(llenado, envío/TY, espera de recarga por event_id, foto para "Datos vs Excel"). Si el form
**no** está en un iframe, no hay bloqueo y usa el motor normal por DOM.

El valor de cada campo sale del mismo camino que en Chrome: la etiqueta visible se traduce a
los ids del mapping del país (`campos_inferencia.ids_para_etiqueta`) y de ahí
`_resolve_field_value`, `_pide_omitir_campo`, IDs dinámicos, `_normalizar_documento_brasil`,
`_sanitize_peru_document` y la regla de no mapeados (obligatorio → plausible, opcional → 50%).

#### La capa nativa (`ios_nativo.py`)

La capa nativa ve la pantalla como la ve el sistema del iPhone. No ve ids, pero cada
control trae **su etiqueta visible** adentro ("Nombre", "Cédula", "Ciudad"…):

| Control | Cómo lo ve el iPhone |
|---|---|
| Texto | `XCUIElementTypeTextField`, con un `StaticText` (la etiqueta) adentro |
| Select | `XCUIElementTypeOther` cuyo `value` es la opción elegida ("Seleccionar" al principio). Al tocarlo, iOS 16 abre un menú con un botón por opción |
| Checkbox | `XCUIElementTypeSwitch`, `value` `0`/`1` |
| Enviar / Siguiente | `XCUIElementTypeButton` |
| El iframe | un contenedor `Other` con el `title` del iframe ("RQA new form") |

Qué hace con cada cosa, **con la lógica de la app** (no reglas propias):

- **Qué dato va en cada campo**: la etiqueta pasa por `_infer_data_key` /
  `_get_inferred_value` (`lt_runner`), los mismos que usa la app con los ids. Antes se
  busca la columna exacta del Excel (CPF, CEP…). "Sobrenome" se resuelve como apellido
  antes que como nombre, porque contiene "nome".
- **Datos**: un lead de `generar_fila_datos(pais, device="iphone")`, el generador de
  Excels. El país sale de la URL del form (`/commonwealth/colombia/`) o del dominio.
- **Selects**: opción del Excel si viene, si no **una real al azar**. Los placeholders se
  descartan con `_is_placeholder` (nunca "Seleccionar", "Seleccione", "Escolha"…).
- **Dependencias** (región → ciudad → concesionario): se llena en orden visual. Un hijo
  deshabilitado se espera hasta que se habilite, con los mismos tiempos que
  `desplegables.py`: 8 s, revisando cada 0,2 s, 2 reintentos.
- **Campos que el lead no conoce**: obligatorio (`*` en la etiqueta) → valor plausible
  siempre; opcional → 50%, una vez por campo (README, "Campos que no están en el Excel").
- **Checkbox**: se marca si es de términos/privacidad o obligatorio, y se verifica que
  quede en `1`.
- **Forms en pasos**: avanza con los mismos textos de botón que la app
  (`_NEXT_KEYWORDS`, `_SUBMIT_KEYWORDS`).
- **TY**: mismo criterio que `_ty_visible` de la app: el form **reemplazado** por otro
  contenido (dentro del contenedor del iframe ya no hay campos ni botones y hay texto).
  Textos de agradecimiento como respaldo. El texto de la TY queda en el log.
- **Error de event_id**: textos de `BaseFormFiller._EVENT_ID_PATTERNS`. **Solo en forms
  `gm_forms`** (en `gm_front`/`gm_frontend` no ocurre) se reintenta una vez como en los
  navegadores locales: esperar que el form se recargue solo, volver a llenar y reenviar.
  Si falla de nuevo: *"Error event_id: formulario no pudo enviarse por falta de event_id
  (2 intentos fallidos)"*.
- **Sin TY**: deja en el log qué textos, campos y botones quedaron en el form, para poder
  revisar y reintentar la corrida.

Detalles técnicos que costaron encontrar:

- En la capa nativa **no hay JavaScript**: `is_displayed()` / `get_attribute()` de
  Selenium inyectan JS y fallan con *"Method is not implemented"*. Se lee `page_source`
  (XML, una llamada por pantalla) y se usa `get_dom_attribute()` / `rect`.
- Los comandos de contexto (`getContexts`, `switchToContext`) son de Appium y Selenium no
  los trae: se registran a mano.
- **Safari hace zoom al tipear** y cambian posiciones y tamaños. Por eso no se calcula
  nada de antemano: después de cada acción se relee la pantalla, y lo pendiente se decide
  por lo que *muestra* el control (vacío, placeholder, `0`).
- Las lecturas de valor por la capa nativa no siempre son confiables: la referencia es el
  video.

## Cómo probar

Desde la app: *Envío de Leads* → tildar **iPhone BS / Android BS / Mac BS** → cargar las
credenciales en el panel **CREDENCIALES BROWSERSTACK** → ejecutar. Necesita el Excel del
dispositivo (`…_Iphone.xlsx`, `…_Android.xlsx`, `…_Mac.xlsx`), que se genera en *Generar
Excels con Datos*.

Por línea de comandos (con envío, leads REALES en producción):

```powershell
python run.py --run-browserstack iphone --pais Colombia
python run.py --run-browserstack mac --pais Brasil --excel-suffix _CDC
```

Herramienta de desarrollo, para una landing puntual (sólo envía con `--enviar`):

```powershell
python -m osocio.providers.browserstack.prueba_conexion --url https://... --form-url https://...
```

Tests sin red: `pytest tests/test_estructura_browserstack.py tests/test_browserstack_integracion.py`.
**En pruebas, máximo 2 sesiones de BrowserStack a la vez.**

## Resultados hasta ahora

Con el motor de LambdaTest (primeras pruebas) y, desde la integración, con el motor de la
app (`runner.run`): Colombia `raq` en iPhone, Android y Mac.

| Form | Tipo | iPhone | Android | Mac |
|---|---|---|---|---|
| Colombia `raq` | gm_front | ✅ | ✅ | ✅ |
| Bolivia `model-overviews` (Tracker) | gm_front | ✅ | ✅ | ✅ |
| Bolivia `raq` | gm_front | ✅ | pendiente | pendiente |
| Brasil `myr2026` (S10) | gm_forms | ✅ (con reintento de event_id) | ✅ | ✅ |

## Plan de BrowserStack

`Automate - Desktop & Mobile Pro`: **10 sesiones en paralelo por cuenta** (compartidas
por todo el equipo, no por build) y 10 más en cola. Se probaron 6 a la vez dentro del
mismo build sin cola: se puede correr, por ejemplo, un build por mercado con sus forms en
paralelo. Conviene un tope configurable para no acaparar las 10. **En pruebas, máximo 2
sesiones a la vez.**

## Cambios fuera de `browserstack/`

- `base_form_filler.py`: `_EVENT_ID_PATTERNS` reconoce el cartel nuevo de Brasil
  (*"Não conseguimos concluir seu envio"*). Afecta también a Chrome, Firefox y Edge.
- `lt_runner.py`, **solo el camino de Mac/Safari desktop** (Android e iPhone no cambian):
  - `_fill_text_sendkeys`: Safari 18 (BrowserStack) no le da el foco a un campo de un
    iframe de otro dominio al hacer `send_keys`, y las teclas no llegaban: el campo quedaba
    vacío sin error, y después lo pisaba el relleno de "requerido sin dato" (el lead
    viajaba con apellido y email inventados). Ahora se enfoca antes
    (`focus({preventScroll:true})`) y, si igual no quedó, se carga con el setter de React.
  - Scroll: en desktop ya no se baja 90 px por campo (se acumulaba hasta dejar el footer en
    el video). Como en los navegadores locales, **solo se scrollea si el campo no se ve**
    (`_centrar_landing_en_campo`); eso también corrige cuando la landing salta sola al
    footer, cosa que pasa a veces. Mobile sigue scrolleando a medida que llena.
- `.gitignore`: `browserstack_credentials.txt` y `browserstack.yml`.
- `base_form_filler.py`: `_entrar_al_iframe()` (punto de extensión para el iPhone) y
  `capturas=False` → `ScreenshotManagerNulo`.
- `variantes.py`: nombres `BS_iPhone` / `BS_Android` / `BS_Mac` para los resultados.
- `excel_layout.py`: columna **Video**.
- UI (`main_interface.py`, `envio_leads.py`, `helpers_interface.py`,
  `dealer_comparator_ui.py`), `run.py` (`--run-browserstack`; `--run-lambdatest` sigue
  aceptado y va a BrowserStack), `autonomous_runner.py` (las programaciones guardadas con
  LambdaTest van a su equivalente), `build.bat` (plantilla `browserstack_credentials.txt`).

## Pendiente

- [x] Bolivia y Brasil en Android y Mac.
- [x] Integración a la app: UI, CLI, runner autónomo, mails, columna "Video", resultados,
      build.
- [x] Código común fuera de `lambdatest_mac/` (`core/campos_inferencia.py`).
- [x] iPhone con las reglas de datos de la app (`-`, Excel antes que IDs dinámicos,
      documentos de Brasil, no mapeados).
- [ ] Probar desde la app (UI) los tres dispositivos, un T3/CDC de Brasil y "Reintentar
      fallidos".
- [ ] **Optimizar el motor en remoto** (menos comandos por lead).
- [ ] Más mercados en iPhone, sobre todo Perú (tipo de documento) y un form en pasos.
- [ ] iPhone: columnas SI/NO de checkboxes opcionales (la capa nativa no ve el `name` del
      checkbox) y envío vacío inicial.
- [ ] Probar el `.exe` en otra PC con las credenciales de otra persona.
- [ ] Retirar LambdaTest cuando BrowserStack esté validado.
