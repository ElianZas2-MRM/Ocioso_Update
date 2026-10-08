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

Mientras dure la migración, LambdaTest **no se toca**: los dos proveedores conviven y
BrowserStack vive en su propia carpeta.

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
  bs_driver.py        conexión: credenciales, dispositivos, driver, estado y video
  ios_nativo.py       motor del iPhone (capa nativa, ver abajo)
  prueba_conexion.py  prueba de punta a punta contra una landing real
```

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

### Android y Mac: el flujo de la app, tal cual

En estos dos se puede entrar al iframe con `switch_to.frame()`, así que se usa
**exactamente** el flujo de LambdaTest: `lt_runner._run_single_lead`, el mismo que
llaman `run_lt_batch` (Mac) y `run_lt_android_batch` (Android) por cada lead. Solo
cambia el driver. Mapping por id, IDs dinámicos, dependencias, checkboxes, forms en
pasos, puerta RAQ de Brasil, documentos de Brasil, TY y "Datos vs Excel" vienen solos.

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

### iPhone: cómo funciona `ios_nativo.py`

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

Desde la raíz del proyecto, con las credenciales cargadas:

```powershell
# iPhone, form de Colombia, SIN enviar
python -m osocio.providers.browserstack.prueba_conexion

# con envío (lead REAL en producción, email @mrm.com)
python -m osocio.providers.browserstack.prueba_conexion --enviar

# otra plataforma / otra landing
python -m osocio.providers.browserstack.prueba_conexion --plataforma android --enviar
python -m osocio.providers.browserstack.prueba_conexion --url https://... --form-url https://secure-developments.com/...
```

`--enviar` es lo único que manda un lead. Tests sin red:
`pytest tests/test_estructura_browserstack.py`.

## Resultados hasta ahora

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

## Pendiente

- [x] Bolivia y Brasil en Android y Mac.
- [ ] Más mercados en iPhone, sobre todo Perú (tipo de documento) y un form en pasos.
- [ ] ¿Reintento de event_id en `gm_forms` también para Android/Mac? Hoy siguen a
      LambdaTest: sin reintento.
- [ ] Integración a la app: UI, CLI, runner autónomo, Excel con la columna "Video",
      carpetas de resultados, build (`.spec`).
- [ ] Sacar el código común de `lambdatest_mac/` a un módulo neutro, para que
      BrowserStack no dependa de archivos que se llaman "lambdatest".
- [ ] En el iPhone, lo que depende del Excel real: `-` para omitir campos, columnas
      SI/NO de checkboxes opcionales, IDs dinámicos, envío vacío inicial para capturar
      errores.
- [ ] Retirar LambdaTest cuando BrowserStack esté validado.
