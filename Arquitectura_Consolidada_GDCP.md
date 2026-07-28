# Arquitectura Consolidada del Sistema GDCP

## 1. Explicación general del sistema

### Qué hace el proyecto

Este sistema automatiza la carga y validación de formularios web de distintos países usando Selenium, con una interfaz de escritorio en Tkinter para operar el flujo sin tocar código.

El proyecto permite:

- crear y mantener los Excels de entrada por país,
- ejecutar formularios reales en distintos navegadores y viewports,
- capturar evidencia visual del proceso,
- guardar resultados en Excel,
- parametrizar IDs fijos y dinámicos sin editar el código base,
- validar reglas de campos desde una pestaña específica de QA,
- programar ejecuciones automáticas y enviar resultados por email.

### Cómo funciona de punta a punta

En términos prácticos, el sistema toma un Excel de entrada por país, abre una landing real, encuentra el iframe correcto del formulario, dispara validaciones visuales iniciales, completa los campos visibles paso a paso, envía el lead, toma screenshots del recorrido y deja un Excel de resultados con el estado de cada fila.

La interfaz gráfica es la puerta de entrada operativa. Desde ahí se crean Excels, se editan datos, se lanzan formularios por país, se administran mappings/IDs y se ejecutan validaciones. La ejecución real del navegador y del formulario vive en la capa `core/`, especialmente en `BaseFormFiller`.

### Flujo principal step by step

1. El usuario inicia la app con `run.py`.
2. `run.py` asegura carpetas de runtime y abre la UI con `interface.main_interface.iniciar_interfaz()`.
3. Desde la UI se selecciona país, navegador y viewport.
4. La UI carga dinámicamente el módulo `forms/Formulario_<Pais>_Main.py` y ejecuta su función `run_formularios_<Pais>()`.
5. El runner del país instancia la clase base del país definida en `core/Formulario_<Pais>_Base.py`, que hereda de `core/base_form_filler.py`.
6. `BaseFormFiller.run()` prepara carpetas de resultados, crea navegador, abre el Excel de entrada y recorre fila por fila.
7. Para cada fila:
   - abre la landing,
   - maneja cookies,
   - encuentra el iframe correcto del formulario,
   - captura errores con el formulario vacío,
   - completa campos visibles según mapping fijo, IDs dinámicos y dependencias,
   - navega pasos intermedios si el formulario es multipaso,
   - marca radios/checkboxes requeridos,
   - envía el formulario,
   - detecta éxito o error,
   - guarda screenshots y resultados en Excel.
8. Si está habilitado, se envía un email con el Excel y las evidencias comprimidas.
9. Opcionalmente, un proceso aparte (`iniciar_ejecutor_autonomo.py` / `json/ejecutor_autonomo.py`) monitorea programaciones y ejecuta corridas automáticas.

### Qué componentes participan en cada etapa

| Etapa | Componentes principales | Rol |
|---|---|---|
| Inicio de aplicación | `run.py`, `interface/main_interface.py` | Arranque, UI, selección de ejecución |
| Gestión de datos | `interface/helpers_interface.py`, `utils/fixed_field_mapping_store.py`, `data/`, `json/` | Crear/sincronizar Excels, persistir config y mappings |
| Ejecución por país | `forms/Formulario_*_Main.py`, `core/Formulario_*_Base.py` | Adaptar configuración país + lanzar flujo |
| Automatización central | `core/base_form_filler.py`, `core/browser_manager.py`, `core/screenshot_manager.py` | Navegación Selenium, llenado, envío, evidencia |
| Reglas de negocio de campos | `json/fixed_field_mappings.json`, `json/ids_dinamicos.json`, `core/field_dependencies.py` | Determinar qué llenar y en qué orden |
| Validación QA | `interface/field_validation_ui.py`, `validation/selenium_validation_runner.py`, `validation/error_message_validator.py`, `validation/validation_exporter.py` | Ejecutar pruebas de campos y exportar resultados |
| Automatización programada | `utils/scheduling.py`, `json/ejecutor_autonomo.py` | Guardar agenda y ejecutar corridas futuras |
| Notificación | `interface/helpers_interface.py`, `validation/validation_email.py` | Envío de emails y adjuntos |

### Dónde está la lógica central del sistema

La lógica central está concentrada en cuatro zonas:

1. `core/base_form_filler.py`
   Es el núcleo real de la automatización. Ahí vive el flujo de Selenium, el llenado step-by-step, la resolución de mappings, dependencias, IDs dinámicos, captura de errores, screenshots y guardado de resultados.

2. `interface/main_interface.py`
   Es la orquestación operativa. No ejecuta la lógica del navegador, pero decide qué módulo cargar, qué entorno usar, qué Excel abrir y cómo dispara los flujos desde UI.

3. `utils/fixed_field_mapping_store.py`
   Centraliza la configuración efectiva por país. Mezcla defaults definidos en `core/Formulario_*_Base.py` con overrides persistidos en JSON. Es la pieza que evita editar código base para cambios operativos.

4. `validation/selenium_validation_runner.py`
   Es el motor de QA automatizado. Recorre formularios como un usuario, dispara validaciones, compara errores esperados vs reales y exporta evidencia estructurada.

## 2. Mapa de responsabilidades

### Capa de entrada / operación

- `run.py`
  Punto de entrada principal de la app de escritorio.
- `interface/main_interface.py`
  Construye la UI, carga módulos por país, deja editar Excels, configura ejecución y programación.
- `interface/field_validation_ui.py`
  Construye la pestaña de validación de campos y dispara el runner de QA.

### Capa de automatización de formularios

- `forms/`
  Runners por país. Son wrappers de entrada que traducen una acción de UI a una ejecución concreta.
- `core/Formulario_*_Base.py`
  Configuración específica de cada país: Excel, país, mapping de campos, required fields y particularidades.
- `core/base_form_filler.py`
  Motor común reutilizado por todos los países.
- `core/browser_manager.py`
  Crea el WebDriver correcto usando drivers locales.
- `core/screenshot_manager.py`
  Genera las capturas del flujo.

### Capa de configuración y persistencia

- `json/config_global.json`
  Config global de envío y comportamiento.
- `json/fixed_field_mappings.json`
  Overrides persistidos de mappings fijos por país.
- `json/ids_dinamicos.json`
  IDs dinámicos, valores fijos y dependencias configurables.
- `utils/fixed_field_mapping_store.py`
  Resuelve la configuración final efectiva que luego usa el motor.
- `utils/scheduling.py`
  Guarda/carga la programación de ejecuciones automáticas.

### Capa de validación / QA

- `validation/selenium_validation_runner.py`
  Ejecuta validaciones Selenium sobre campos visibles y dropdowns.
- `validation/error_message_validator.py`
  Genera inputs de prueba y decide qué error debería ser el principal según reglas.
- `validation/text_field_validator.py`
  Valida comportamiento de inputs a nivel regex/caracter.
- `validation/validation_exporter.py`
  Exporta el resultado a Excel con formato legible y foco en errores.
- `validation/validation_email.py`
  Envía el reporte de validación.

### Capa de automatización programada

- `iniciar_ejecutor_autonomo.py`
  Script de arranque del scheduler.
- `json/ejecutor_autonomo.py`
  Loop residente que revisa si existe una programación pendiente y dispara múltiples corridas.

### Cómo se conectan entre sí

La conexión entre módulos sigue esta secuencia:

- UI -> `forms/*_Main.py`
- `forms/*_Main.py` -> `core/Formulario_*_Base.py`
- `core/Formulario_*_Base.py` -> `core/base_form_filler.py`
- `base_form_filler.py` -> `browser_manager.py`, `screenshot_manager.py`, `utils/fixed_field_mapping_store.py`, `json/ids_dinamicos.json`
- UI de validación -> `validation/selenium_validation_runner.py` -> `validation/error_message_validator.py` / `validation_exporter.py`
- UI de programación -> `utils/scheduling.py` -> `json/ejecutor_autonomo.py` -> `forms/*_Main.py`
- Ejecuciones -> `interface/helpers_interface.py` para análisis de resultados, compresión y email.

## 3. Funciones y métodos más importantes

Se listan solo las piezas que cambian realmente el comportamiento del sistema o coordinan flujos completos.

### `iniciar_interfaz`

- Archivo: `interface/main_interface.py`
- Qué hace: construye la aplicación Tkinter, pestañas, controles, tablas y acciones principales.
- Inputs: no recibe parámetros.
- Outputs: abre la UI y deja registrado el flujo operativo del usuario.
- Por qué es importante: es la puerta de entrada humana del sistema. Sin esta función no existe operación diaria del proyecto.
- Ejemplo breve: al ejecutar `python run.py`, esta función levanta la app y habilita la ejecución de formularios por país.

### `ejecutar_script_configurable`

- Archivo: `interface/main_interface.py`
- Qué hace: toma país, navegador y viewport elegidos en UI, carga dinámicamente el módulo de formularios del país y ejecuta su runner.
- Inputs: `nombre_script_base`, `selected_browser`, `selected_viewport`.
- Outputs: dispara un hilo de ejecución del formulario seleccionado.
- Por qué es importante: es el puente entre la interfaz y el motor real de automatización.
- Ejemplo breve: si el usuario elige Argentina + Chrome + mobile, resuelve `Formulario_Argentina_Main.py` y corre `run_formularios_Argentina(browser="chrome", viewport="600x738")`.

### `run_formularios_<Pais>`

- Archivo: `forms/Formulario_Argentina_Main.py` y equivalentes por país.
- Qué hace: inicializa la clase del país, prepara el ciclo de lectura del Excel, procesa cada fila y opcionalmente envía email al final.
- Inputs: `browser`, `viewport`, `headless`, `enviar_email`.
- Outputs: genera Excel de resultados, screenshots y eventualmente envío de correo.
- Por qué es importante: es el entry point ejecutable por país, tanto desde la UI como desde el scheduler.
- Ejemplo breve: `run_formularios_Argentina("chrome", "fullscreen")` procesa `Lead_information_Formulario_Argentina_Main.xlsx`.

### `load_effective_country_form_config`

- Archivo: `utils/fixed_field_mapping_store.py`
- Qué hace: combina la configuración base de un país con los overrides guardados en `json/fixed_field_mappings.json`.
- Inputs: `country_name`, `fallback_config` opcional.
- Outputs: dict de configuración efectiva con `field_mapping`, `required_fields`, `excel_file`, etc.
- Por qué es importante: desacopla la operación del código fuente. Permite cambiar mappings sin editar los `Base.py`.
- Ejemplo breve: si un ID fijo cambió en producción, esta función hace que el motor use el override persistido y no el default hardcodeado.

### `build_excel_columns_for_country`

- Archivo: `utils/fixed_field_mapping_store.py`
- Qué hace: arma los encabezados del Excel según el mapping efectivo del país.
- Inputs: `country_name`.
- Outputs: lista de columnas, empezando por `URL` y `Formulario`.
- Por qué es importante: sincroniza la estructura del Excel con la lógica real de captura/llenado.
- Ejemplo breve: cuando se crea por primera vez el Excel de un país, esta función determina qué columnas verá el usuario.

### `abrir_excel`

- Archivo: `interface/helpers_interface.py`
- Qué hace: crea el Excel si no existe, lo sincroniza si la estructura cambió y luego lo abre.
- Inputs: `nombre_archivo`.
- Outputs: abre el archivo Excel en el sistema.
- Por qué es importante: reduce errores operativos. El usuario no necesita preparar manualmente la estructura de entrada.
- Ejemplo breve: si falta el Excel de Paraguay, lo crea con columnas acordes al mapping efectivo actual.

### `run`

- Archivo: `core/base_form_filler.py`
- Qué hace: ejecuta el proceso end-to-end para un país completo, recorriendo todas las filas del Excel.
- Inputs: usa la configuración del objeto (`self.config`) y el Excel asociado al país.
- Outputs: Excel de resultados, screenshots, logs y cierre ordenado del navegador.
- Por qué es importante: es el orquestador principal del motor Selenium.
- Ejemplo breve: toma cada fila, abre landing, encuentra formulario, llena datos, envía lead y guarda el resultado.

### `process_landing_page`

- Archivo: `core/base_form_filler.py`
- Qué hace: navega a la landing, espera carga completa, maneja popups/cookies y toma la captura inicial.
- Inputs: `landing_url`, `ss_counter`.
- Outputs: nombre del screenshot de la landing.
- Por qué es importante: estandariza la entrada al flujo y evita errores tempranos por contenido dinámico o cookies.

### `find_and_position_to_form`

- Archivo: `core/base_form_filler.py`
- Qué hace: localiza el iframe correcto del formulario esperado y posiciona la vista sobre él.
- Inputs: `expected_form_url`.
- Outputs: referencia al iframe objetivo o `None`.
- Por qué es importante: si falla esta etapa, no existe automatización posible. Es la bisagra entre landing y formulario real.

### `capture_error_messages`

- Archivo: `core/base_form_filler.py`
- Qué hace: intenta enviar el formulario vacío para capturar mensajes de error iniciales.
- Inputs: `ss_counter`.
- Outputs: nombre del screenshot de errores o marcador de no capturado.
- Por qué es importante: guarda evidencia base de validación visual y documenta el estado inicial del formulario.

### `fill_form_fields_auto_step`

- Archivo: `core/base_form_filler.py`
- Qué hace: completa campos visibles, maneja formularios de uno o varios pasos y avanza automáticamente con botón siguiente hasta el paso final.
- Inputs: `form_data`.
- Outputs: nombre del screenshot del formulario completado.
- Por qué es importante: es uno de los métodos más críticos del proyecto. Resuelve el flujo multipaso sin configuración extra por formulario.
- Ejemplo breve: en un formulario con región -> ciudad -> concesionario, completa el padre primero, espera el dropdown hijo y recién después avanza.

### `_fill_visible_fields_from_mapping`

- Archivo: `core/base_form_filler.py`
- Qué hace: llena solo los campos visibles en el DOM actual, aplicando dependencias, selects, IDs dinámicos y fallback de texto.
- Inputs: `form_data`, `dependencies`.
- Outputs: no retorna dato funcional; muta el formulario y registra campos aplicados.
- Por qué es importante: concentra la lógica fina de cómo se decide qué completar en cada pantalla.

### `_auto_fill_unmapped_dropdowns`

- Archivo: `core/base_form_filler.py`
- Qué hace: detecta selects e inputs visibles que no están en el mapping y los completa automáticamente con opciones válidas o valores dinámicos.
- Inputs: `field_mapping` opcional.
- Outputs: `True/False` indicando si completó algo.
- Por qué es importante: le da resiliencia al sistema frente a cambios parciales del DOM o campos nuevos no mapeados aún.

### `_sync_tracked_with_dom_before_submit`

- Archivo: `core/base_form_filler.py` (equivalente en LambdaTest: `_sync_tracked_with_dom` en `lambdatest_mac/lt_runner.py`, que usan Mac y Android).
- Qué hace: justo antes del click en Enviar relee el DOM completo en un solo `execute_script` y pisa el tracking con el valor efectivo de cada campo.
- Inputs: `form_data` de la fila.
- Outputs: actualiza `current_row_field_values` y setea `_datos_vs_excel` / `_datos_mismatch`.
- Por qué es importante: el tracking por paso guarda lo que se *intentó* escribir. Si el form re-renderiza, un dropdown se resetea, o un reintento vuelve a sortear modelo/ciudad/concesionario, lo trackeado deja de ser lo que viaja en el lead y la comparación posterior contra la base de datos queda mal. Este paso garantiza que el Excel refleje siempre el último valor real. Además compara los dropdowns contra lo pedido en el Excel y avisa (columna `Datos vs Excel`, en ámbar) si quedó otra cosa.

### `submit_and_verify_form`

- Archivo: `core/base_form_filler.py`
- Qué hace: envía el formulario, busca errores visibles y, si el lead parece enviado, captura la TY page.
- Inputs: `current_ss_number`, `expected_form_url`.
- Outputs: tupla `(resultado_texto, ty_page_name)`.
- Por qué es importante: define el resultado final funcional de cada fila.

### `write_tracked_fields_to_sheet`

- Archivo: `core/base_form_filler.py`
- Qué hace: escribe en el Excel de salida los valores efectivamente usados para campos mapeados y dinámicos.
- Inputs: `sheet`, `row_index`.
- Outputs: actualiza la hoja Excel.
- Por qué es importante: deja trazabilidad de qué valor se aplicó realmente, útil para debugging y auditoría.

### `create_browser`

- Archivo: `core/browser_manager.py`
- Qué hace: crea el WebDriver correcto según navegador, viewport y modo headless usando drivers locales.
- Inputs: `browser_type`, `viewport`, `headless`.
- Outputs: instancia de Selenium WebDriver.
- Por qué es importante: abstrae la infraestructura del navegador y evita acoplar cada runner a Selenium puro.

### `run_field_validations`

- Archivo: `validation/selenium_validation_runner.py`
- Qué hace: abre landing e iframe, recorre steps, detecta campos visibles, dispara errores, prueba reglas y arma filas de resultado estructuradas.
- Inputs: `field_mapping`, URLs, navegador, viewport y parámetros de ejecución.
- Outputs: estructura con `rows`, `fields`, `error_rows`, `unmapped_ids`, `summary`.
- Por qué es importante: es el motor de QA del producto, separado del flujo de carga de leads.
- Ejemplo breve: puede validar si un campo acepta caracteres inválidos, si muestra el mensaje correcto y si un dropdown dependiente exige contexto previo.

### `generar_inputs_test`

- Archivo: `validation/error_message_validator.py`
- Qué hace: construye inputs de prueba a partir de reglas como `required`, `min_length`, `max_length`, `invalid_chars` y patrones.
- Inputs: `config` del campo.
- Outputs: lista de casos de prueba.
- Por qué es importante: convierte reglas declarativas en casos ejecutables de QA.

### `export_validation_results`

- Archivo: `validation/validation_exporter.py`
- Qué hace: exporta resultados de validación a Excel, ordena errores primero y resalta visualmente las fallas.
- Inputs: `validation_result`, `output_path` opcional.
- Outputs: archivo Excel de resultados de validación.
- Por qué es importante: transforma la salida técnica del runner en un artefacto usable por QA o negocio.

### `guardar_programacion` / `cargar_programacion`

- Archivo: `utils/scheduling.py`
- Qué hacen: persisten y recuperan la programación futura de tests automáticos.
- Inputs: dict `programacion` en el guardado.
- Outputs: `True/False` o un dict con fecha, países, navegadores y viewports.
- Por qué son importantes: habilitan la ejecución desatendida sin depender de memoria en runtime.

### `ejecutar_tests`

- Archivo: `json/ejecutor_autonomo.py`
- Qué hace: recorre la programación pendiente y ejecuta todos los scripts por país/navegador/viewport, recopilando resultados generados.
- Inputs: `programacion`.
- Outputs: lista de resultados de ejecución con rutas a Excel y screenshots.
- Por qué es importante: es el corazón del modo autónomo.

### `main`

- Archivo: `json/ejecutor_autonomo.py`
- Qué hace: loop residente que verifica cada minuto si debe lanzar una programación.
- Inputs: no recibe parámetros.
- Outputs: dispara ejecución automática y limpieza de programación obsoleta o cumplida.
- Por qué es importante: mantiene vivo el scheduler del sistema.

### `enviar_email_resultados` / `enviar_email_resultados_consolidados`

- Archivo: `interface/helpers_interface.py`
- Qué hacen: analizan el Excel generado, preparan adjuntos y encolan el envío por Outlook.
- Inputs: rutas de Excel, screenshots y lista de resultados según el caso.
- Outputs: `True/False` de éxito.
- Por qué son importantes: cierran el ciclo operativo y vuelven observable el resultado fuera de la aplicación.

## 4. Flujo resumido con funciones

Este es el recorrido principal del código, en el orden en que un desarrollador nuevo debería entenderlo:

1. `run.py`
   Arranca la aplicación y delega en `iniciar_interfaz()`.

2. `interface.main_interface.iniciar_interfaz()`
   Levanta la UI y expone acciones por país.

3. `interface.main_interface.ejecutar_script_configurable()`
   Traduce una acción del usuario en una ejecución concreta de runner por país.

4. `forms/Formulario_<Pais>_Main.py::run_formularios_<Pais>()`
   Instancia la clase del país y pone en marcha el proceso.

5. `utils.fixed_field_mapping_store.load_effective_country_form_config()`
   Resuelve el mapping que realmente se usará en runtime.

6. `core.base_form_filler.run()`
   Recorre filas del Excel y coordina el procesamiento end-to-end.

7. `core.base_form_filler.process_landing_page()`
   Abre la landing, espera carga, maneja cookies y toma la captura inicial.

8. `core.base_form_filler.find_and_position_to_form()`
   Encuentra el iframe del formulario correcto.

9. `core.base_form_filler.capture_error_messages()`
   Dispara evidencia inicial de validación con formulario vacío.

10. `core.base_form_filler.fill_form_fields_auto_step()`
    Ejecuta el llenado real del formulario.

11. `core.base_form_filler._fill_visible_fields_from_mapping()`
    Decide qué campo completar en el step actual.

12. `core.base_form_filler._auto_fill_unmapped_dropdowns()`
    Absorbe campos no mapeados visibles para mejorar robustez.

13. `core.base_form_filler._sync_tracked_with_dom_before_submit()`
    Relee el DOM antes de enviar y deja en el tracking el valor real de cada campo.

14. `core.base_form_filler.submit_and_verify_form()`
    Envía el lead y determina si hubo éxito o error.

15. `core.base_form_filler.write_tracked_fields_to_sheet()`
    Deja trazabilidad de valores efectivamente usados.

15. `interface.helpers_interface.enviar_email_resultados()`
    Si está habilitado, distribuye el resultado al finalizar.

## Observaciones útiles para un desarrollador nuevo

- La UI es importante para operar, pero no contiene la lógica central de automatización.
- El comportamiento real del llenado se decide en `core/base_form_filler.py`.
- Los `forms/*_Main.py` son runners, no el núcleo del negocio.
- Los `core/Formulario_*_Base.py` definen defaults por país y sirven como contrato de configuración.
- Los cambios operativos de mapping ya no deberían resolverse editando `Base.py` si pueden persistirse en `json/fixed_field_mappings.json`.
- Los IDs dinámicos y ciertas dependencias viven en `json/ids_dinamicos.json`; no todo está hardcodeado.
- El módulo de validación es un subsistema separado, útil para QA y descubrimiento de problemas del DOM.
- El scheduler autónomo no reemplaza la UI: la complementa para corridas programadas.

## Punto de partida recomendado para onboarding

Si un desarrollador nuevo tiene poco tiempo, el orden recomendado para leer el proyecto es:

1. `run.py`
2. `interface/main_interface.py`
3. `forms/Formulario_Argentina_Main.py` como ejemplo representativo
4. `core/Formulario_Argentina_Base.py`
5. `core/base_form_filler.py`
6. `utils/fixed_field_mapping_store.py`
7. `interface/helpers_interface.py`
8. `validation/selenium_validation_runner.py`
9. `json/ejecutor_autonomo.py`

Con esa secuencia se entiende casi todo el sistema sin necesidad de recorrer cada archivo país por país.