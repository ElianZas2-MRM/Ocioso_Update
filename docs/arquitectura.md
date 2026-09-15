# Arquitectura

Qué hace cada parte y por dónde pasa una corrida. Para el detalle de cómo levantar el
entorno, ver [`desarrollo.md`](desarrollo.md).

## Qué es esto

Una app de escritorio en Tkinter que automatiza el llenado y envío de formularios de GM en
9 mercados de Latinoamérica, y verifica que el lead haya entrado. Se distribuye como un
`.exe` armado con PyInstaller.

No es una app web: no hay cliente ni servidor, es **un solo proceso**. Lo que sí hay es una
separación entre la capa visual (`interface/`) y el motor (`core/`, `validation/`,
`providers/`).

## Los cuatro caminos de entrada

`run.py` es el único arranque. Según cómo se lo llame:

```
python run.py                     -> interface/main_interface.py     abre la ventana
python run.py --run-country X     -> forms/_runner_common.get_runner  un país, sin ventana
python run.py --autonomous        -> autonomous_runner.main           el scheduler en background
python run.py --run-lambdatest X  -> providers/lambdatest_*           dispositivos en la nube
```

Los cuatro terminan en el mismo motor. La diferencia es quién dispara y si hay alguien
mirando.

## El recorrido de una corrida de leads

```
   run.py
     |
     v
   main_interface.iniciar_interfaz()        el usuario elige país, navegador y aprieta Iniciar
     |
     v
   GenericCountryBase(pais, browser=...)    arma la config del país desde country_configs
     |                                       (hereda de BaseFormFiller)
     v
   BaseFormFiller.run()                      el corazon: una corrida completa
     |
     +--> initialize_browser()               crea el navegador via browser_factory
     +--> process_landing_page()             abre la landing, ubica el iframe del form
     +--> fill_form_fields_auto_step()       completa el formulario paso a paso
     +--> submit_and_verify_form()           envia y verifica que haya thank-you page
     +--> write_tracked_fields_to_sheet()    deja constancia en el Excel de resultados
     |
     v
   resultados/                               Excel + capturas + mail final
```

## Las piezas

### `core/` — el motor

`BaseFormFiller` es la clase central: abre el navegador, detecta el formulario, lo completa
con los datos del Excel y lo envía. Tenía 156 métodos en un solo archivo de 7.000 líneas;
hoy está partida en mixins bajo `core/form_filler/`, uno por responsabilidad:

| Mixin | Qué resuelve |
|---|---|
| `navegacion.py` | scroll, iframes y popups de cookies |
| `casillas.py` | checkboxes y radio buttons |
| `desplegables.py` | `<select>`, incluidas las dependencias padre→hijo |
| `ids_dinamicos.py` | campos que la herramienta descubre sola y persiste |
| `por_mercado.py` | documentos de Brasil, Libro de Reclamaciones de Perú |
| `aem.py` | formularios 2.0 de AEM, que tienen otra estructura |

Son mixins, no composición: comparten `self`. Desde afuera se sigue importando
`from osocio.core.base_form_filler import BaseFormFiller` y nada cambió.

`GenericCountryBase` hereda de `BaseFormFiller` y le arma la configuración de cada país
desde `country_configs.py`. Antes había 9 archivos `Formulario_<País>_Main.py`; se
eliminaron y hoy los resuelve `forms/_runner_common.get_runner()`.

### `interface/` — la capa visual

Una pestaña por archivo, salvo `main_interface.py`, que tiene la ventana principal y el
scheduler. Ese archivo son 4.900 líneas en **una sola función**: está mapeado aparte en
[`mapa_interfaz.md`](mapa_interfaz.md).

### `providers/` — ejecución en la nube

`lambdatest_mac/` y `lambdatest_android/` corren los mismos formularios sobre dispositivos
reales de LambdaTest en vez de un navegador local. Android reutiliza casi toda la lógica de
Mac; lo único propio es el driver y la carpeta de resultados.

### `validation/` — verificación sin enviar

La pestaña Validación de Campos y el Comparador de Dealers abren el navegador para chequear
reglas o listados de concesionarios **sin llegar a enviar un lead real**.

### `utils/` — lo compartido

`paths.py` es el más importante de todos y merece su propia sección.

## `paths.py`: la única fuente de rutas

**Todas** las rutas del proyecto salen de `osocio/utils/paths.py`. Ningún otro módulo
calcula la raíz del proyecto ni arma carpetas de salida a mano, y hay un test que lo hace
cumplir (`test_rutas.py`).

```python
BASE_DIR                  la raiz (o la carpeta del .exe, si esta empaquetado)
DATA_DIR                  data/         Excels de entrada
JSON_DIR                  json/         configuracion persistente
RESULTS_DIR               resultados/
  RESULTS_LT_MAC_DIR        resultados/lambdatest_mac/
  RESULTS_LT_ANDROID_DIR    resultados/lambdatest_android/
  RESULTS_DEALERS_DIR       resultados/dealers/
  RESULTS_MASIVA_DIR        resultados/resultado_urlsinsertas/
TEMPORALES_DIR            temporales/   logs y respaldos antes de pisar Excels
DRIVERS_DIR               drivers/      los baja y actualiza la app sola
```

Es la única que sabe distinguir entre correr como script y correr empaquetada con
PyInstaller. Esa es exactamente la razón por la que tiene que ser una sola: cuando esa
lógica estaba copiada en 16 archivos, un movimiento de carpetas las dejó a todas mal y la
app escribió en el lugar equivocado durante una semana sin dar un solo error.

## `json/`: la memoria de la app

Sin esta carpeta la app abre pero **no sabe llenar ningún formulario**. Dos grupos:

- **Conocimiento de los formularios** — `field_validation_rules_<pais>.json`,
  `fixed_field_mappings.json`, `ids_dinamicos.json`. Viaja dentro del portable.
- **Estado local** — programación del scheduler, config de email, preferencias. No se
  versiona.

`ids_dinamicos.json` guarda además las **dependencias entre campos** (region → city →
dealer) y los valores que el usuario corrigió a mano.

## Cómo se prueba todo esto

El motor recibe con qué navegar en vez de fabricarlo:

```python
BaseFormFiller(config, browser_factory=mi_doble)
```

Por defecto es `BrowserManager.create_browser`, así que en producción no cambia nada. En
los tests se le pasa un doble y el motor corre entero sin abrir un navegador. Esa costura
es lo que hace testeable un motor de 4.700 líneas.
