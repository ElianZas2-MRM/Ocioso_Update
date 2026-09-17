# Guía de desarrollo

Para quien va a tocar el código. Si solo vas a **usar** la app, el manual está en
[`../README.md`](../README.md).

## Levantar el entorno

Hace falta Python 3.10 o superior (el build busca hasta 3.14) y Windows: la app usa
`pywin32` para el ícono de la bandeja y el Programador de tareas.

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements-dev.txt
```

> **Usá el venv de verdad, no el Python del sistema.** Los tests pasan igual sin
> `pywin32` instalado —los imports están protegidos— pero la app queda a medias sin
> avisar: `pywin32` es lo que usa para mandar los mails por Outlook y para el ícono de la
> bandeja. Si corrés desde el Python del sistema, el envío de emails falla y lo único que
> vas a ver es una línea en `temporales/runtime.log`:
>
> ```
> [ERROR] No module named 'win32com'
> ```
>
> El `.exe` no tiene este problema: `build.bat` instala `requirements.txt` en el venv que
> empaqueta.

### Por qué hay dos archivos de requirements

No es un descuido, es a propósito:

- **`requirements.txt`** — las 8 dependencias que la app necesita para correr. Es el único
  que `build.bat` instala en el venv que se empaqueta dentro del `.exe`.
- **`requirements-dev.txt`** — incluye al anterior con `-r requirements.txt` y le suma
  `pytest`.

Para desarrollar instalás **solo `requirements-dev.txt`** y ya tenés todo. Si se
fusionaran, `pytest` terminaría empaquetado dentro del `.exe` que se le entrega al
usuario final: peso que no usa nadie.

Las versiones usan rangos `~=`, que aceptan parches y minors pero no saltos de major
(que son los que rompen). `build.bat` recrea el venv desde cero en cada build: sin acotar,
dos builds del mismo código podían quedar con dependencias distintas.

## Correr la app

```bash
python run.py                          # abre la interfaz
python run.py --help                   # todas las opciones
python run.py --run-country Argentina  # un país puntual, sin abrir ventana
python run.py --autonomous --once      # ejecuta la programación guardada una vez
```

`run.py` es el **único** punto de entrada. Según cómo lo llames abre la interfaz, arranca
el programador autónomo o corre un país directo por consola.

## Correr los tests

```bash
pytest              # los 416
pytest -q           # resumido
pytest tests/test_rutas.py -v   # un archivo puntual
```

Tarda unos 20 segundos, a propósito: tiene que poder correrse seguido. Ningún test abre un
navegador ni sale a la red.

### Cómo está organizada la suite

| Archivo | Qué cubre |
|---|---|
| `test_rutas.py` | Que la raíz del proyecto se resuelva bien, y que **nadie la calcule por su cuenta** |
| `test_configuracion.py` | Que los JSON de `json/` parseen y tengan lo que el motor espera |
| `test_humo_general.py` | Que los 53 módulos importen y expongan sus símbolos públicos |
| `test_humo_funcionalidades.py` | Un smoke test por funcionalidad |
| `test_costura_navegador.py` | Que el motor pueda correr con un navegador falso |
| `test_estructura_form_filler.py` | Que partir `BaseFormFiller` en mixins no pierda métodos |
| `test_carpetas_de_resultados.py` | Que todo lo que la app produce caiga dentro de `resultados/` |
| el resto | Lógica de negocio: autovalores, regex del CRM, checkboxes, evidencia |

### Dos reglas que la suite hace cumplir

Hay dos tests que no prueban una función sino que **vigilan una decisión de diseño**. Los
dos existen por bugs reales que costaron caro:

1. **`test_solo_paths_py_puede_calcular_la_raiz`** — ningún módulo puede encadenar
   `os.path.dirname()` sobre `__file__` para llegar a la raíz del proyecto. Eso lo hace
   `osocio/paths.py` y nadie más. Cuando el código se movió a `osocio/`, había 16
   copias de esa cuenta y todas quedaron cortas en un nivel: la app escribió en la carpeta
   equivocada durante una semana sin quejarse.

2. **`test_ningun_modulo_arma_carpetas_de_resultados_a_mano`** — las carpetas de salida
   salen de `paths.py`, no de un `os.path.join` suelto. Antes había cuatro carpetas de
   resultados hermanas colgando de la raíz.

Si agregás una ruta nueva, va en `paths.py`.

## Armar el `.exe`

```bash
build.bat
```

Recrea el venv desde cero, instala `requirements.txt` y PyInstaller, y arma
`dist/OsocioFormAutomation_portable/` más su ZIP. Tarda varios minutos.

**El `.exe` compila bien y revienta al abrir** si falta un `hiddenimport`, así que después
de buildear hay que **abrirlo** — que compile no alcanza como verificación.

`FormAutomation.spec` usa `collect_submodules('osocio')` en vez de una lista a mano, así
que los módulos nuevos entran solos. La única excepción está anotada ahí: los módulos que
empiezan con guion bajo hay que declararlos aparte, porque `collect_submodules` los saltea.

## Dónde está cada cosa

```
run.py                     punto de entrada unico
osocio/
  core/                    el motor de llenado
    form_filler/           BaseFormFiller partida en mixins por responsabilidad
    browser/               acciones sobre el navegador
  interface/               la UI en Tkinter, una pestaña por archivo
  validation/              motor de la pestaña Validación de Campos
  providers/               LambdaTest Mac y Android
  utils/                   rutas, generación de datos, drivers, programación
  forms/                   runner compartido de los países
tests/                     503 tests
data/   json/   Asset/     datos, configuración y recursos del .exe
resultados/                todo lo que la app produce
docs/                      esta carpeta
```

Las carpetas `data/`, `resultados/`, `temporales/` y `drivers/` no se versionan: las crea
y llena la app sola.
