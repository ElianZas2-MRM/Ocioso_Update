# Cómo trabajar en este repo

## Lo básico

`main` está protegida: no se commitea directo. Todo entra por rama y pull request.

```bash
git checkout main
git pull
git checkout -b fix/lo-que-arreglas
# ... trabajás ...
git push -u origin fix/lo-que-arreglas
```

Antes de pedir el PR, que esto esté en verde:

```bash
pytest                              # los 416
python -m pyflakes osocio run.py tests   # sin nombres indefinidos
python run.py                       # que la app abra
```

## Nombres de rama

| Prefijo | Para qué |
|---|---|
| `fix/` | arreglar algo roto |
| `feat/` | funcionalidad nueva |
| `refactor/` | mover o reorganizar sin cambiar comportamiento |
| `test/` | solo tests |
| `chore/` | mantenimiento, documentación, build |

## Si apilás ramas, mergealas en orden

Esto nos costó un rato, así que queda escrito.

Cuando una rama se apoya sobre otra que todavía no se mergeó, **hay que mergear la base
primero**. GitHub hace *squash merge*: aplasta todos los commits de la rama en uno nuevo con
un SHA distinto. Si mergeás la de arriba antes que la de abajo, git ve la historia duplicada
y te tira conflictos en archivos que nadie tocó dos veces.

Nos pasó con dos PRs invertidos y quedó una carpeta `lambdatest_mac/` huérfana en la raíz,
creada por el PR que entró segundo. Nadie lo notó hasta días después.

Si ya te pasó, el arreglo no es resolver el conflicto a mano:

```bash
git rebase --onto origin/main <la-rama-base> <tu-rama>
```

Eso se queda solo con tus commits propios, apoyados sobre el `main` nuevo.

## Dos reglas que los tests hacen cumplir

No son estilo: son decisiones de diseño que ya se rompieron una vez y ahora están vigiladas.

### Las rutas salen de `paths.py`

Ningún módulo puede calcular la raíz del proyecto por su cuenta. Nada de:

```python
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # NO
```

Va así:

```python
from osocio.utils.paths import BASE_DIR, JSON_DIR, RESULTS_DIR           # SI
```

`test_solo_paths_py_puede_calcular_la_raiz` lo verifica recorriendo el paquete por AST.

**Por qué importa:** cuando el código se movió a `osocio/`, había 16 copias de esa cuenta y
todas quedaron cortas en un nivel. La app siguió abriendo sin quejarse y escribió en la
carpeta equivocada durante una semana. Las dependencias entre campos dejaron de aplicarse y
el scheduler guardaba donde nadie leía. Ningún test lo detectó porque todos parcheaban las
rutas para aislarse.

Si agregás una carpeta nueva, va en `paths.py`.

### Los resultados van adentro de `resultados/`

Nada de crear carpetas hermanas en la raíz. Cada tipo de ejecución tiene su subcarpeta y
sale de `paths.py`. `test_ningun_modulo_arma_carpetas_de_resultados_a_mano` lo vigila.

## Sobre los tests

Hay dos clases y hacen falta las dos:

- **Los que aíslan** usan `tmp_path` y `monkeypatch` para probar lógica sin tocar datos
  reales. Son la mayoría.
- **Los que NO aíslan** (`test_rutas.py`, `test_configuracion.py`) verifican la resolución
  real, sin parchear nada.

El bug de rutas se coló porque **solo existía el primer tipo**: al parchear `JSON_DIR`,
ningún test ejercitaba la resolución verdadera. Si escribís un test que parchea una ruta,
preguntate si además hace falta uno que la verifique de verdad.

## Cosas que no van al repo

`resultados/`, `temporales/`, `drivers/`, `venv/`, `build/`, `dist/` y los Excel de datos
están en `.gitignore`.

Prestá atención a `dist/` en particular: **este repo pesa 388 MB por eso**. En su momento se
commitearon el `.exe`, un `.zip` de 91 MB, un `.rar` de 83 MB y cuatro copias de los drivers.
Siguen en el historial aunque ya no estén en el árbol. Antes de commitear, mirá `git status`.

## Al agregar código nuevo

- **Un módulo nuevo** entra solo al `.exe`: el `.spec` usa `collect_submodules('osocio')`.
  La excepción son los que empiezan con guion bajo, que hay que declarar a mano.
- **Después de `build.bat`, abrí el `.exe`.** PyInstaller compila bien y revienta al abrir
  cuando falta un `hiddenimport`. Que compile no alcanza.
