# Cómo colaborar

La rama `main` está **protegida**: nadie —ni el dueño— puede hacer `git push` directo a
`main`. Todo cambio entra por **Pull Request**. Tampoco se puede hacer `force-push` ni
borrar `main`.

## Quién puede tocar el repo

| Quién | Qué puede hacer |
|---|---|
| Colaboradores con acceso de escritura (los invita el dueño, uno por uno) | Crear ramas, pushear a esas ramas, abrir PRs y mergearlos |
| Cualquier otra persona (repo público) | Forkear y abrir un PR **desde su fork**. No puede pushear ni mergear nada; el PR no toca `main` hasta que un colaborador le da merge |

Ser colaborador no es automático: el dueño manda la invitación por usuario de GitHub y la
persona la acepta. Nadie se agrega solo.

## Flujo para un colaborador (tenés acceso de escritura)

```bash
# 1. Partí siempre de main actualizada
git checkout main
git pull

# 2. Rama nueva con nombre descriptivo: tipo/descripcion-corta
git checkout -b feat/nombre-del-cambio      # o fix/... , chore/... , docs/...

# 3. Trabajás y commiteás (ver convención abajo)
git add -A
git commit -m "feat: descripción corta en imperativo"

# 4. Subís la rama
git push -u origin feat/nombre-del-cambio

# 5. Abrís el PR contra main
gh pr create --base main --fill        # o desde la web de GitHub

# 6. Mergeás el PR (no hace falta aprobación de terceros)
gh pr merge --squash --delete-branch   # o el botón "Merge" en la web

# 7. Volvés a main y actualizás
git checkout main
git pull
```

## Flujo desde afuera (sin acceso de escritura)

1. Fork del repo (botón *Fork* en GitHub).
2. Cloná tu fork, hacé una rama, commiteá y pusheá a **tu** fork.
3. Abrí un PR desde tu fork hacia `main` de este repo.
4. Un colaborador lo revisa y, si está ok, le da merge.

## Convención de commits

- **Conventional commits**, en español, en imperativo:
  `feat:` (funcionalidad nueva), `fix:` (bug), `chore:` (mantenimiento/config),
  `docs:` (documentación), `refactor:`, `test:`.
- Un commit = un cambio con sentido propio. Si el PR mezcla cosas, partilo en varios commits.
- **Sin** `Co-Authored-By` ni firmas de herramientas de IA.

Ejemplos reales del repo:

```
feat: autovalores para campos detectados y marcado selectivo de checkboxes
fix: dropdowns no ignoran el valor cargado en el Excel
docs: actualiza README y capturas con reintento de fallidos y drivers
```

## Antes de abrir el PR

```powershell
.\venv\Scripts\activate
pip install -r requirements-dev.txt   # solo la primera vez (trae pytest)
python -m pytest                      # tiene que dar todo verde
python -m pyflakes osocio run.py tests   # sin nombres indefinidos
python run.py                         # que la app abra
```

Si tocaste lógica de llenado de formularios, además probá una corrida real chica
(`python run.py`, un país, pocas filas del Excel de datos) antes de mergear.

---

# Lo que aprendimos a los golpes

Lo de acá abajo no es estilo: son cosas que ya se rompieron una vez y ahora están
vigiladas por tests.

## Si apilás ramas, mergealas en orden

Cuando una rama se apoya sobre otra que todavía no se mergeó, **hay que mergear la base
primero**. GitHub hace *squash merge*: aplasta todos los commits de la rama en uno nuevo
con un SHA distinto. Si mergeás la de arriba antes que la de abajo, git ve la historia
duplicada y te tira conflictos en archivos que nadie tocó dos veces.

Nos pasó con dos PRs invertidos y quedó una carpeta `lambdatest_mac/` huérfana en la raíz,
creada por el PR que entró segundo. Nadie lo notó hasta días después.

Si ya te pasó, el arreglo no es resolver el conflicto a mano:

```bash
git rebase --onto origin/main <la-rama-base> <tu-rama>
```

Eso se queda solo con tus commits propios, apoyados sobre el `main` nuevo.

## Las rutas salen de `paths.py`

Ningún módulo puede calcular la raíz del proyecto por su cuenta. Nada de:

```python
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # NO
```

Va así:

```python
from osocio.paths import BASE_DIR, JSON_DIR, RESULTS_DIR                 # SI
```

`test_solo_paths_py_puede_calcular_la_raiz` lo verifica recorriendo el paquete por AST.

**Por qué importa:** cuando el código se movió a `osocio/`, había 16 copias de esa cuenta y
todas quedaron cortas en un nivel. La app siguió abriendo sin quejarse y escribió en la
carpeta equivocada durante una semana. Las dependencias entre campos dejaron de aplicarse y
el scheduler guardaba donde nadie leía. Ningún test lo detectó porque todos parcheaban las
rutas para aislarse.

Si agregás una carpeta nueva, va en `paths.py`.

## Los resultados van adentro de `resultados/`

Nada de crear carpetas hermanas en la raíz. Cada tipo de ejecución tiene su subcarpeta y
sale de `paths.py`. `test_ningun_modulo_arma_carpetas_de_resultados_a_mano` lo vigila.

## Los tests necesitan las dos clases

- **Los que aíslan** usan `tmp_path` y `monkeypatch` para probar lógica sin tocar datos
  reales. Son la mayoría.
- **Los que NO aíslan** (`test_rutas.py`, `test_configuracion.py`) verifican la resolución
  real, sin parchear nada.

El bug de rutas se coló porque **solo existía el primer tipo**: al parchear `JSON_DIR`,
ningún test ejercitaba la resolución verdadera. Si escribís un test que parchea una ruta,
preguntate si además hace falta uno que la verifique de verdad.

## Que un módulo importe no significa que ande

Python no resuelve los nombres hasta que se llama la función. Un módulo puede importar
perfecto y reventar con `NameError` recién en producción. Pasó dos veces:

- `import a.b.c` liga el nombre `a`, no `c`. El import pasa y el atributo falta al usarlo.
- Al mover funciones entre archivos, las que quedaron usando nombres del archivo original
  importan bien y fallan al llamarlas.

Por eso `test_humo_general.py` no solo importa cada módulo: también verifica que los
símbolos públicos **existan**.

## Cosas que no van al repo

`resultados/`, `temporales/`, `drivers/`, `venv/`, `build/`, `dist/` y los Excel de datos
están en `.gitignore`.

Prestá atención a `dist/` en particular. El repo llegó a pesar **388 MB** porque en su
momento se commitearon el `.exe`, un `.zip` de 91 MB, un `.rar` de 83 MB y cuatro copias de
los drivers. En septiembre de 2026 se reescribió el historial para purgarlos y quedó en
**13 MB**, pero eso obligó a que todos re-clonaran: no hay forma de deshacerlo sin volver a
hacer lo mismo. Antes de commitear, mirá `git status`.

> Una advertencia que quedó de ahí: reescribir el historial **no saca un secreto de
> GitHub**. El repo tiene refs `refs/pull/*`, una por PR, que apuntan a los commits
> originales; las gestiona GitHub y no se pueden pushear ni borrar. Si se filtra una
> credencial, lo único que la vuelve inservible es **rotarla**.

## Al agregar código nuevo

- **Un módulo nuevo** entra solo al `.exe`: el `.spec` usa `collect_submodules('osocio')`.
  La excepción son los que empiezan con guion bajo, que hay que declarar a mano.
- **Después de `build.bat`, abrí el `.exe`.** PyInstaller compila bien y revienta al abrir
  cuando falta un `hiddenimport`. Que compile no alcanza.
