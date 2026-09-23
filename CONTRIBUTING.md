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

## Lo que el usuario cargó en el Excel gana siempre

Ningún valor guardado, sembrado o sorteado puede pisar una celda con dato. El orden está
documentado en el README (*De dónde sale el valor de cada campo*) y es, en corto:

```
valor forzado por formulario  >  celda del Excel  >  IDs Dinámicos  >  al azar  >  vacío
```

Si agregás una fuente de valores nueva, va **después** de la celda del Excel.

**Por qué importa:** el llenado consultaba `json/ids_dinamicos.json` *antes* de mirar el
Excel, y cortaba ahí con un `return`. Un autovalor que la app se había sembrado sola le
ganaba al CPF cargado a mano, en cada corrida, durante meses. Lo peor es que el invariante
ya estaba escrito tres líneas más arriba, en el comentario del sembrador de autovalores:
*"El Excel siempre gana; nunca pisa lo que el usuario ya cargó a mano"*. El sembrador lo
respetaba al escribir; el llenado no lo respetaba al leer.

`test_prioridad_del_excel.py` lo vigila, y además verifica que el llenado consulte la regla
en vez de mirar los IDs dinámicos por su cuenta.

## Los ids del mapping no son los del formulario

El mapping de cada país usa los `id` clásicos; los formularios `gm_frontend` usan otros
(`cpf` → `document`, `cep` → `zip_code`, `firstname` → `name`…). Esa tabla vive en
`osocio/utils/field_id_aliases.py` y es la única fuente.

**Si agregás un alias, revisá quién más decide por el `id`.** Al resolverse el alias, el id
que viaja por el resto del llenado es el del DOM (`document`), no el del mapping (`cpf`).
Cualquier `if "cpf" in field_id` que haya por ahí deja de ser cierto. Pasó con la
normalización del documento y con la detección de campos enmascarados: agregar el alias
solo, sin tocar esos dos, arreglaba una cosa y rompía otra.

## Lo que hacen todos los navegadores, escrito una sola vez

Hay cuatro cosas en la app que abren un navegador: Envío de Leads, Comparar Dealers,
Validación de Campos y LambdaTest. Cuando las cuatro necesitan el mismo comportamiento del
navegador, **va en un solo lugar** y las cuatro lo llaman.

**Por qué importa:** el pre-scroll que carga el contenido diferido estaba escrito cuatro
veces. Las cuatro copias tenían el mismo bug (medían la altura de la página una sola vez,
antes de bajar, y una página con lazy-loading crece al bajar), más dos diferencias que
nadie decidió: dos de ellas no disparaban el evento de scroll, y dos tenían un corte de
seguridad que no podía ejecutarse nunca. Arreglar una sola no arreglaba nada.

Hoy vive en `osocio/utils/scroll_dinamico.py`. Las cuatro funciones conservan su nombre y
su firma — solo delegan — así que los lugares que las llaman no se enteraron.

Lo que sí es legítimo que difiera son los tiempos de espera, y por eso son parámetros: en
headless el navegador necesita más para procesar los eventos, y una sesión remota de
LambdaTest ya paga su latencia en cada paso.

`test_pre_scroll.py` corre la misma batería por los cuatro accesos. Si mañana aparece un
quinto, el test avisa.

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
están en `.gitignore`. Además, **por nombre y estén donde estén**: los binarios de driver
(`chromedriver*`, `geckodriver*`, `msedgedriver*`…), cualquier `.exe`, y los documentos de
Office (`.docx`, `.pptx`). Esos van a SharePoint; en `docs/` quedan los links.

Prestá atención a `dist/` en particular. El repo llegó a pesar **388 MB** porque en su
momento se commitearon el `.exe`, un `.zip` de 91 MB, un `.rar` de 83 MB y cuatro copias de
los drivers. En septiembre de 2026 se reescribió el historial para purgarlos y quedó en
**14 MB**, pero eso obligó a que todos re-clonaran: no hay forma de deshacerlo sin volver a
hacer lo mismo. Antes de commitear, mirá `git status`.

**Por qué esto además se testea.** `.gitignore` solo es la primera capa: se puede saltear
con `git add -f`, a veces sin querer, y hasta ahora solo cubría la *carpeta* `drivers/` —
un binario suelto en cualquier otro lado entraba igual, que es exactamente como llegaron
los 86 MB que hubo que purgar. `test_nada_pesado_versionado.py` mira lo que está realmente
versionado: ningún driver, ningún formato que git no pueda delta-comprimir, y nada de más
de 2 MB.

Un `.docx` o un `.zip` ya vienen comprimidos, así que git no puede guardar solo las
diferencias entre versiones: **cada versión pesa entera y para siempre**. Por eso un
archivo así es mucho más caro que su tamaño. Y una vez adentro no sale sin otro rewrite.

> Si alguna vez hace falta subir uno a propósito: `git add -f <archivo>` y sumalo a
> `PERMITIDOS` en ese test, con el motivo escrito.

> Una advertencia que quedó de ahí: reescribir el historial **no saca un secreto de
> GitHub**. El repo tiene refs `refs/pull/*`, una por PR, que apuntan a los commits
> originales; las gestiona GitHub y no se pueden pushear ni borrar. Si se filtra una
> credencial, lo único que la vuelve inservible es **rotarla**.

## Si clonaste antes de septiembre de 2026, no hagas `pull`

El rewrite le cambió el SHA a **todos** los commits. Un clon viejo y el repo de hoy son, para
git, dos historias que no se conocen. Si hacés `git pull`, git intenta fusionarlas y te
devuelve esto en casi todos los archivos:

```
CONFLICT (add/add): Merge conflict in run.py
CONFLICT (add/add): Merge conflict in requirements-dev.txt
CONFLICT (add/add): Merge conflict in osocio/utils/scheduling.py
...
```

**`add/add` es la firma del problema**: significa "este archivo lo crearon los dos lados por
separado, no tienen ancestro en común". No es un conflicto de contenido — no hay nada que
resolver a mano. Y como el merge deja los marcadores `<<<<<<< HEAD` adentro de los `.py`, la
app deja de arrancar con un `SyntaxError`.

Cómo salir, sin perder `venv/`, `drivers/` ni tus Excels (están ignorados, sobreviven a todo
esto):

```bash
# 1. Cancelar el merge a medias
git merge --abort

# 2. ¿Tenías trabajo propio sin subir? Miralo antes de seguir
git log --oneline origin/main..HEAD
git status

# 3. Si había algo, guardalo en una rama antes de tocar nada
git branch respaldo-de-mi-clon-viejo

# 4. Apuntar a la historia buena
git fetch origin
git checkout main
git reset --hard origin/main
```

El paso 4 **descarta los cambios sin commitear**, así que mirá `git status` en el paso 2.

Después de eso, borrá las ramas locales que sigan colgando de la historia vieja: si trabajás
sobre una, el problema vuelve igual.

> Si el trabajo que guardaste en el paso 3 lo querés traer, **no lo mergees**: sacá los
> archivos sueltos (`git checkout respaldo-de-mi-clon-viejo -- ruta/al/archivo.py`) o copialos
> a mano. Mergear vuelve a cruzar las dos historias.

## Al agregar código nuevo

- **Un módulo nuevo** entra solo al `.exe`: el `.spec` usa `collect_submodules('osocio')`.
  La excepción son los que empiezan con guion bajo, que hay que declarar a mano.
- **Después de `build.bat`, abrí el `.exe`.** PyInstaller compila bien y revienta al abrir
  cuando falta un `hiddenimport`. Que compile no alcanza.
