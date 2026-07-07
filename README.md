# Osocio — Form Automation

App de escritorio (Windows, Python + Tkinter + Selenium) para automatizar el llenado y envío de formularios de leads en varios países, generar los Excels de datos de prueba, validar reglas de campos, programar ejecuciones recurrentes, y chequear que los concesionarios (dealers) de un formulario coincidan con un Excel de referencia.

> El historial de cambios detallado de versiones anteriores quedó guardado en `README_HISTORIAL_ANTERIOR.md` (no se perdió, solo se sacó de este archivo para dejar una guía limpia).

## Instalación

1. Python 3.10+ (o el que ya tengas en `venv/`).
2. Activar el entorno virtual:
   ```powershell
   .\venv\Scripts\activate
   ```
3. Instalar dependencias:
   ```powershell
   pip install -r requirements.txt
   ```
4. Colocar los drivers de navegador (`chromedriver.exe`, `geckodriver.exe`, `msedgedriver.exe`) dentro de `drivers/` — la app no los descarga automáticamente, solo usa los locales.

## Cómo correr la app

```powershell
python run.py
```

O usá el ejecutable ya compilado: `dist/OsocioFormAutomation.exe` (o la carpeta portable, ver [Build portable](#build-portable)).

Al abrir, vas a ver la barra superior con el logo, el campo de **Email destinatario** y el checkbox **Enviar mail** (configuración compartida por todas las pestañas), y debajo las 5 pestañas de la app.

---

## Pestaña: Envío de Leads

Es la pestaña principal: rellena y envía los formularios reales usando los Excels de datos generados.

![Envío de Leads](Asset/screenshots/01_envio_de_leads.png)

**Paso a paso:**

1. **Email destinatario** (arriba a la derecha): si querés recibir un resumen por mail, escribí el destinatario y tildá **Enviar mail**. Se habilitan "Adjuntar resultados", "Adjuntar screenshots" y el modo de envío (**1 por país** o **Consolidado**).
2. **⚙ Configurar**: abre la configuración avanzada — elegís "Un Excel por dispositivo" o "Un Excel compartido para todos los dispositivos", y tildás "Ver navegador mientras corre" / "Minimizar a la bandeja al cerrar" si los necesitás.
3. **MERCADOS** y **EXCELS POR MERCADO**: Consecutivo (uno detrás del otro) o Paralelo (todos a la vez).
4. **DISPOSITIVOS / NAVEGADORES**: selección múltiple — Chrome, Firefox, Edge, Mac LT, Android LT. Al elegir Mac LT o Android LT aparece el panel para cargar las credenciales de LambdaTest (usuario + access key).
5. **PAÍSES A EJECUTAR**: hacé click en las tarjetas de los mercados que querés correr (AR/BO/BR/CL/CO/EC/PY/PE/UY). El contador te dice cuántos elegiste.
6. **EJECUTAR ENVÍO**: se habilita apenas elegís al menos un país. Abre un modal que bloquea la ventana mientras corre, mostrando el progreso mercado por mercado.
7. **Ver Resultados**: abre la carpeta `resultados/` con el Excel de resultados y las capturas de pantalla de esa corrida.

> Antes de ejecutar necesitás tener generado el Excel de datos correspondiente — si falta, andá primero a la pestaña **Generar Excels con Datos**.

---

## Pestaña: Programación de Tests

Programa la ejecución automática y recurrente de "Envío de Leads" (por ejemplo, todos los martes a las 03:00) sin que tengas que apretar nada.

![Programación de Tests](Asset/screenshots/04_programacion_tests.png)

**Paso a paso:**

1. **MERCADOS** y **DISPOSITIVOS / NAVEGADORES**: mismos selectores que en "Envío de Leads".
2. **⚙ Configurar automatización**: abre el calendario semanal — elegís los días (Lun a Dom) y, dentro de cada día, los horarios en franjas de 15 minutos (o uno personalizado). También elegís los países a testear ahí.
3. Al guardar, la app valida que ya existan los Excel necesarios en `data/` para cada combinación país + dispositivo elegida (si falta alguno, avisa "Archivos Excel Faltantes").
4. **Programar test automático**: activa la programación. El botón cambia a **Iniciar ahora** (para disparar ya, sin esperar el horario) + **Desactivar**.
5. **La app tiene que estar abierta** para que el monitor (revisa cada 60 segundos) detecte el horario y dispare la ejecución — si estaba cerrada, al reabrirla corre el horario pendiente del día.

---

## Pestaña: Validación de Campos

Chequea que las reglas de validación (regex, largo, campo obligatorio, etc.) de cada campo del formulario real coincidan con lo esperado — sin llegar a enviar un lead real.

![Validación de Campos](Asset/screenshots/03_validacion_campos.png)

**Paso a paso:**

1. **Tabla de URLs**: se carga desde un Excel (columnas País / URL / Formulario). Usá **Abrir Excel** para editarlo por fuera, o **Actualizar** para recargarlo después de un cambio.
2. **Ver navegador**: tildalo si querés ver el browser mientras corre la validación.
3. **▶ Configuración de ID**: desplegá esta sección para mapear cada campo del form — el id del elemento HTML, su descripción, si es un dropdown, si es numérico, el regex completo y por carácter, y reglas rápidas (letras minúsculas, email, obligatorio...). También podés definir "Mensaje de error" esperado y "Dependencia" (por ejemplo, que Ciudad dependa de Región). La tabla de abajo lista las reglas ya cargadas — click para editar, o **Eliminar regla**.
4. **Ejecutar validación**: corre la validación real contra el/los formularios configurados.
5. **Resultados**: abre la carpeta con el detalle de la validación.

---

## Pestaña: Generar Excels con Datos

Genera el Excel de datos de prueba (nombre, documento, teléfono, email, modelo, ciudad, dealer, etc.) que después usan "Envío de Leads" y "Programación de Tests".

![Generar Excels con Datos](Asset/screenshots/02_generar_excels.png)

**Paso a paso:**

1. **MERCADO A GENERAR**: elegís un país a la vez (se genera un Excel por mercado).
2. **DISPOSITIVOS PARA EL EXCEL**: selección múltiple (Chrome, Firefox, Edge, Mac LT, Android LT) — se genera un Excel por dispositivo elegido. Tildá **"Es formulario T3 2.0"** si el form es la versión nueva Adobe AEM (el archivo se genera con sufijo `_T3.xlsx`).
3. **URLS A PROCESAR**: elegí el formato con las pills — **"URL Landing + URL Form"** (`url landing • url form • ...`) o **"Solo URL Form"** — y pegá las URLs en el cuadro de texto.
4. Botones de la barra inferior:
   - **GENERAR EXCELS**: crea los archivos con datos aleatorios nuevos.
   - **REGENERAR DATOS**: recrea los datos manteniendo las URLs ya cargadas.
   - **Borrar URLs**: limpia el cuadro de texto.
5. Los archivos quedan en `data/`, con el patrón `Lead_information_Formulario_<País>_<Dispositivo>.xlsx` (o `_T3.xlsx`).

---

## Pestaña: Comparador Dealers

Chequea que los concesionarios (dealers) de una marca estén correctamente cargados en un formulario real, comparando contra un Excel de dealers esperados — reemplaza el bookmarklet manual que antes se pegaba en la consola del navegador.

![Comparador Dealers](Asset/screenshots/05_comparador_dealers.png)

**Paso a paso:**

1. **MERCADO A CHEQUEAR**: tarjeta del país. Cada país guarda su propia configuración (columnas, URLs, etc.) automáticamente al cambiar de mercado.
2. **URL DEL FORMULARIO A CHEQUEAR**: pegá una o varias URLs en el cuadro de texto:
   - Modo **"URL Landing + URL Form"**: de a 2 líneas por form (`url landing` / `url form`).
   - Modo **"Solo URL Form"**: una URL de form por línea.
   - Si las URLs pegadas parecen ser de un país distinto al mercado seleccionado, aparece un aviso (como se ve en la captura de arriba).
   - Elegí el navegador (Chrome/Firefox/Edge) y tildá **"Ver navegador mientras corre"** si querés verlo (apagado por defecto).
3. **EXCEL DE DEALERS A CHEQUEAR**: botón **"Seleccionar Excel"** (cualquier .xlsx/.xls), fila de encabezado configurable (el Excel real no siempre arranca en la fila 1), y el toggle **"Este Excel: Tiene columna de filtro"** / **"No tiene filtro (usar todas las filas)"**. Con filtro, definís la columna, el valor a buscar y la condición: **Incluir**, **Excluir**, o **Buscar extras** (esta última dispara además la búsqueda de dealers EXTRA y DUPLICADOS).
4. **Columnas del Excel**: pills **region / city / dealer** (son los ids reales del HTML del `<select>` — no cambian de país a país aunque el label visible sí, ej. en Argentina se ve "Provincia" pero el id sigue siendo `region`); `dealer` siempre es obligatorio. Definís qué columna del Excel corresponde a cada uno, más BAC (opcional). Podés tildar **"buscar dealers EXTRA y DUPLICADOS"** y agregar **columnas adicionales a comprobar en el form** (columna del Excel + id de cualquier otro campo del HTML).
5. **Modelos** (opcional, solo forms T1): si el form tiene selector de modelo, corré la comparación para todos los modelos detectados o para una lista específica.
6. **Modo de salida**: "Solo Excel" o "Excel + Capturas (ZIP)".
7. **Configuraciones guardadas**: guardá el mapeo completo de columnas con el nombre que quieras (ej. "GMUY Livianos") para reusarlo sin reconfigurar todo de nuevo.
8. **EJECUTAR**: valida todo antes de arrancar — si falta algo (Excel, URL de form, etc.) muestra un aviso corto sin llegar a abrir nada. Si está todo bien, abre un modal de progreso que bloquea la ventana principal (con botón **Detener**), y al terminar muestra el resumen PASS/FAIL/EXTRA/DUPLICADO.
9. Los reportes (Excel con colores por estado, y el ZIP de capturas si corresponde) quedan en la carpeta `Dealerscheck_resultados/`.

---

## Estructura de carpetas

- `core/`: lógica base de automatización y formularios por país, incluido `dealer_comparator_runner.py`.
- `forms/`: scripts de entrada para ejecutar los formularios de cada país.
- `interface/`: UI de administración (una pestaña por archivo) y utilidades de soporte.
- `data/`: archivos Excel de entrada (datos de prueba por país/dispositivo).
- `drivers/`: drivers locales del navegador (`chromedriver.exe`, `geckodriver.exe`, `msedgedriver.exe`).
- `resultados/`: resultados y capturas de "Envío de Leads" / LambdaTest.
- `Dealerscheck_resultados/`: reportes y capturas del Comparador Dealers.
- `json/`: configuración persistente (email, programación, reglas de validación, presets del Comparador Dealers).

## Drivers locales (sin descargas automáticas)

La app usa exclusivamente drivers locales desde `drivers/` (junto al proyecto, o junto al `.exe` al compilar). Si el navegador se actualiza y el driver queda desfasado, reemplazá el `.exe` correspondiente ahí adentro.

## Build portable

```powershell
.\build.bat
```

Genera:
- `dist/OsocioFormAutomation.exe`
- `dist/OsocioFormAutomation_portable/` (con `data/`, `drivers/`, `json/`, `resultados/`, `temporales/`, `Dealerscheck_resultados/`, `lambdatest_mac/` y `lambdatest_android/`)
- `dist/OsocioFormAutomation_portable.zip`

El portable arranca siempre limpio: sin schedule activo, sin configuración personal del Comparador Dealers, y sin `config_global.json` (evita llevarse el email o la access key de LambdaTest de la PC donde se compiló). Los drivers deben seguir distribuyéndose manualmente dentro de `drivers/`.

## Seguridad — credenciales

- `lambdatest_credentials.txt` y `json/config_global.json` (contiene el email y la access key de LambdaTest) están en `.gitignore` — nunca se suben al repositorio.
- El build portable tampoco los incluye (ver arriba).
- Si necesitás correr LambdaTest, creá `lambdatest_credentials.txt` en la raíz del proyecto con:
  ```
  username=TU_USUARIO
  access_key=TU_ACCESS_KEY
  ```
