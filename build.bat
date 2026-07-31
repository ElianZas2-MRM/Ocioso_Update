@echo off
REM -- Changelog -----------------------------------------------------------------
REM Jul 2026: Excel de resultados: columnas de RESULTADO primero y los datos de entrada al
REM           final (antes era al reves). Se reordena al cerrar la corrida, no antes: los
REM           runners leen los datos del lead de la misma hoja (URL en A, Formulario en B)
REM           y moverlas durante la corrida desalinearia esa lectura
REM           Si la landing no cargo (404 / sin respuesta / redirect), "Form coincide" y
REM           "Formulario Inserto" ya no dicen "no coincide con el esperado" — el form nunca
REM           se busco. Dicen "N/D — la landing no cargo (...)" con el estado real
REM Jul 2026: Excel: SIEMPRE se dice POR QUE fallo un lead. Nueva columna "Motivo" con la
REM           causa corta (form no inserto / no coincide / validacion del form + el campo /
REM           sin TY page / event_id / 404). Antes habia que leer el Resultado entero
REM           "Formulario Inserto" pasa a TRES estados en vez de dos:
REM             verde   "Form inserto"
REM             AMBAR   "Form inserto NO coincide con el esperado, se envio lead igualmente"
REM             rojo    "Form NO inserto (iframe sin src / sin form en la landing)"
REM           "Form coincide" acompana: FAIL en ambar si el lead viajo igual (hay dato en la
REM           base, entro por otro form), FAIL en rojo si no salio o no habia form
REM           Nuevas columnas "Estado URL landing" / "Estado URL form" (utils/url_status.py):
REM           status HTTP real de cada URL — 200 verde, redirect ambar, 404/503/sin respuesta
REM           rojo. Selenium no expone el status, asi que una landing caida se veia solo como
REM           "form no encontrado". Si la URL falla, se antepone al Motivo como causa raiz
REM           Aplica a browsers (escritorio) y a LambdaTest Mac/Android
REM Jul 2026: Resultados = lo que REALMENTE se envio: antes del click en Enviar se relee el
REM           DOM (1 solo execute_script) y se pisa el tracking con el valor efectivo. El
REM           reintento (recarga la landing y rellena de cero) re-sorteaba modelo/ciudad/
REM           concesionario y el Excel quedaba con lo del PRIMER intento -> no coincidia con
REM           la base de datos. Los dropdowns random sin trackear quedan como Final::<campo>
REM           Nueva columna "Datos vs Excel": verde OK si los dropdowns quedaron como se
REM           pidio; ambar con el detalle (pedido 'X' -> quedo 'Y') si no, y ahi el Resultado
REM           tampoco queda verde. Solo dropdowns: en texto el flujo transforma a proposito
REM           (CPF/CNPJ regenerados, maxlength) y darian avisos falsos
REM           Fix Android: el.clear() no vaciaba el campo NI lanzaba excepcion -> el
REM           re-ingreso concatenaba ('ApellidoApellido', email duplicado) y el lead fallaba
REM           por validacion. Vaciado verificado en cascada; si no se puede, no escribe encima
REM           Fix modelo por ?model=: se leia driver.current_url, que dentro del iframe es la
REM           landing (sin el parametro) -> esas filas quedaban sin modelo registrado. Ahora
REM           usa la URL del form del Excel, con filtro de tokens que no son modelo
REM           Aplica a LambdaTest Android/Mac (lt_runner.py) y al runner de escritorio
REM Jul 2026: iframe GM: siempre priorizar src con gm_forms/gm_front/gm_admin (evita agarrar
REM           el iframe equivocado cuando la landing tiene varios) — browsers y LambdaTest
REM           CTA: mas selectores por texto ES/PT + barrido generico del <form> (fix "no
REM           encuentra el boton que si esta"); retry recargando SOLO el iframe (no la landing)
REM           Capturas: landing_inicial -> form_vacio -> form_errores -> form_completado -> TY
REM           -> landing_final; sin capturas de landing completa en el medio; multipaso por paso
REM           Foco: ventana real off-screen SIN robar foco via Win32 SW_SHOWNOACTIVATE (no headless)
REM           Modal: progreso por FORMULARIOS (no sesiones) + mercados en ejecucion + filas con
REM           error y motivo corto (form ausente / incorrecto / sin TYP / landing 404 / ...)
REM           Errores: cualquier fallo del form = FAIL en Excel/UI/email (TYP no vista, form
REM           incorrecto, landing 404, campo/dropdown sin completar). "FORMULARIO AUSENTE" si no
REM           hay form; "Formulario incorrecto (distinto al esperado)" si es otro
REM           Resultados: columna Modelo = modelo elegido en el dropdown o ?model= real del form
REM           LambdaTest: NUNCA capturas (evidencia = video); mismo iframe GM / CTA / modelo
REM Jul 2026: Nueva pestaña "Comparador Dealers": chequea region/ciudad/dealer/BAC/modelos
REM           contra un Excel de dealers (fila de encabezado y columnas configurables,
REM           múltiples forms por pasada, detección de duplicados y extras de forma jerárquica,
REM           desbloqueo de nivel dealer, columnas adicionales como píldoras-checkbox en vivo,
REM           ejecución en dos fases (comparación rápida primero + capturas ZIP opcionales después),
REM           corrección de screenshots en Chrome dentro de iframe, e icono de la app propio.
REM Jun 2026: Generar Excels: panel Brasil visible al seleccionar pais (pack before fix)
REM           Generar Excels: boton "Regenerar datos (URLs actuales)" sin ingresar URLs
REM           Generar Excels: radio buttons modo envio (redondelitos, sin indicatoron=0)
REM           Ejecucion: forms standalone (sin landing URL) ya no se saltean
REM           Browser: viewport desktop garantizado (1366x768) en Chrome/FF/Edge
REM           Browser: ventana movida fuera de pantalla (x=10000) en los 3 navegadores
REM           Build: lambdatest_android incluido en portable y dist
REM Jun 2026: LambdaTest: fix terms checkbox (React fiber onChange + shadow DOM cookies)
REM           LambdaTest: zoom 80% en URLs comprar-carro para visibilidad completa
REM           LambdaTest: scroll al fondo tras 2do campo en forms comprar-carro
REM           LambdaTest: cookie popup GM via shadowRoot (gb-legal-notification)
REM Jun 2026: Fix Firefox: screenshots faltantes en form_errores/completado/typage
REM           CPF/CNPJ/CEP: cero inicial recuperado si Excel lo omio (zfill)
REM           Chevrolet BR: click en #contact-by-form antes de entrar al form
REM           Error COM de Outlook muestra mensaje real (antes generico)
REM           Tracking por paso en Excel de resultados (columnas PasoN::campo)
REM           Seleccion de dropdowns ahora solo exacta (fuzzy matching eliminado)
REM           Generacion de documentos Brasil via API 4devs
REM           Retirado xlsxwriter (reemplazado por openpyxl)
REM           Soporte checkboxes custom con opacity:0 (fake-terms, stat-radio)
REM           Retry generico: cualquier fallo de envio recarga y reintenta (max 2)
REM           Sanitizacion 'test' a 'prueba' en campos de texto
REM           Email auto-generado: formato {nombre}{apellido}_{pais}{nn}@mrm.com
REM           Fix URLs form esperado/encontrado por fila en capturas
REM           pywin32 requerido para envio de email via Outlook (win32com)
REM           libro-reclamaciones (chevrolet.com.pe): IDs fijos cc_name/
REM           cc_telephone/cc_ci/cc_email + ciudad/dealer aleatorio + envio real
REM           Deteccion exacta de error transitorio de envio: BR (Desculpe...) y
REM           resto de mercados (Lo siento, ocurrio un inconveniente...)
REM ------------------------------------------------------------------------------
setlocal
set "APP_NAME=OsocioFormAutomation"
set "PORTABLE_DIR=dist\%APP_NAME%_portable"
cd /d "%~dp0"

REM Borrar venv viejo (puede tener rutas de otra PC hardcodeadas)
if exist ".\venv" (
    echo Eliminando venv anterior...
    rmdir /s /q ".\venv"
)

REM Buscar Python instalado, ignorar stub de WindowsApps
set "PY="

for %%V in (314 313 312 311 310 39) do (
    if exist "%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe" (
        set "PY=%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe"
        goto :create_venv
    )
    if exist "C:\Python%%V\python.exe" (
        set "PY=C:\Python%%V\python.exe"
        goto :create_venv
    )
    if exist "C:\Program Files\Python%%V\python.exe" (
        set "PY=C:\Program Files\Python%%V\python.exe"
        goto :create_venv
    )
)

REM Buscar en PATH: probar cada "python" encontrado y quedarse con el primero que
REM realmente funcione (--version imprime "Python 3..."). Un Python instalado desde
REM la Microsoft Store tambien vive bajo WindowsApps y es valido, a diferencia del
REM stub vacio que solo abre la Store (ese no imprime version real).
for /f "delims=" %%i in ('where python 2^>nul') do (
    if not defined PY (
        for /f "delims=" %%v in ('"%%i" --version 2^>^&1') do (
            echo %%v | findstr /b /c:"Python 3" >nul
            if not errorlevel 1 set "PY=%%i"
        )
    )
)
if defined PY goto :create_venv

echo ERROR: No se encontro Python instalado correctamente.
echo Instala Python desde https://www.python.org/downloads/
goto :error

:create_venv
echo Python encontrado: %PY%
echo Creando entorno virtual limpio...
%PY% -m venv venv
if errorlevel 1 goto :error
set "PY=.\venv\Scripts\python.exe"

echo Instalando dependencias...
%PY% -m pip install --upgrade pip -q
if errorlevel 1 goto :error
%PY% -m pip install pyinstaller -q
if errorlevel 1 goto :error
%PY% -m pip install -r requirements.txt -q
if errorlevel 1 goto :error

echo Compilando con PyInstaller...
%PY% -m PyInstaller --clean FormAutomation.spec
if errorlevel 1 goto :error

echo Armando carpeta portable...
if exist "%PORTABLE_DIR%" rmdir /s /q "%PORTABLE_DIR%"
mkdir "%PORTABLE_DIR%"
REM Build onedir: PyInstaller deja una CARPETA dist\%APP_NAME%\ (el .exe + _internal\)
REM en vez de un unico .exe, asi que se copia entera a la raiz del portable.
robocopy "dist\%APP_NAME%" "%PORTABLE_DIR%" /E /NFL /NDL /NJH /NJS /NC /NS >nul
if errorlevel 8 goto :error
if not exist "%PORTABLE_DIR%\%APP_NAME%.exe" goto :error

REM data/ se copia aparte: hay que dejar afuera los Excel reales de clientes
REM (rankings/listados de dealers), que no deben viajar en el portable ni en el ZIP
if exist ".\data" (
    robocopy ".\data" "%PORTABLE_DIR%\data" /E /NFL /NDL /NJH /NJS /NC /NS /XF "*Ranking Dealers*.xls*" >nul
    if errorlevel 8 goto :error
) else (
    mkdir "%PORTABLE_DIR%\data"
)

for %%D in (drivers resultados temporales Dealerscheck_resultados resultados_lambdatestmac resultados_lambdatest_android) do (
    if exist ".\%%D" (
        robocopy ".\%%D" "%PORTABLE_DIR%\%%D" /E /NFL /NDL /NJH /NJS /NC /NS >nul
        if errorlevel 8 goto :error
    ) else (
        mkdir "%PORTABLE_DIR%\%%D"
    )
)

REM Copiar json/ sin los archivos de estado del scheduler, la config personal del
REM Comparador Dealers, ni config_global.json (tiene el email y la access key de
REM LambdaTest en texto plano) — el portable arranca limpio, sin datos de otra PC
if exist ".\json" (
    robocopy ".\json" "%PORTABLE_DIR%\json" /E /NFL /NDL /NJH /NJS /NC /NS /XF programacion_test.json scheduler_triggered.json dealer_comparator_settings.json config_global.json ejecutor_autonomo.log >nul
    if errorlevel 8 goto :error
) else (
    mkdir "%PORTABLE_DIR%\json"
)

(
    echo @echo off
    echo cd /d "%%~dp0"
    echo start "" "%APP_NAME%.exe"
) > "%PORTABLE_DIR%\Abrir_Osocio_Form_Automation.bat"

REM Crear plantilla de credenciales si no existe en el portable
if not exist "%PORTABLE_DIR%\lambdatest_credentials.txt" (
    echo # LambdaTest credentials - completar con tus datos> "%PORTABLE_DIR%\lambdatest_credentials.txt"
    echo username=TU_USUARIO>> "%PORTABLE_DIR%\lambdatest_credentials.txt"
    echo access_key=TU_ACCESS_KEY>> "%PORTABLE_DIR%\lambdatest_credentials.txt"
    echo Plantilla lambdatest_credentials.txt creada en portable.
)

REM Solo se entrega la carpeta portable: se borra la carpeta cruda de PyInstaller
REM dist\%APP_NAME%\ (ya fue copiada adentro del portable) y no se arma ZIP.
if exist "dist\%APP_NAME%" rmdir /s /q "dist\%APP_NAME%"
if exist "dist\%APP_NAME%.exe" del /f /q "dist\%APP_NAME%.exe"
if exist "dist\%APP_NAME%_portable.zip" del /f /q "dist\%APP_NAME%_portable.zip"

echo.
echo Build completado correctamente.
echo   Portable: %PORTABLE_DIR%\
echo   (Abrir con: %PORTABLE_DIR%\Abrir_Osocio_Form_Automation.bat)
pause
exit /b 0

:error
echo.
echo Ocurrio un error durante la compilacion.
pause
exit /b 1
