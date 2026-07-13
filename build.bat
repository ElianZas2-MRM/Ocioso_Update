@echo off
REM -- Changelog -----------------------------------------------------------------
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

REM Buscar en PATH ignorando WindowsApps
for /f "delims=" %%i in ('where python 2^>nul') do (
    echo %%i | findstr /i "WindowsApps" >nul
    if errorlevel 1 (
        set "PY=%%i"
        goto :create_venv
    )
)

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
copy /y "dist\%APP_NAME%.exe" "%PORTABLE_DIR%\%APP_NAME%.exe" >nul
if errorlevel 1 goto :error

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

echo Generando ZIP...
if exist "dist\%APP_NAME%_portable.zip" del /f /q "dist\%APP_NAME%_portable.zip"
powershell -NoProfile -Command "Compress-Archive -Path '%PORTABLE_DIR%' -DestinationPath 'dist\%APP_NAME%_portable.zip'"
if errorlevel 1 echo AVISO: No se pudo generar el ZIP (no es critico).

echo.
echo Build completado correctamente.
echo   EXE:      dist\%APP_NAME%.exe
echo   Portable: %PORTABLE_DIR%\
echo   ZIP:      dist\%APP_NAME%_portable.zip
pause
exit /b 0

:error
echo.
echo Ocurrio un error durante la compilacion.
pause
exit /b 1
