@echo off
REM Arma el .exe y el portable de Osocio Form Automation.
REM El historial de cambios de este script esta en docs/historial_build.md
REM (ocupaba 116 de las 280 lineas de aca y tapaba lo que el script hace).
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
REM Build onefile: PyInstaller deja UN solo dist\%APP_NAME%.exe con todo comprimido
REM adentro (sin carpeta _internal\ con las librerias sueltas).
copy /y "dist\%APP_NAME%.exe" "%PORTABLE_DIR%\%APP_NAME%.exe" >nul
if errorlevel 1 goto :error
if not exist "%PORTABLE_DIR%\%APP_NAME%.exe" goto :error

REM data/: NO viaja ningun Excel de datos. Los de leads los genera la propia app desde
REM "Generar Excels con Datos", el Excel matriz de Revision Masiva lo elige el usuario, y
REM los listados de dealers son datos reales de clientes. Lo unico que se lleva es el Excel
REM de URLs de Validacion de Campos, que la pestana espera encontrar ahi.
mkdir "%PORTABLE_DIR%\data"
if exist ".\data\Field_Validation_URLs.xlsx" (
    copy /y ".\data\Field_Validation_URLs.xlsx" "%PORTABLE_DIR%\data\Field_Validation_URLs.xlsx" >nul
)

REM drivers/ si viaja, pero ya NO se versionan en el repo: los baja la app sola. Antes de
REM empaquetar se refrescan para que el portable salga con los que corresponden a los
REM navegadores de esta PC. Si no hay internet se usa lo que haya en .\drivers y, en el peor
REM caso, el portable los descarga solo en el primer arranque.
echo Verificando drivers de navegador...
%PY% -c "import truststore; truststore.inject_into_ssl(); from osocio.utils.driver_updater import ensure_drivers_ready; ensure_drivers_ready()"

if exist ".\drivers" (
    robocopy ".\drivers" "%PORTABLE_DIR%\drivers" /E /XD ".tmp_update" /NFL /NDL /NJH /NJS /NC /NS >nul
    if errorlevel 8 goto :error
) else (
    mkdir "%PORTABLE_DIR%\drivers"
)

REM Las carpetas de salida se crean VACIAS. Antes se copiaban enteras y el portable
REM viajaba con los resultados, capturas y reportes de revision masiva de esta PC
REM (cientos de MB de datos de otra corrida). El portable arranca limpio.
for %%D in (resultados temporales) do (
    if not exist "%PORTABLE_DIR%\%%D" mkdir "%PORTABLE_DIR%\%%D"
)

REM Copiar json/ sin los archivos de estado del scheduler, la config personal del
REM Comparador Dealers, ni config_global.json (tiene el email y la access key de
REM LambdaTest en texto plano) — el portable arranca limpio, sin datos de otra PC
if exist ".\json" (
    robocopy ".\json" "%PORTABLE_DIR%\json" /E /NFL /NDL /NJH /NJS /NC /NS /XF programacion_test.json programacion_leads.json programacion_masivo.json scheduler_triggered.json dealer_comparator_settings.json config_global.json ejecutor_autonomo.log >nul
    if errorlevel 8 goto :error
) else (
    mkdir "%PORTABLE_DIR%\json"
)

REM Red de seguridad: el portable tiene que arrancar SIN ninguna programacion activa (ni
REM Envio de Leads Programados ni Revision Masiva). El /XF de arriba ya los excluye, pero
REM se borra por patron ademas, para que un archivo de programacion con nombre nuevo no se
REM cuele por estar fuera de esa lista.
if exist "%PORTABLE_DIR%\json\programacion_*.json" del /f /q "%PORTABLE_DIR%\json\programacion_*.json"
if exist "%PORTABLE_DIR%\json\scheduler_triggered.json" del /f /q "%PORTABLE_DIR%\json\scheduler_triggered.json"

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

REM Se borra el .exe crudo de PyInstaller (ya fue copiado adentro del portable) y la carpeta
REM dist\%APP_NAME%\ por si quedo de un build onedir anterior.
if exist "dist\%APP_NAME%" rmdir /s /q "dist\%APP_NAME%"
if exist "dist\%APP_NAME%.exe" del /f /q "dist\%APP_NAME%.exe"

REM ZIP del portable, para repartir un unico archivo comprimido.
if exist "dist\%APP_NAME%_portable.zip" del /f /q "dist\%APP_NAME%_portable.zip"
echo Comprimiendo el portable...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Compress-Archive -Path '%PORTABLE_DIR%' -DestinationPath 'dist\%APP_NAME%_portable.zip' -CompressionLevel Optimal -Force"
if not exist "dist\%APP_NAME%_portable.zip" (
    echo ADVERTENCIA: no se pudo armar el ZIP, queda solo la carpeta portable.
)

echo.
echo Build completado correctamente.
echo   Portable: %PORTABLE_DIR%\
echo   ZIP:      dist\%APP_NAME%_portable.zip
echo   (Abrir con: %PORTABLE_DIR%\Abrir_Osocio_Form_Automation.bat)
pause
exit /b 0

:error
echo.
echo Ocurrio un error durante la compilacion.
pause
exit /b 1
