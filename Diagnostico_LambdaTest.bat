@echo off
REM Diagnostico de conectividad a LambdaTest.
REM Copiar este .bat junto con Diagnostico_LambdaTest.ps1 dentro de la carpeta portable
REM (al lado de OsocioFormAutomation.exe) y hacer doble click.
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Diagnostico_LambdaTest.ps1"
