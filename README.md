# Form Automation Project

Automatización de formularios para varios países con Selenium, generación de resultados y UI de administración.

## Requisitos

- Python 3.10+ o compatible
- Activar el entorno virtual provisto con:
  ```powershell
  .\venv\Scripts\activate
  ```
- Instalar dependencias:
  ```powershell
  pip install -r requirements.txt
  ```

## Ejecución

1. Asegúrate de ejecutar desde la carpeta raíz del proyecto.
2. Inicia la interfaz con:
   ```powershell
   python run.py
   ```
3. Desde la interfaz puedes:
   - cargar o crear el Excel de datos
   - editar filas y guardar cambios
   - lanzar la ejecución de formularios por país
   - abrir resultados y capturas

## Estructura de carpetas

- `core/`: lógica base de automatización y formularios por país.
- `forms/`: scripts de entrada para ejecutar los formularios de cada país.
- `interface/`: UI de administración y utilidades de soporte.
- `data/`: archivos Excel de entrada.
- `drivers/`: drivers locales del navegador (`chromedriver.exe`, `geckodriver.exe`, `msedgedriver.exe`).
- `resultados/`: resultados generados y capturas de pantalla.
- `json/`: configuración persistente y programación.

## Drivers locales (sin descargas automáticas)

La app usa exclusivamente drivers locales desde la carpeta `drivers/` ubicada junto al proyecto (o junto al `.exe` al compilar con PyInstaller).

Estructura esperada:

```text
Form_Automation_Project/
|-- run.py
|-- drivers/
|   |-- chromedriver.exe
|   |-- geckodriver.exe
|   `-- msedgedriver.exe
```

Si el navegador se actualiza y el driver queda desfasado, reemplaza el `.exe` correspondiente dentro de `drivers/`.

## Programación automática

El ejecutor automático usa `json/programacion_test.json` para planificar ejecuciones.
La interfaz mantiene su propio monitor mientras la ventana está abierta.
Si necesitas que la programación siga activa con la interfaz cerrada, inicia el ejecutor autónomo manualmente.

Si necesitas iniciarlo manualmente para pruebas o diagnóstico, puedes usar:

```powershell
python run.py --autonomous
```

También se mantiene compatibilidad con:

```powershell
python json\ejecutor_autonomo.py
```

## Build portable

El script `build_exe.bat` genera:

- `dist/OsocioFormAutomation.exe`
- `dist/OsocioFormAutomation_portable/`

Dentro de `dist/OsocioFormAutomation_portable/` se crean:

- `OsocioFormAutomation.exe`
- `Abrir_Osocio_Form_Automation.bat`
- carpetas externas `data/`, `drivers/`, `json/`, `resultados/` y `temporales/`

Los drivers deben seguir distribuyéndose manualmente dentro de `drivers/`.

## Notas importantes

- El proyecto ya usa `sys.executable` para ejecutar los scripts desde el entorno activo.
- Se normalizó la carpeta `json` para evitar problemas de mayúsculas/minúsculas.
- `requirements.txt` contiene las dependencias necesarias.
