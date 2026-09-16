# Documentación de Osocio Form Automation

## Para empezar

| Documento | Para qué sirve |
|---|---|
| [`../README.md`](../README.md) | Manual de uso: qué hace cada pestaña, con capturas. Empezá por acá si vas a **usar** la app. |
| [`desarrollo.md`](desarrollo.md) | Cómo levantar el entorno, correr los tests y armar el `.exe`. Empezá por acá si vas a **tocar el código**. |
| [`arquitectura.md`](arquitectura.md) | Qué hace cada parte y por dónde pasa una corrida. |
| [`../CONTRIBUTING.md`](../CONTRIBUTING.md) | Cómo trabajar con ramas y PRs en este repo. |

## Referencia

| Documento | Contenido |
|---|---|
| [`mapa_interfaz.md`](mapa_interfaz.md) | Mapa de `main_interface.py`, que son 4.900 líneas en una sola función. Dice en qué rango de líneas está cada cosa. |
| [`historial_build.md`](historial_build.md) | Historial de cambios del script de build. Vivía dentro de `build.bat`. |
| [`README_HISTORIAL_ANTERIOR.md`](README_HISTORIAL_ANTERIOR.md) | Historial de versiones anteriores de la app. |
| [`resumen_de_cambios.md`](resumen_de_cambios.md) | Resumen de cambios acumulado. |
| [`Arquitectura_Consolidada_GDCP.md`](Arquitectura_Consolidada_GDCP.md) | **Histórico.** Describe la estructura de carpetas anterior al paquete `osocio/`, así que las rutas que menciona ya no existen. Para la estructura vigente, ver [`arquitectura.md`](arquitectura.md). |

## Documentos que no viven en el repo

Estos tres pesaban 3,9 MB entre los tres y no se pueden leer ni comparar desde GitHub, así
que se guardan aparte:

| Documento | Dónde está |
|---|---|
| `Osocio_Form_Automation_Guia.pptx` (3,0 MB) | _pendiente: link de SharePoint_ |
| `Documentation_Osocio_Form_Automation_EN_v2.docx` (884 KB) | _pendiente: link de SharePoint_ |
| `Arquitectura_Consolidada_Osocio_Form_Automation.docx` (35 KB) | _pendiente: link de SharePoint_ |

> Los archivos siguen estando en el historial de git, así que se pueden recuperar con
> `git show <commit>:docs/<archivo>` si hiciera falta.

## Capturas

`screenshots/` tiene las 28 capturas que usa el manual de uso, y `como_se_usa/` el
documento de bienvenida. Ninguna de las dos se empaqueta en el `.exe`: son solo
documentación.
