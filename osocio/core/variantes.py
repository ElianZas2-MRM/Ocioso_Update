"""Variantes de un mercado: el mismo pais corrido contra otro tipo de formulario.

Una variante NO es un mercado nuevo. Comparte el `field_mapping`, los ids y el generador
de datos del mercado del que cuelga; lo unico propio es contra que formularios corre. Por
eso se resuelve con un sufijo en el nombre del Excel en vez de duplicar configuracion.

Hoy hay dos (más el T1, que es el mercado tal cual):

- `_T1`: los formularios de siempre. No es una variante: solo le pone nombre a sus Excels
  para distinguirlos de los T3 en `data/` (`…_Chrome_T1.xlsx`). Los Excels viejos, sin
  sufijo, se siguen leyendo si no hay uno `_T1`. Resultados, reportes y mail no cambian.
- `_T3`: los formularios 2.0 de Adobe AEM. Se estan migrando y van a terminar iguales a
  los T1, pero mientras tanto se corren aparte y sus resultados van a `resultados/t3/`.
- `_CDC`: Cadillac. Usa los mismos ids que Brasil y el mismo generador de datos, pero
  **es T1**: sus resultados van a `resultados/t1/` como los de cualquier otro mercado.

Por que existe este modulo
--------------------------
Mientras hubo una sola variante, el codigo tomo un atajo: preguntaba "¿tiene sufijo?" y
respondia "entonces es T3". Con dos variantes ese atajo es falso, y falla en silencio:
Cadillac tiene sufijo y NO es T3, asi que sus resultados se habrian guardado en la carpeta
de T3 y los reportes lo habrian etiquetado como tal.

Ademas el nombre de los archivos de una corrida se armaba en dos lugares distintos que
tenian que coincidir a mano — el motor, que los ESCRIBE, y el ejecutor autonomo, que los
BUSCA para adjuntarlos al mail. Ya se desincronizaron una vez y el mail dejo de salir. Acá
se arma una sola vez.
"""

import os

SIN_VARIANTE = ""
T1 = "_T1"
T3 = "_T3"
CDC = "_CDC"

# El modo que se elige arriba en la app. Cadillac (CDC) no es un modo: es T1.
MODO_T1 = "T1"
MODO_T3 = "T3"

# Como se nombra cada variante en los reportes y en el mail. El "{pais}" se reemplaza; una
# etiqueta sin el es un nombre propio que no depende del mercado (Cadillac es Cadillac).
_ETIQUETAS = {
    T3: "{pais} T3",
    CDC: "CDC",
}

# Nombre con el que cada navegador aparece en los archivos de salida.
_DISPOSITIVOS = {
    "chrome": "Chrome", "firefox": "Firefox", "edge": "Edge",
    # Corridas en la nube: resultados_Brasil_CDC_BS_Mac1.xlsx, en las mismas carpetas t1/t3.
    "browserstack_iphone": "BS_iPhone", "browserstack_android": "BS_Android",
    "browserstack_mac": "BS_Mac",
}


def normalizar(sufijo):
    """El sufijo en su forma canonica: sin espacios y en mayusculas.

    Llega de varios lados —la UI, `--excel-suffix` por linea de comandos, la
    programacion guardada en disco— asi que conviene no confiar en como viene escrito.
    """
    return str(sufijo or "").strip().upper()


def es_t3(sufijo):
    """¿Esta corrida es de formularios T3 (Adobe AEM 2.0)?

    Ojo: NO es lo mismo que "tiene sufijo". Cadillac tiene sufijo y no es T3. Confundir
    las dos cosas manda sus resultados a la carpeta equivocada sin que nadie se entere.
    """
    return normalizar(sufijo) == T3


def variante_del_modo(modo):
    """La variante con la que corre el modo elegido en la app: T3 o ninguna (T1)."""
    return T3 if str(modo or "").strip().upper() == MODO_T3 else SIN_VARIANTE


def nombre_excel_leads(pais, dispositivo, sufijo=SIN_VARIANTE):
    """Nombre del Excel de leads: `Lead_information_Formulario_Brasil_Chrome_T1.xlsx`.

    Sin variante es T1: el Excel nuevo lleva `_T1` para no confundirse con el `_T3`.
    """
    sufijo = normalizar(sufijo) or T1
    return f"Lead_information_Formulario_{pais}_{dispositivo}{sufijo}.xlsx"


def ruta_excel_leads(carpeta, pais, dispositivo, sufijo=SIN_VARIANTE):
    """Ruta del Excel de leads a leer o crear.

    Para T1 se prefiere el `…_T1.xlsx`; si no existe pero hay uno viejo sin sufijo, se usa
    ese. Así nadie tiene que renombrar los Excels que ya tenía.
    """
    ruta = os.path.join(carpeta, nombre_excel_leads(pais, dispositivo, sufijo))
    if normalizar(sufijo) in (SIN_VARIANTE, T1) and not os.path.exists(ruta):
        viejo = os.path.join(carpeta, f"Lead_information_Formulario_{pais}_{dispositivo}.xlsx")
        if os.path.exists(viejo):
            return viejo
    return ruta


def etiqueta_de_corrida(pais, sufijo=SIN_VARIANTE):
    """Como se nombra esta corrida en los reportes y en el mail."""
    plantilla = _ETIQUETAS.get(normalizar(sufijo))
    if not plantilla:
        return pais
    return plantilla.format(pais=pais)


def _nombre_dispositivo(navegador):
    limpio = str(navegador or "").strip().lower()
    if not limpio:
        return ""
    return _DISPOSITIVOS.get(limpio, limpio.capitalize())


def basenames_de_corrida(pais, navegador, programada=False, sufijo=SIN_VARIANTE):
    """Prefijos de los archivos que deja una corrida: (excel, carpeta de capturas).

    Al final se les pega el numero de corrida: `resultados_Brasil_Chrome1.xlsx`.

    La variante va en el nombre. Sin eso, una corrida de Cadillac escribiria
    `resultados_Brasil_Chrome1.xlsx` y se pisaria en la numeracion con las corridas
    normales de Brasil, que van a la misma carpeta.
    """
    dispositivo = _nombre_dispositivo(navegador)
    partes = [pais]
    variante = "" if normalizar(sufijo) == T1 else normalizar(sufijo).lstrip("_")
    if variante:
        partes.append(variante)
    if dispositivo:
        partes.append(dispositivo)
    cola = "_".join(partes)

    if programada:
        return f"Automatizacion_{cola}", f"Automatizacion_screenshots_{cola}"
    return f"resultados_{cola}", f"screenshots_{cola}"
