"""Variantes de un mercado: el mismo pais corrido contra otro tipo de formulario.

Una variante NO es un mercado nuevo. Comparte el `field_mapping`, los ids y el generador
de datos del mercado del que cuelga; lo unico propio es contra que formularios corre. Por
eso se resuelve con un sufijo en el nombre del Excel en vez de duplicar configuracion.

Hoy hay dos:

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

SIN_VARIANTE = ""
T3 = "_T3"
CDC = "_CDC"

# Como se nombra cada variante en los reportes y en el mail. El "{pais}" se reemplaza; una
# etiqueta sin el es un nombre propio que no depende del mercado (Cadillac es Cadillac).
_ETIQUETAS = {
    T3: "{pais} T3",
    CDC: "CDC",
}

# Nombre con el que cada navegador aparece en los archivos de salida.
_DISPOSITIVOS = {"chrome": "Chrome", "firefox": "Firefox", "edge": "Edge"}


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
    variante = normalizar(sufijo).lstrip("_")
    if variante:
        partes.append(variante)
    if dispositivo:
        partes.append(dispositivo)
    cola = "_".join(partes)

    if programada:
        return f"Automatizacion_{cola}", f"Automatizacion_screenshots_{cola}"
    return f"resultados_{cola}", f"screenshots_{cola}"
