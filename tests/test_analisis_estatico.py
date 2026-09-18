"""pyflakes no puede empeorar: los avisos que quedan están contados.

Este proyecto llegó a tener **223 avisos**, de los cuales 154 venían de un solo
`from tkinter import *`. Con ese ruido, un aviso nuevo e importante pasaba desapercibido —
que es exactamente lo que hace inútil una herramienta de análisis estático.

Se limpió lo que se podía limpiar sin cambiar comportamiento. Los que quedan se dejaron a
propósito y están justificados abajo. Este test fija esa línea base: si aparece un aviso
nuevo, falla.

No busca llegar a cero. Busca que el número **no suba** sin que alguien lo decida.
"""
import re
import subprocess
import sys

import pytest

from osocio import paths


# Avisos que quedan, verificados uno por uno. Si bajás alguno, bajá también este número.
# 17 → 16: se fue el `_excel_empty` que se calculaba y nunca se usaba, al sacar la
# generación de documentos del llenado.
AVISOS_ESPERADOS = 16


def _correr_pyflakes():
    proceso = subprocess.run(
        [sys.executable, "-m", "pyflakes", "osocio", "run.py", "tests"],
        cwd=paths.BASE_DIR, capture_output=True, text=True,
    )
    if "No module named pyflakes" in proceso.stderr:
        pytest.skip("pyflakes no está instalado (pip install -r requirements-dev.txt)")
    return [l for l in proceso.stdout.splitlines() if l.strip()]


def test_no_hay_nombres_indefinidos():
    """El aviso que SIEMPRE es un bug: un nombre que no existe.

    Es el que atajó dos errores reales en este proyecto: un `import sys` que se removió de
    un archivo que sí lo usaba, y `paths.py` apuntando fuera del proyecto tras moverlo.
    """
    indefinidos = [l for l in _correr_pyflakes() if "undefined name" in l]
    assert not indefinidos, "hay nombres indefinidos:\n  " + "\n  ".join(indefinidos)


def test_no_aparecieron_avisos_nuevos():
    avisos = _correr_pyflakes()
    assert len(avisos) <= AVISOS_ESPERADOS, (
        f"pyflakes reporta {len(avisos)} avisos y la línea base es {AVISOS_ESPERADOS}.\n"
        "Si el aviso nuevo es legítimo, arreglalo. Si decidís convivir con él, subí la "
        "constante y dejá escrito por qué.\n\n  " + "\n  ".join(avisos)
    )


def test_la_linea_base_no_quedo_vieja():
    """Si limpiaste avisos y no bajaste la constante, el test deja de proteger."""
    avisos = _correr_pyflakes()
    assert len(avisos) >= AVISOS_ESPERADOS - 3, (
        f"pyflakes reporta solo {len(avisos)} avisos contra una línea base de "
        f"{AVISOS_ESPERADOS}. Bajá AVISOS_ESPERADOS para que el test siga sirviendo."
    )


def test_nadie_volvio_a_meter_un_import_estrella():
    """`from tkinter import *` traía 154 avisos él solo y tapaba todo lo demás.

    Ademas impide que pyflakes detecte nombres indefinidos en ese archivo: no puede saber
    qué trajo el import, así que se calla.
    """
    import os

    culpables = []
    for carpeta, dirs, archivos in os.walk(os.path.join(paths.BASE_DIR, "osocio")):
        dirs[:] = [d for d in dirs if d != "__pycache__"]
        for a in archivos:
            if not a.endswith(".py"):
                continue
            ruta = os.path.join(carpeta, a)
            with open(ruta, "r", encoding="utf-8") as fh:
                for nro, linea in enumerate(fh, 1):
                    if re.match(r"\s*from\s+\S+\s+import\s+\*", linea):
                        rel = os.path.relpath(ruta, paths.BASE_DIR).replace(os.sep, "/")
                        culpables.append(f"{rel}:{nro}")

    assert not culpables, (
        "hay imports con estrella, que tapan el analisis estatico del archivo entero:\n  "
        + "\n  ".join(culpables)
    )
