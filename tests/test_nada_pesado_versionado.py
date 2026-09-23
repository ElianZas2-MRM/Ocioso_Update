"""No vuelven a entrar binarios al repo.

En septiembre de 2026 hubo que reescribir el historial para purgar ~86 MB de drivers, un
.exe, un .zip de 91 MB y un .rar de 83 MB. El repo paso de 388 MB a 14 MB, pero eso obligo
a que todo el equipo re-clonara: reescribir el historial le cambia el SHA a todos los
commits, y un clon viejo ya no se puede fusionar con el repo nuevo.

Por eso esto se testea y no se deja solo en `.gitignore`:

- `.gitignore` se puede saltear con `git add -f`, a veces sin querer.
- Solo ignoraba la carpeta `drivers/`. Un binario en cualquier otro lado entraba igual.
- Un archivo que ya entro **no se saca sin otro rewrite**, y el rewrite es carisimo. La
  unica defensa barata es que no entre.

El umbral no es un numero magico: es "bastante mas que el archivo versionado mas grande
que tenemos hoy". Si algo lo pasa, o es un binario que no deberia estar, o es un cambio
deliberado que merece una linea en este archivo explicando por que.
"""
import os
import subprocess

import pytest

from osocio import paths

# El versionado mas grande hoy es un PNG de la interfaz (~0,9 MB). Dos megas deja lugar de
# sobra para una captura nueva y sigue atajando un driver (~15 MB) o un .docx con imagenes.
LIMITE_MB = 2.0

# Nombres de binarios de driver, sin importar en que carpeta esten.
DRIVERS = ("chromedriver", "geckodriver", "msedgedriver", "operadriver",
           "iedriverserver", "selenium-manager")

# Formatos que git no puede delta-comprimir: cada version pesa entera y para siempre.
FORMATOS_PESADOS = (".exe", ".dll", ".zip", ".rar", ".7z", ".msi",
                    ".docx", ".pptx", ".dotx", ".potx", ".xls", ".mp4", ".mov")

# Lo que ya estaba adentro antes de esta regla. No se saca por las malas: desversionarlo
# es una decision aparte, y de todos modos ya vive en el historial. Esta lista existe para
# que se vea, no para que se agrande: si sumas algo aca, deja escrito por que.
PERMITIDOS = {
    "docs/como_se_usa/Bienvenidos a Osocio Form Automation versión GDCP.docx",
}


def _archivos_versionados():
    """Lo que git tiene trackeado, con los nombres bien decodificados.

    Se usa `-z` a proposito: sin eso git escapa los nombres con acentos
    (`versi\\303\\263n`) y los deja entre comillas, asi que un `.docx` con tilde en el
    nombre no matchea por extension. Paso: un chequeo dijo "no hay ninguno" y habia uno.
    """
    try:
        salida = subprocess.run(
            ["git", "ls-files", "-z"], cwd=paths.BASE_DIR,
            capture_output=True, check=False,
        )
    except (OSError, FileNotFoundError):
        pytest.skip("git no esta disponible en este entorno")
    if salida.returncode != 0:
        pytest.skip("no es un repositorio git")
    return [b.decode("utf-8") for b in salida.stdout.split(b"\x00") if b]


def _tam_mb(rel):
    ruta = os.path.join(paths.BASE_DIR, rel)
    return os.path.getsize(ruta) / (1024 * 1024) if os.path.exists(ruta) else 0.0


def test_no_hay_binarios_de_driver_versionados():
    """Los drivers los baja y actualiza la app sola (osocio/utils/driver_updater.py).

    Versionarlos ademas de pesar no sirve: quedan desfasados solos cada vez que Chrome o
    Edge se autoactualizan.
    """
    culpables = [
        a for a in _archivos_versionados()
        if any(d in os.path.basename(a).lower() for d in DRIVERS)
        and not a.lower().endswith(".py")   # el codigo que los gestiona si va
    ]
    assert not culpables, (
        "hay binarios de driver versionados:\n  " + "\n  ".join(culpables)
    )


def test_no_entraron_formatos_que_git_no_puede_comprimir():
    nuevos = [
        a for a in _archivos_versionados()
        if a.lower().endswith(FORMATOS_PESADOS) and a not in PERMITIDOS
    ]
    assert not nuevos, (
        "entraron archivos que git no puede delta-comprimir (cada version pesa entera y "
        "para siempre):\n  " + "\n  ".join(f"{a} ({_tam_mb(a):.2f} MB)" for a in nuevos)
        + "\n\nSi es a proposito, sumalo a PERMITIDOS con el motivo."
    )


def test_ningun_archivo_versionado_pasa_el_limite():
    grandotes = [
        (a, _tam_mb(a)) for a in _archivos_versionados()
        if _tam_mb(a) > LIMITE_MB
    ]
    assert not grandotes, (
        f"hay archivos versionados de mas de {LIMITE_MB} MB:\n  "
        + "\n  ".join(f"{a} ({mb:.2f} MB)" for a, mb in grandotes)
        + "\n\nUn archivo que entra al historial no se saca sin reescribirlo, y eso "
          "obliga a todo el equipo a re-clonar."
    )


def test_la_lista_de_permitidos_sigue_siendo_real():
    """Si alguien desversiona el .docx, esta entrada sobra y conviene borrarla.

    Una lista de excepciones que ya no corresponden es como un test apagado: deja de
    proteger y nadie se entera.
    """
    versionados = set(_archivos_versionados())
    fantasmas = [a for a in PERMITIDOS if a not in versionados]
    assert not fantasmas, (
        "PERMITIDOS nombra archivos que ya no estan versionados; sacalos de la lista:\n  "
        + "\n  ".join(fantasmas)
    )
