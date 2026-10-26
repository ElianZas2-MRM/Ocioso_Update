"""
rotacion_dealers.py
===================
Región, ciudad y concesionario que se eligen al azar no se repiten en las corridas del día.

Cuando el Excel no pide un valor para esos dropdowns, la app elige uno al azar. Elegido
"a ciegas" el mismo concesionario podía salir en dos forms del mismo día (o dos veces en
el mismo form). Esto lleva la cuenta de lo usado HOY, por país:

- Se elige al azar entre las opciones MENOS usadas hoy: nada se repite hasta que se
  usaron todas.
- Lo que pide el Excel siempre manda; igual se anota como usado, así el azar lo evita.
- Las corridas de prueba (sin envío) eligen con la rotación pero no anotan nada.

Se guarda en temporales/rotacion/<AAAA-MM-DD>.json (por PC, no va al repo). Varias
corridas en paralelo comparten el archivo: cada lectura+escritura se hace con un lock.
"""

import datetime
import json
import os
import random
import time
import unicodedata

DEALER, CIUDAD, REGION = "concesionario", "ciudad", "region"

# Palabras que identifican cada tipo de campo por id / name / etiqueta. El orden importa:
# "concesionario" antes que nada (un id como "dealer-city" es del concesionario).
_CLAVES = (
    (DEALER, ("dealer", "concesionari", "concessionari", "distribuidor", "agencia", "agência")),
    (CIUDAD, ("city", "ciudad", "cidade", "localidad", "comuna", "municipio")),
    (REGION, ("region", "regiao", "região", "state", "estado", "provincia", "departamento", "depto")),
)


def _norm(texto):
    s = unicodedata.normalize("NFKD", str(texto or "").lower())
    return " ".join("".join(c for c in s if not unicodedata.combining(c)).split())


def tipo_de_campo(*nombres):
    """'concesionario' / 'ciudad' / 'region' según el id, name o etiqueta del campo; None si
    no es ninguno de los tres (esos se siguen eligiendo al azar sin rotación)."""
    texto = " ".join(_norm(n) for n in nombres if n)
    if not texto:
        return None
    for tipo, claves in _CLAVES:
        if any(_norm(c) in texto for c in claves):
            return tipo
    return None


def _dir_rotacion():
    from osocio import paths
    return os.path.join(paths.TEMPORALES_DIR, "rotacion")


class RotacionDelDia:
    """Cuenta de uso del día por país y tipo de campo. Un objeto por corrida está bien:
    cada operación relee el archivo, así ve lo que anotaron las corridas en paralelo."""

    def __init__(self, carpeta=None, fecha=None):
        self.carpeta = carpeta or _dir_rotacion()
        self.fecha = fecha or datetime.date.today().isoformat()
        self.ruta = os.path.join(self.carpeta, f"{self.fecha}.json")
        self._lock = self.ruta + ".lock"

    # ── archivo con lock ──────────────────────────────────────────────────────────────
    # En Windows un archivo que otro está borrando, o que tiene abierto el antivirus /
    # OneDrive / otra corrida leyéndolo, da PermissionError por un instante. Eso es "ocupado,
    # probá de nuevo", no un error: antes mataba el hilo y la anotación se perdía.
    _REINTENTOS, _PAUSA = 40, 0.05

    def _tomar_lock(self, espera=5.0):
        os.makedirs(self.carpeta, exist_ok=True)
        fin = time.time() + espera
        while True:
            try:
                fd = os.open(self._lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                os.close(fd)
                return True
            except (FileExistsError, PermissionError):
                # Un lock de más de 30 s es de una corrida que se cortó: se descarta.
                try:
                    if time.time() - os.path.getmtime(self._lock) > 30:
                        os.remove(self._lock)
                        continue
                except OSError:
                    pass
                if time.time() > fin:
                    return False
                time.sleep(0.05)

    def _reintentar(self, accion):
        """accion() reintentando mientras Windows diga PermissionError (archivo ocupado)."""
        for _ in range(self._REINTENTOS - 1):
            try:
                return accion()
            except PermissionError:
                time.sleep(self._PAUSA)
        return accion()

    def _soltar_lock(self):
        try:
            self._reintentar(lambda: os.remove(self._lock))
        except OSError:
            pass

    def _leer(self, estricto=False):
        """Lo anotado hoy. Sin archivo (o roto) = nada anotado. Si el archivo existe pero no
        se puede abrir, `estricto` lo informa con la excepción en vez de devolver {}: quien
        va a escribir no puede tomar "no pude leer" por "no hay nada" y pisar lo anotado."""
        def abrir():
            with open(self.ruta, encoding="utf-8") as f:
                return json.load(f)
        try:
            datos = self._reintentar(abrir)
            return datos if isinstance(datos, dict) else {}
        except (FileNotFoundError, ValueError):
            return {}
        except OSError:
            if estricto:
                raise
            return {}

    def _escribir(self, datos):
        tmp = self.ruta + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(datos, f, ensure_ascii=False, indent=1)
        self._reintentar(lambda: os.replace(tmp, self.ruta))

    # ── consultas y elección ──────────────────────────────────────────────────────────
    def usos(self, pais, tipo):
        """{opción normalizada: veces usada hoy} para ese país y tipo."""
        return dict(self._leer().get(_norm(pais), {}).get(tipo, {}))

    def veces(self, pais, tipo, texto):
        return self.usos(pais, tipo).get(_norm(texto), 0)

    def elegir(self, pais, tipo, opciones, excluir=()):
        """Índice de la opción a elegir: al azar entre las menos usadas hoy. `excluir`
        (textos) se descarta si queda alguna otra. None si no hay opciones."""
        if not opciones:
            return None
        usos = self.usos(pais, tipo)
        excl = {_norm(e) for e in excluir}
        indices = [i for i, t in enumerate(opciones) if _norm(t) not in excl] or list(range(len(opciones)))
        minimo = min(usos.get(_norm(opciones[i]), 0) for i in indices)
        return random.choice([i for i in indices if usos.get(_norm(opciones[i]), 0) == minimo])

    def registrar(self, pais, elegidos):
        """Anota como usados hoy los valores {tipo: texto} de un form."""
        elegidos = {t: v for t, v in (elegidos or {}).items() if t and v}
        if not elegidos:
            return
        tomado = self._tomar_lock()
        try:
            datos = self._leer(estricto=True)
            por_pais = datos.setdefault(_norm(pais), {})
            for tipo, texto in elegidos.items():
                cuenta = por_pais.setdefault(tipo, {})
                cuenta[_norm(texto)] = cuenta.get(_norm(texto), 0) + 1
            self._escribir(datos)
        except OSError as e:
            # Mejor perder esta anotación que borrar las del día; la corrida sigue.
            print(f"⚠️ Rotación: no se pudo anotar {elegidos} ({e})")
        finally:
            if tomado:
                self._soltar_lock()
