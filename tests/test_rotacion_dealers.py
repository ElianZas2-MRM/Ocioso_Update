"""Rotación de región / ciudad / concesionario elegidos al azar en las corridas del día."""
import threading

from osocio.utils.rotacion_dealers import CIUDAD, DEALER, REGION, RotacionDelDia, tipo_de_campo


def test_tipo_de_campo_por_id_o_etiqueta():
    assert tipo_de_campo("dealer") == DEALER
    assert tipo_de_campo("", "Concessionária *") == DEALER
    assert tipo_de_campo("city") == CIUDAD
    assert tipo_de_campo("", "CIUDAD") == CIUDAD
    assert tipo_de_campo("region") == REGION
    assert tipo_de_campo("", "Provincia") == REGION
    assert tipo_de_campo("", "Estado") == REGION
    assert tipo_de_campo("models", "Modelo") is None


def test_no_repite_hasta_usar_todas(tmp_path):
    rot = RotacionDelDia(carpeta=str(tmp_path), fecha="2026-10-23")
    opciones = ["KOVACS", "DIVEMOTOR", "SALAZAR"]
    usados = []
    for _ in range(3):
        i = rot.elegir("Chile", DEALER, opciones)
        usados.append(opciones[i])
        rot.registrar("Chile", {DEALER: opciones[i]})
    assert sorted(usados) == sorted(opciones)            # los 3 distintos
    i = rot.elegir("Chile", DEALER, opciones)            # todas usadas una vez: vuelve a rotar
    assert opciones[i] in opciones


def test_cuenta_por_pais_y_normaliza(tmp_path):
    rot = RotacionDelDia(carpeta=str(tmp_path), fecha="2026-10-23")
    rot.registrar("Chile", {DEALER: "Kovacs Valparaíso"})
    assert rot.veces("chile", DEALER, "KOVACS VALPARAISO") == 1
    assert rot.veces("Peru", DEALER, "Kovacs Valparaíso") == 0
    assert rot.elegir("Chile", DEALER, ["Kovacs Valparaíso", "Otro"]) == 1


def test_excluir_y_otro_dia(tmp_path):
    rot = RotacionDelDia(carpeta=str(tmp_path), fecha="2026-10-23")
    assert rot.elegir("Chile", CIUDAD, ["A", "B"], excluir=["A"]) == 1
    assert rot.elegir("Chile", CIUDAD, ["A"], excluir=["A"]) == 0      # si no queda otra, la única
    rot.registrar("Chile", {CIUDAD: "A"})
    otro_dia = RotacionDelDia(carpeta=str(tmp_path), fecha="2026-10-24")
    assert otro_dia.veces("Chile", CIUDAD, "A") == 0


def _falla_las_primeras(n, real, error=PermissionError):
    """Envuelve `real` para que las primeras n llamadas den el error de Windows de
    "archivo ocupado" y después ande."""
    cuenta = {"n": 0}

    def envuelta(*a, **k):
        cuenta["n"] += 1
        if cuenta["n"] <= n:
            raise error(13, "Permission denied")
        return real(*a, **k)
    return envuelta


def test_lock_ocupado_por_windows_se_reintenta(tmp_path, monkeypatch):
    """La causa del test intermitente de abajo: crear el lock mientras otro hilo lo está
    borrando da PermissionError en Windows (no FileExistsError) y mataba el hilo."""
    import os
    rot = RotacionDelDia(carpeta=str(tmp_path), fecha="2026-10-23")
    monkeypatch.setattr(os, "open", _falla_las_primeras(3, os.open))
    rot.registrar("Chile", {DEALER: "Kovacs"})
    assert rot.veces("Chile", DEALER, "Kovacs") == 1


def test_archivo_ocupado_al_escribir_se_reintenta(tmp_path, monkeypatch):
    import os
    rot = RotacionDelDia(carpeta=str(tmp_path), fecha="2026-10-23")
    monkeypatch.setattr(os, "replace", _falla_las_primeras(3, os.replace))
    rot.registrar("Chile", {DEALER: "Kovacs"})
    assert rot.veces("Chile", DEALER, "Kovacs") == 1


def test_si_no_se_puede_leer_no_pisa_lo_anotado(tmp_path, monkeypatch):
    """Un "no pude leer" no es "no hay nada": antes se escribía encima sólo la anotación
    nueva y se perdían las del día."""
    import builtins
    rot = RotacionDelDia(carpeta=str(tmp_path), fecha="2026-10-23")
    rot.registrar("Chile", {DEALER: "Kovacs"})
    abrir = builtins.open

    def open_bloqueado(ruta, modo="r", *a, **k):
        if str(ruta) == rot.ruta and "r" in modo:
            raise PermissionError(13, "Permission denied")
        return abrir(ruta, modo, *a, **k)
    monkeypatch.setattr(builtins, "open", open_bloqueado)
    rot._PAUSA = 0
    rot.registrar("Chile", {DEALER: "Divemotor"})          # no se anota, pero no rompe
    monkeypatch.setattr(builtins, "open", abrir)
    assert rot.veces("Chile", DEALER, "Kovacs") == 1        # lo de antes sigue ahí


def test_corridas_en_paralelo_no_pierden_anotaciones(tmp_path):
    rot = RotacionDelDia(carpeta=str(tmp_path), fecha="2026-10-23")
    hilos = [threading.Thread(target=rot.registrar, args=("Chile", {DEALER: f"D{i}"})) for i in range(20)]
    for h in hilos:
        h.start()
    for h in hilos:
        h.join()
    assert sum(rot.usos("Chile", DEALER).values()) == 20
