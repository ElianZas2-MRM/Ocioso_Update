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


def test_corridas_en_paralelo_no_pierden_anotaciones(tmp_path):
    rot = RotacionDelDia(carpeta=str(tmp_path), fecha="2026-10-23")
    hilos = [threading.Thread(target=rot.registrar, args=("Chile", {DEALER: f"D{i}"})) for i in range(20)]
    for h in hilos:
        h.start()
    for h in hilos:
        h.join()
    assert sum(rot.usos("Chile", DEALER).values()) == 20
