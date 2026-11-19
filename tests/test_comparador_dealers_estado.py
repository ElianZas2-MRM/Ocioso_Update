"""Tests de clasificación de estado y colores del Comparador de Dealers."""

from osocio.interface.dealer_comparator_ui import clasificar_resultado_comparador


def test_todo_ok_sin_avisos_es_verde():
    estado, icono, color, titulo, sufijo = clasificar_resultado_comparador({"PASS": 41, "FAIL": 0})
    assert estado == "ok"
    assert icono == "✓"
    assert color == "#82E0AA"
    assert titulo == "Comparación completada"
    assert "Todo OK" in sufijo


def test_con_extras_es_amarillo_aviso():
    estado, icono, color, titulo, sufijo = clasificar_resultado_comparador(
        {"PASS": 41, "FAIL": 0, "EXTRA": 9}
    )
    assert estado == "aviso"
    assert icono == "⚠"
    assert color == "#F8C471"
    assert "avisos" in titulo.lower()
    assert "9 aviso(s)" in sufijo


def test_con_duplicado_oculto_nota_o_nombre_distinto_es_amarillo():
    for counts, disc in [
        ({"PASS": 10, "DUPLICADO": 2}, 0),
        ({"PASS": 10, "OCULTO": 1}, 0),
        ({"PASS": 10, "NOTA": 3}, 0),
        ({"PASS": 10}, 2),
    ]:
        estado, icono, color, _titulo, _sufijo = clasificar_resultado_comparador(
            counts, disclaimer_count=disc
        )
        assert estado == "aviso"
        assert icono == "⚠"
        assert color == "#F8C471"


def test_con_fail_o_missing_es_rojo_aunque_haya_extras():
    for counts in [
        {"PASS": 10, "FAIL": 1, "EXTRA": 5},
        {"PASS": 10, "MISSING": 2, "EXTRA": 1},
    ]:
        estado, icono, color, _titulo, sufijo = clasificar_resultado_comparador(counts)
        assert estado == "error"
        assert icono == "✕"
        assert color == "#F1948A"
        assert "con problemas" in sufijo


def test_detenido_y_error_de_ejecucion_son_rojos():
    estado_det, icono_det, color_det, titulo_det, _ = clasificar_resultado_comparador(
        {"PASS": 5}, detenido=True
    )
    assert (estado_det, icono_det, color_det) == ("detenido", "■", "#F1948A")
    assert "detenida" in titulo_det.lower()

    estado_err, icono_err, color_err, titulo_err, _ = clasificar_resultado_comparador(
        {"PASS": 0}, error_msg="No abrió Chrome"
    )
    assert (estado_err, icono_err, color_err) == ("error", "✕", "#F1948A")
    assert "error" in titulo_err.lower()
