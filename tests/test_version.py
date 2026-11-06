import re

from osocio.version import VERSION


def test_version_tiene_formato_x_y():
    assert re.fullmatch(r"\d+\.\d+(\.\d+)?", VERSION)
