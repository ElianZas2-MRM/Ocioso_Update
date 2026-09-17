"""Cuando el Excel pide expresamente NO completar un campo.

Hay dos razones distintas para querer dejar un campo sin llenar, y las dos salieron del
mismo formulario (Cadillac Brasil T1):

- Campos opcionales que a veces se quieren probar vacios. CPF y CEP son opcionales ahi, y
  la app los genera sola cuando la celda viene vacia. Eso esta bien como default, pero no
  habia forma de pedir lo contrario.
- Campos que el formulario tiene bloqueados. El dropdown de concesionario viene `disabled`
  a proposito, asi que no lo puede elegir nadie; igual quedaba anotado como "quedo sin
  elegir" y ensuciaba el resultado de una corrida que en realidad estuvo bien.

La regla, en una linea: **una celda vacia se comporta como siempre; para omitir un campo
hay que decirlo.** Es al revés de lo que podria parecer natural, y es deliberado: los
Excels que ya existen tienen celdas vacias por todos lados y ninguna de ellas queria decir
"omitir". Cambiar el significado de lo vacio habria cambiado, en silencio, el resultado de
cada corrida vieja.

No hace falta ninguna columna nueva: la marca va en la celda del propio campo, que ya
existe (Brasil tiene columnas CPF, CNPJ, CEP y Concesionario). Por eso tambien funciona en
los Excels que ya estan generados, y por eso no corre ninguna columna de lugar - la lectura
del Excel es por posicion.
"""

# Se acepta cualquiera de estas para no tener que acordarse de una sola. La canonica, la
# que conviene documentar y escribir, es el guion.
#
# "NO" queda deliberadamente afuera: en las columnas de checkbox ya significa otra cosa
# (destildar la casilla), y dos significados para la misma palabra en el mismo Excel es
# pedirle a alguien que se equivoque.
MARCAS = frozenset({
    "-",
    "--",
    "---",
    "vacio",
    "vacío",
    "omitir",
    "no completar",
    "nocompletar",
    "skip",
    "n/a",
})


def pide_omitir(valor):
    """¿Esta celda pide expresamente que el campo NO se complete?

    Vacio, None o cualquier valor real devuelven False: el default es completar como
    siempre. Un numero tampoco es una marca, aunque sea un 0 o un negativo - los CPF y CEP
    llegan del Excel como numeros.
    """
    if valor is None or isinstance(valor, (int, float, bool)):
        return False
    return str(valor).strip().lower() in MARCAS
