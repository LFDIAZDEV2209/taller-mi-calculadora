"""
TALLER "MI CALCULADORA" - Implementacion de referencia en Python
================================================================

Este modulo reproduce, instruccion por instruccion, el algoritmo
desarrollado en ``src/pseint/MiCalculadora.pseint``.

Proposito academico:
    1. Verificar la logica del pseudocodigo antes de digitarla en PSeInt.
    2. Servir como evidencia de ejecucion real (traza de consola).
    3. Documentar la equivalencia entre pseudocodigo y lenguaje tipado.

Convencion de indexacion:
    PSeInt permite configurar la base de los arreglos en 0 o en 1. Para
    que la equivalencia sea exacta e independiente de la configuracion de
    la maquina, la clase ``Arreglo`` replica la convencion de indices desde
    1 usada por el pseudocodigo, que a su vez fueaszada con una posicion
    adicional de holgura en la instruccion ``Dimension``.
"""

from __future__ import annotations

import sys
from dataclasses import dataclass, field

# ---------------------------------------------------------------------------
# Estructura de datos: arreglo con indexacion basada en 1
# ---------------------------------------------------------------------------


class Arreglo:
    """Arreglo unidimensional que acepta indices desde 1.

    Equivale a la instruccion ``Dimension <nombre>[n]`` de PSeInt. El
    parametro ``holgura`` reserva la posicion 0 para que el arreglo sea
    valido tanto si el interprete usa base 1 como base 0.
    """

    __slots__ = ("nombre", "_datos", "_tipo")

    def __init__(self, nombre: str, dimension: int, tipo: type) -> None:
        self.nombre = nombre
        self._tipo = tipo
        # indice logico i  ->  posicion i en la lista (base 1)
        self._datos: list[object] = [self._tipo() for _ in range(dimension + 1)]

    def __setitem__(self, indice: int, valor: object) -> None:
        if not 1 <= indice < len(self._datos):
            raise IndexError(f"indice {indice} fuera de rango para {self.nombre}")
        self._datos[indice] = valor

    def __getitem__(self, indice: int) -> object:
        if not 1 <= indice < len(self._datos):
            raise IndexError(f"indice {indice} fuera de rango para {self.nombre}")
        return self._datos[indice]

    def __len__(self) -> int:
        return len(self._datos) - 1

    def __repr__(self) -> str:
        return f"Arreglo({self.nombre}, {self!r})"


def es_numero_valido(texto: str) -> bool:
    """Equivale a la Funcion EsNumeroValido del pseudocodigo.

    Retorna ``True`` si la cadena representa un numero real valido:
    unicamente digitos, signo negativo opcional y como maximo un
    separador decimal.

    >>> es_numero_valido("3.5")
    True
    >>> es_numero_valido("-12")
    True
    >>> es_numero_valido("1.2.3")
    False
    >>> es_numero_valido("abc")
    False
    """
    digitos_validos = "0123456789"
    longitud = len(texto)
    puntos = 0
    digitos_encontrados = 0
    error = False
    pos = 1

    # Una cadena vacia nunca es un numero valido
    if longitud == 0:
        error = True

    # Se admite signo negativo unicamente al inicio
    if not error and texto[0] == "-":
        pos = 2

    while pos <= longitud and not error:
        caracter = texto[pos - 1]
        if caracter == ".":
            puntos += 1
            if puntos > 1:
                error = True
        else:
            # Se comprueba que el caracter pertenezca al conjunto de digitos
            digito_valido = caracter in digitos_validos
            if digito_valido:
                digitos_encontrados += 1
            else:
                error = True
        pos += 1

    return (not error) and puntos <= 1 and digitos_encontrados > 0


# ---------------------------------------------------------------------------
# Entrada / salida del interprete
# ---------------------------------------------------------------------------


class Consola:
    """Abstraccion minima de entrada y salida, equivalente a Escribir y Leer.

    En modo traza (cuando se entregan entradas predefinidas) se hace eco de
    cada dato digitado, de modo que la salida se lea igual que una sesion
    real registrada en la consola de PSeInt.
    """

    def __init__(self, entradas: list[str] | None = None) -> None:
        self.entradas = list(entradas or [])
        self.eco = bool(entradas)
        self._pos = 0
        self.lineas: list[str] = []

    def escribir(self, *partes: object) -> None:
        linea = "".join(_a_texto(p) for p in partes)
        self.lineas.append(linea)
        print(linea)

    def leer(self) -> str:
        if self._pos >= len(self.entradas):
            raise EOFError("se agotaron las entradas de la traza simulada")
        valor = self.entradas[self._pos]
        self._pos += 1
        if self.eco:
            # Se imite lo que PSeInt muestra cuando el usuario digita un dato
            self.escribir(valor)
        return valor

    def leer_numero(self) -> float:
        entrada = self.leer()
        while not es_numero_valido(entrada):
            self.escribir("Entrada invalida. Se admiten solo numeros (ej: -12, 3.5).")
            entrada = self.leer()
        return float(entrada)


def _a_texto(valor: object) -> str:
    """Convierte a la notacion que PSeInt produce en la instruccion Escribir."""
    if isinstance(valor, bool):
        return "Verdadero" if valor else "Falso"
    if isinstance(valor, float):
        if valor == int(valor):
            return str(int(valor))
        return f"{valor:g}"
    return str(valor)


# ---------------------------------------------------------------------------
# Programa principal
# ---------------------------------------------------------------------------


@dataclass
class Resultado:
    salidas: list[str] = field(default_factory=list)


def calcular(consola: Consola) -> Resultado:
    """Ejecuta la calculadora replicando el algoritmo del pseudocodigo."""

    # --- PASO I: declaracion (dimensionamiento) de los arreglos ------------
    numeros = Arreglo("numeros", 6, float)
    nombres_operacion = Arreglo("nombresOperacion", 5, str)
    simbolo_operacion = Arreglo("simboloOperacion", 5, str)
    operaciones = Arreglo("operaciones", 5, str)
    resultados = Arreglo("resultados", 5, float)
    estado_operacion = Arreglo("estadoOperacion", 5, str)

    consola.escribir("=" * 56)
    consola.escribir("              MI CALCULADORA - TALLER ARREGLOS")
    consola.escribir("            Estructuras de Datos - PSeInt")
    consola.escribir("=" * 56)
    consola.escribir("")

    # --- PASO II: solicitacion y almacenamiento de los numeros ------------
    while True:
        consola.escribir("Cuantos numeros desea operar (entre 2 y 5)?: ")
        cantidad = int(consola.leer())
        if cantidad < 2 or cantidad > 5:
            consola.escribir("Dato invalido: la cantidad debe estar entre 2 y 5.")
            continue
        break

    for i in range(1, cantidad + 1):
        consola.escribir(f"Ingrese el numero {i} de {cantidad}: ")
        numeros[i] = consola.leer_numero()

    consola.escribir("")
    consola.escribir("Contenido del arreglo numeros:")
    for i in range(1, cantidad + 1):
        consola.escribir(f"   numeros[{i}] = {_a_texto(numeros[i])}")
    consola.escribir("")

    # --- PASO III: catalogo de operaciones -------------------------------
    nombres_operacion[1] = "SUMA"
    nombres_operacion[2] = "RESTA"
    nombres_operacion[3] = "MULTIPLICACION"
    nombres_operacion[4] = "DIVISION"

    simbolo_operacion[1] = "+"
    simbolo_operacion[2] = "-"
    simbolo_operacion[3] = "*"
    simbolo_operacion[4] = "/"

    consola.escribir("Operaciones disponibles en el arreglo nombresOperacion:")
    for i in range(1, 5):
        consola.escribir(
            f"   Opcion {i}: {nombres_operacion[i]}  ({simbolo_operacion[i]})"
        )
    consola.escribir("")

    # --- PASO IV: seleccion y almacenamiento de las operaciones -----------
    consola.escribir("Seleccione las 4 operaciones a registrar en el arreglo operaciones.")
    for j in range(1, 5):
        while True:
            consola.escribir("")
            consola.escribir(f"Seleccion numero {j} de 4. Elija una opcion del catalogo:")
            for op in range(1, 5):
                consola.escribir(
                    f"   {op}. {nombres_operacion[op]}  ({simbolo_operacion[op]})"
                )
            consola.escribir("Opcion elegida: ")
            opcion = int(consola.leer())
            if opcion < 1 or opcion > 4:
                consola.escribir("Opcion invalida: digite un numero entre 1 y 4.")
                continue
            break
        operaciones[j] = nombres_operacion[opcion]
        consola.escribir(f"Registro -> operaciones[{j}] = {operaciones[j]}")

    # --- PASO V: lazo que itera el arreglo operaciones ---------------------
    # --- PASO VI: condicional multiple que ejecuta cada operacion ---------
    consola.escribir("")
    consola.escribir("Procesando el arreglo operaciones con el lazo Para...")
    for k in range(1, 5):
        # Busqueda secuencial del codigo de la operacion seleccionada
        codigo = 0
        for j in range(1, 5):
            if operaciones[k] == nombres_operacion[j]:
                codigo = j

        if codigo == 1:  # SUMA
            acumulado = numeros[1]
            for j in range(2, cantidad + 1):
                acumulado = acumulado + numeros[j]
            resultados[k] = acumulado
            estado_operacion[k] = "CORRECTA"
        elif codigo == 2:  # RESTA
            acumulado = numeros[1]
            for j in range(2, cantidad + 1):
                acumulado = acumulado - numeros[j]
            resultados[k] = acumulado
            estado_operacion[k] = "CORRECTA"
        elif codigo == 3:  # MULTIPLICACION
            acumulado = 1
            for j in range(1, cantidad + 1):
                acumulado = acumulado * numeros[j]
            resultados[k] = acumulado
            estado_operacion[k] = "CORRECTA"
        elif codigo == 4:  # DIVISION
            divisor_en_cero = False
            for j in range(2, cantidad + 1):
                if numeros[j] == 0:
                    divisor_en_cero = True
            if divisor_en_cero:
                resultados[k] = 0
                estado_operacion[k] = "ERROR: NO SE PUEDE DIVIDIR ENTRE CERO"
            else:
                acumulado = numeros[1]
                for j in range(2, cantidad + 1):
                    acumulado = acumulado / numeros[j]
                resultados[k] = acumulado
                estado_operacion[k] = "CORRECTA"
        else:  # De Otro Modo
            resultados[k] = 0
            estado_operacion[k] = "ERROR: OPERACION NO RECONOCIDA"

    # --- PASO VII: presentacion de resultados ------------------------------
    consola.escribir("")
    consola.escribir("=" * 56)
    consola.escribir("   PASO VII - RESULTADOS DE LAS OPERACIONES")
    consola.escribir("=" * 56)
    operaciones_con_error = 0
    for k in range(1, 5):
        consola.escribir("")
        consola.escribir(f"Operacion {k} -> {operaciones[k]}")
        if estado_operacion[k] == "CORRECTA":
            consola.escribir(f"   Resultado = {_a_texto(resultados[k])}")
        else:
            consola.escribir(f"   {estado_operacion[k]}")
            operaciones_con_error += 1

    consola.escribir("")
    consola.escribir("-" * 56)
    consola.escribir("Operaciones registradas en el arreglo : 4")
    consola.escribir(f"Operaciones ejecutadas correctamente: {4 - operaciones_con_error}")
    consola.escribir(f"Operaciones con error                 : {operaciones_con_error}")
    consola.escribir("-" * 56)

    # --- HISTORIAL: trazabilidad de la calculadora ------------------------
    consola.escribir("")
    consola.escribir("HISTORIAL DE OPERACIONES (operaciones / resultados / estado)")
    for k in range(1, 5):
        # La linea se emite en dos llamadas para no exceder el ancho util
        consola.escribir(
            f"   [{k}] {operaciones[k]}  |  resultado: {_a_texto(resultados[k])}"
        )
        consola.escribir(f"             |  estado: {estado_operacion[k]}")
    consola.escribir("")
    consola.escribir("Fin de la ejecucion de MiCalculadora.")

    return Resultado(salidas=consola.lineas)


# ---------------------------------------------------------------------------
# Trazas de verificacion
# ---------------------------------------------------------------------------

TRAZAS: dict[str, dict[str, object]] = {
    # Caso 1: los cuatro pasos del taller con resultado correcto
    "exito": {
        "descripcion": "Cuatro operaciones correctas sobre 4 numeros.",
        "entradas": ["4", "10", "4", "2", "8", "1", "2", "3", "4"],
    },
    # Caso 2: division entre cero
    "division_cero": {
        "descripcion": "La operacion DIVISION recibe un divisor en cero.",
        "entradas": ["3", "50", "0", "5", "1", "2", "3", "4"],
    },
    # Caso 3: validacion de entradas invalidas
    "validacion": {
        "descripcion": "Cantidad fuera de rango, texto no numerico y opcion invalida.",
        "entradas": [
            "9",        # cantidad invalida -> reintento
            "3",
            "8",        # numero valido
            "abc",      # no numerico  -> reintento
            "2.5",      # decimal valido
            "-4",       # negativo valido
            "7",        # opcion invalida -> reintento
            "1", "2", "3", "4",
        ],
    },
    # Caso 4: orden libre (se comprueba que el despacho por nombre funciona
    # aunque el usuario elija las operaciones fuera del orden del catalogo)
    "orden_libre": {
        "descripcion": "Operaciones elegidas fuera del orden del catalogo.",
        "entradas": ["2", "20", "5", "4", "2", "1", "4", "3"],
    },
}


def main(argv: list[str]) -> int:
    if len(argv) > 1 and argv[1] in TRAZAS:
        nombre = argv[1]
        caso = TRAZAS[nombre]
        print(f"# TRAZA: {nombre}")
        print(f"# {caso['descripcion']}")
        print("# " + "-" * 55)
        print()
        calcular(Consola(list(caso["entradas"])))  # type: ignore[arg-type]
        return 0

    if len(argv) > 1 and argv[1] == "--demo":
        for nombre, caso in TRAZAS.items():
            print(f"\n########## TRAZA: {nombre} ##########")
            print(f"# {caso['descripcion']}")
            print("# " + "-" * 55 + "\n")
            calcular(Consola(list(caso["entradas"])))  # type: ignore[arg-type]
        return 0

    print("Uso: python mi_calculadora.py "
          "[exito|division_cero|validacion|orden_libre|--demo]")
    print("Sin argumentos la calculadora interactua solicita los datos por teclado.")
    return 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))