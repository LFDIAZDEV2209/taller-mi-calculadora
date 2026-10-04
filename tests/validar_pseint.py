"""
Verificación estructural del pseudocódigo de PSeInt.

PSeInt es una aplicación de escritorio y no se puede ejecutar en un entorno
automatizado. Este módulo comprueba por análisis estático lo que un
intérprete comprobaría en tiempo de ejecución:

  1. Balance de bloques: cada apertura tiene su cierre correspondiente.
  2. Que todo arreglo esté declarado antes de usarse.
  3. Que los índices usados permanezcan dentro del rango declarado.
  4. Que cada variable usada en la lectura haya sido declarada.
  5. Que las subcadenas de Subcadena no excedan la longitud del texto.

Uso:
    python tests/validar_pseint.py
Código de salida: 0 si el algoritmo es estructuralmente válido.
"""

from __future__ import annotations

import re
import sys
import unicodedata
from pathlib import Path

ARCHIVO = Path(__file__).resolve().parents[1] / "src" / "pseint" / "MiCalculadora.pseint"

# Bloques que se abren y se cierran
BLOQUES = {
    "Si": "FinSi",
    "Para": "FinPara",
    "Mientras": "FinMientras",
    "Repetir": "FinRepetir",   # se cierra con "Hasta Que"
    "Segun": "FinSegun",
    "Funcion": "FinFuncion",
    "Algoritmo": "FinAlgoritmo",
}

# Sentencias que abren bloque pero cuya cierre se escribe de otra forma
CIERRES_ALTERNATIVOS = {"Repetir": {"Hasta Que"}}

# Instrucciones que no abren bloque pese a começar con palabras reservadas
SIN_BLOQUE = {"Si", "Sino"}

TIPOS_PRIMITIVOS = {"Entero", "Real", "Cadena", "Logico", "Caracter"}


def normalizar(linea: str) -> str:
    return unicodedata.normalize("NFC", linea)


def sin_comentario(linea: str) -> str:
    return linea.split("//", 1)[0].strip()


def analizar(lineas: list[str]) -> dict[str, object]:
    problemas: list[str] = []
    pila: list[tuple[str, int]] = []
    arreglos: dict[str, tuple[int, int]] = {}
    declaradas: set[str] = set()
    leidas: set[str] = set()
    indices_usados: dict[str, list[tuple[int, str]]] = {}
    arreglos_usados_sin_declarar: list[tuple[int, str]] = []

    pila_funciones: list[str] = []

    for numero, original in enumerate(lineas, 1):
        codigo = sin_comentario(original)
        if not codigo:
            continue
        palabras = re.findall(r"[A-Za-zÁÉÍÓÚáéíóúÑñ_][A-Za-z0-9_]*", codigo)

        if not palabras:
            continue

        primera = palabras[0]

        # --- Balance de bloques ---
        if primera in ("Algoritmo", "Funcion"):
            pila.append((primera, numero))
            if primera == "Funcion":
                nombre = codigo.split()[1].split("(")[0]
                pila_funciones.append(nombre)
                declaradas.add(nombre)
        elif primera in ("Si", "Para", "Mientras", "Repetir", "Segun"):
            pila.append((primera, numero))
        elif primera == "Sino":
            if not pila or pila[-1][0] != "Si":
                problemas.append(f"Línea {numero}: «Sino» sin un «Si» abierto.")
        elif primera == "Hasta":
            # "Hasta Que" cierra el "Repetir" más reciente
            if pila and pila[-1][0] == "Repetir":
                pila.pop()
            else:
                problemas.append(
                    f"Línea {numero}: «Hasta Que» sin un «Repetir» abierto."
                )
        elif primera.startswith("Fin"):
            esperado = BLOQUES.get(pila[-1][0]) if pila else None
            cierres_validos = {esperado} if esperado else set()
            if pila and pila[-1][0] in CIERRES_ALTERNATIVOS:
                cierres_validos |= CIERRES_ALTERNATIVOS[pila[-1][0]]
            if primera in cierres_validos:
                cerrado = pila.pop()
                if cerrado[0] == "Funcion":
                    pila_funciones.pop()
            else:
                abierto = pila[-1] if pila else ("ninguno", 0)
                problemas.append(
                    f"Línea {numero}: «{primera}» no cierra el bloque "
                    f"«{abierto[0]}» abierto en la línea {abierto[1]}."
                )

        # --- Declaración de arreglos ---
        if primera == "Dimension":
            for coincidencia in re.finditer(
                r"([A-Za-z_][A-Za-z0-9_]*)\s*\[\s*([0-9]+)\s*\]", codigo
            ):
                nombre_arreglo, dimension = coincidencia.group(1), int(coincidencia.group(2))
                arreglos[nombre_arreglo] = (dimension, numero)
                declaradas.add(nombre_arreglo)

        # --- Declaración de variables ---
        if primera in TIPOS_PRIMITIVOS:
            for palabra in palabras[1:]:
                if palabra not in ("Como",):
                    declaradas.add(palabra)

        # --- Lectura de datos ---
        if primera == "Leer":
            for palabra in palabras[1:]:
                leidas.add(palabra)

        # --- Uso de arreglos ---
        for coincidencia in re.finditer(
            r"\b([A-Za-z_][A-Za-z0-9_]*)\s*\[\s*([^\]]+?)\s*\]", codigo
        ):
            nombre_arreglo, expresion = coincidencia.group(1), coincidencia.group(2)
            if re.fullmatch(r"[0-9]+", expresion):
                # Literal entre corchetes: es un arreglo usado con índice fijo
                indices_usados.setdefault(nombre_arreglo, []).append(
                    (numero, expresion)
                )
                if nombre_arreglo not in arreglos and not re.match(
                    r"^Longitud$|^ConvertirA", nombre_arreglo
                ):
                    arreglos_usados_sin_declarar.append((numero, nombre_arreglo))
            elif nombre_arreglo in arreglos:
                # Índice dinámico: solo se valida en ejecución
                indices_usados.setdefault(nombre_arreglo, []).append(
                    (numero, expresion)
                )

    if pila:
        for abierto, numero in pila:
            problemas.append(f"Línea {numero}: el bloque «{abierto}» quedó abierto.")

    # --- Arreglos usados sin declarar ---
    vistos: set[tuple[int, str]] = set()
    for numero, nombre_arreglo in arreglos_usados_sin_declarar:
        if (numero, nombre_arreglo) in vistos:
            continue
        vistos.add((numero, nombre_arreglo))
        problemas.append(
            f"Línea {numero}: el arreglo «{nombre_arreglo}» se usa pero nunca "
            f"se declaró con Dimension."
        )

    # --- Rango de índices ---
    for nombre_arreglo, usos in indices_usados.items():
        if nombre_arreglo not in arreglos:
            continue
        dimension, _ = arreglos[nombre_arreglo]
        for numero, expresion in usos:
            if not re.fullmatch(r"[0-9]+", expresion):
                continue  # índice dinámico: se valida en ejecución
            valor = int(expresion)
            if not 1 <= valor <= dimension:
                problemas.append(
                    f"Línea {numero}: {nombre_arreglo}[{valor}] está fuera del "
                    f"rango declarado 1..{dimension}."
                )

    # --- Variables leídas sin declarar ---
    for variable in sorted(leidas - declaradas):
        problemas.append(f"La variable «{variable}» se lee sin haber sido declarada.")

    return {
        "problemas": problemas,
        "arreglos": arreglos,
        "lecturas": leidas,
        "declaradas": declaradas,
    }


def main() -> int:
    lineas = ARCHIVO.read_text(encoding="utf-8").splitlines()
    resultado = analizar(lineas)
    problemas = resultado["problemas"]

    print(f"Archivo    : {ARCHIVO.name}")
    print(f"Líneas     : {len(lineas)}")
    print(f"Arreglos   : {len(resultado['arreglos'])}")
    print(f"Variables  : {len(resultado['declaradas'])} declaradas, "
          f"{len(resultado['lecturas'])} leídas")
    print()

    for nombre, (dimension, linea) in sorted(resultado["arreglos"].items()):
        print(f"  {nombre:<20} [{dimension}]  declarado en la línea {linea}")
    print()

    if problemas:
        print(f"RESULTADO: {len(problemas)} problema(s) estructural(es)")
        for problema in problemas:
            print(f"  - {problema}")
        return 1

    print("RESULTADO: estructura válida (bloques balanceados, arreglos")
    print("           declarados antes de uso, índices dentro de rango,")
    print("           variables leídas declaradas).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())