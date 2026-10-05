"""
Pruebas automaticas de la calculadora (taller "Mi calculadora").

Ejecucion:
    python -m unittest discover -s tests -v
    python tests/test_mi_calculadora.py -v
"""

from __future__ import annotations

import io
import sys
import unittest
from contextlib import redirect_stdout
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src" / "python"))

from mi_calculadora import Arreglo, Consola, TRAZAS, calcular, es_numero_valido  # noqa: E402


def ejecutar(entradas: list[str]) -> str:
    """Ejecuta la calculadora con una traza de entrada y devuelve la salida."""
    buffer = io.StringIO()
    with redirect_stdout(buffer):
        calcular(Consola(entradas))
    return buffer.getvalue()


class TestEsNumeroValido(unittest.TestCase):
    """Valida la Funcion EsNumeroValido del pseudocodigo."""

    def test_acepta_enteros(self) -> None:
        for texto in ("0", "7", "42", "1234"):
            with self.subTest(texto=texto):
                self.assertTrue(es_numero_valido(texto))

    def test_acepta_decimales(self) -> None:
        for texto in ("3.5", "0.001", "12.0"):
            with self.subTest(texto=texto):
                self.assertTrue(es_numero_valido(texto))

    def test_acepta_signo_negativo(self) -> None:
        for texto in ("-12", "-3.5", "-0"):
            with self.subTest(texto=texto):
                self.assertTrue(es_numero_valido(texto))

    def test_rechaza_cadena_vacia(self) -> None:
        self.assertFalse(es_numero_valido(""))

    def test_rechaza_texto(self) -> None:
        for texto in ("abc", "3a", "a3", "1,5", "5-3", "--3", "3-", " "):
            with self.subTest(texto=texto):
                self.assertFalse(es_numero_valido(texto))

    def test_rechaza_doble_separador_decimal(self) -> None:
        self.assertFalse(es_numero_valido("1.2.3"))

    def test_rechaza_solo_signo(self) -> None:
        self.assertFalse(es_numero_valido("-"))


class TestArreglo(unittest.TestCase):
    """Verifica el indexado desde 1 y el control de rango."""

    def setUp(self) -> None:
        self.arreglo = Arreglo("prueba", 5, float)

    def test_asignacion_y_lectura(self) -> None:
        self.arreglo[1] = 10.5
        self.arreglo[5] = 3.0
        self.assertEqual(self.arreglo[1], 10.5)
        self.assertEqual(self.arreglo[5], 3.0)

    def test_indice_cero_no_pertenece(self) -> None:
        with self.assertRaises(IndexError):
            self.arreglo[0] = 1.0  # noqa: B018

    def test_indice_fuera_de_rango(self) -> None:
        with self.assertRaises(IndexError):
            _ = self.arreglo[6]


class TestOperaciones(unittest.TestCase):
    """Verifica los resultados aritmeticos sobre los cuatro numeros."""

    def setUp(self) -> None:
        self.salida = ejecutar(TRAZAS["exito"]["entradas"])  # type: ignore[arg-type]

    def test_suma_de_cuatro_numeros(self) -> None:
        # 1 + 2 + 3 + 4 = 10
        self.assertIn("Operacion 1 -> SUMA", self.salida)
        self.assertIn("Resultado = 10", self.salida)

    def test_resta_sucesiva(self) -> None:
        # 1 - 2 - 3 - 4 = -8
        self.assertIn("Operacion 2 -> RESTA", self.salida)
        self.assertIn("Resultado = -8", self.salida)

    def test_multiplicacion(self) -> None:
        # 1 * 2 * 3 * 4 = 24
        self.assertIn("Operacion 3 -> MULTIPLICACION", self.salida)
        self.assertIn("Resultado = 24", self.salida)

    def test_division_encadenada(self) -> None:
        # 1 / 2 / 3 / 4 = 0.04166666667
        self.assertIn("Operacion 4 -> DIVISION", self.salida)
        self.assertIn("Resultado = 0.04166666667", self.salida)

    def test_cuatro_operaciones_correctas(self) -> None:
        self.assertIn("Operaciones con error                 : 0", self.salida)


class TestDivisionEntreCero(unittest.TestCase):
    """La division entre cero debe reportarse, nunca dehacer el programa."""

    def setUp(self) -> None:
        self.salida = ejecutar(TRAZAS["division_cero"]["entradas"])  # type: ignore[arg-type]

    def test_reporta_error_de_division(self) -> None:
        self.assertIn("ERROR: NO SE PUEDE DIVIDIR ENTRE CERO", self.salida)

    def test_conta_una_operacion_con_error(self) -> None:
        self.assertIn("Operaciones con error                 : 1", self.salida)

    def test_las_demas_operaciones_siguen(self) -> None:
        self.assertIn("Operaciones ejecutadas correctamente: 3", self.salida)

    def test_registra_error_en_el_historial(self) -> None:
        self.assertIn("estado: ERROR: NO SE PUEDE DIVIDIR ENTRE CERO", self.salida)


class TestValidacionDeEntradas(unittest.TestCase):
    """Las entradas invalidas deben reintentarse, no aceptarse en silencio."""

    def setUp(self) -> None:
        self.salida = ejecutar(TRAZAS["validacion"]["entradas"])  # type: ignore[arg-type]

    def test_rechaza_cantidad_fuera_de_rango(self) -> None:
        self.assertIn("la cantidad debe estar entre 2 y 5", self.salida)

    def test_rechaza_texto_no_numerico(self) -> None:
        self.assertIn("Entrada invalida. Se admiten solo numeros", self.salida)

    def test_rechaza_opcion_invalida(self) -> None:
        self.assertIn("Opcion invalida: digite un numero entre 1 y 4", self.salida)

    def test_termina_con_cuatro_operaciones(self) -> None:
        self.assertIn("Fin de la ejecucion de MiCalculadora.", self.salida)


class TestOrdenLibre(unittest.TestCase):
    """El despacho por nombre debe funcionar con cualquier orden de eleccion."""

    def setUp(self) -> None:
        self.salida = ejecutar(TRAZAS["orden_libre"]["entradas"])  # type: ignore[arg-type]

    def test_registra_orden_ingresado(self) -> None:
        self.assertIn("Operacion 1 -> DIVISION", self.salida)
        self.assertIn("Operacion 2 -> RESTA", self.salida)

    def test_resultados_correctos(self) -> None:
        self.assertIn("Resultado = 4", self.salida)   # 20 / 5
        self.assertIn("Resultado = 15", self.salida)  # 20 - 5
        self.assertIn("Resultado = 25", self.salida)  # 20 + 5

    def test_sin_errores(self) -> None:
        self.assertIn("Operaciones con error                 : 0", self.salida)


class TestTodasLasTrazas(unittest.TestCase):
    """Ninguna traza definida puede terminar de forma inesperada."""

    def test_todas_las_trazas_terminan(self) -> None:
        for nombre, caso in TRAZAS.items():
            with self.subTest(traza=nombre):
                salida = ejecutar(list(caso["entradas"]))  # type: ignore[arg-type]
                self.assertIn("Fin de la ejecucion de MiCalculadora.", salida)


if __name__ == "__main__":
    unittest.main(verbosity=2)