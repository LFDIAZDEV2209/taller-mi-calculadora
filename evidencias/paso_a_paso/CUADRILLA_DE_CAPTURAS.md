# Evidencias del taller

> **Estado: completado.** Las capturas de pantalla de la interfaz de PSeInt
> fueron tomadas e insertadas en el documento de entrega. Este archivo queda
> como registro del procedimiento y como guía de resolución de problemas de
> PSeInt.

---

## Procedimiento aplicado

1. El algoritmo de `src/pseint/MiCalculadora.pseint` se digitó completo en
   PSeInt y se ejecutó una vez para confirmar que corría sin errores.
2. Se agregaron pausas `Esperar 30 Segundos` en los tres puntos donde el
   programa reporta un error de validación, y una pausa final, para que la
   ventana de ejecución permaneciera abierta al momento de capturar.
3. Se tomó una captura por cada paso del taller con la misma sesión de
   entrada, de modo que los resultados fueran comparables entre sí.
4. Las imágenes se insertaron en el documento sustituyendo las cajas
   reservadas que venían marcadas como pendientes.

## Datos de entrada de la sesión documentada

| Dato | Valor |
|---|---|
| Cantidad de números | `4` |
| Número 1 | `10` |
| Número 2 | `4` |
| Número 3 | `2` |
| Número 4 | `8` |
| Operación 1 | `1` (SUMA) |
| Operación 2 | `2` (RESTA) |
| Operación 3 | `3` (MULTIPLICACION) |
| Operación 4 | `4` (DIVISION) |

Resultados obtenidos: **24**, **-4**, **640** y **0.15625**.

## Captura adicional: división entre cero

Además de la sesión anterior, se documentó el caso en que un divisor es cero,
que es el error aritmético más frecuente de una calculadora:

1. Reiniciar el programa.
2. Cantidad `3`, números `50`, `0`, `5`.
3. Operaciones `1`, `2`, `3`, `4`.
4. El programa muestra `ERROR: NO SE PUEDE DIVIDIR ENTRE CERO` y continúa con
   el resumen y el historial sin detenerse.

La traza equivalente de la implementación de referencia está en
`evidencias/consola/02_division_entre_cero_p1.png`.

## Problemas frecuentes con PSeInt

**PSeInt dice que `ConvertirANumero` no existe.**
Está en *Configurar → Opciones del Lenguaje*. Active las funciones de cadena
y de conversión. Si su versión es muy antigua, no tendrá la función
disponible.

**PSeInt reporta un error de tipo en la conversión.**
El resultado de `ConvertirANumero` se asigna a una variable declarada como
`Real`. Si la variable se declaró como `Entero`, cámbiela a `Real`.

**PSeInt dice «índice fuera de rango».**
Revise la base de indexación en *Configurar → Opciones del Lenguaje*. Los
arreglos se declararon con una posición adicional de holgura para ser
independientes de esa configuración, pero si su versión numera desde 0 y el
programa la referencia en una posición no reservada, ajuste los bucles.

**Las tildes no aparecen en la salida.**
Es intencional. Los mensajes de salida están escritos sin tildes por
compatibilidad con la fuente del intérprete; está explicado en la nota de la
sección 4.1 del documento.

**El programa se cierra antes de poder capturar.**
Ese fue el motivo por el que se agregaron las pausas `Esperar 30 Segundos`.
Si su versión de PSeInt no las admite, deje la ventana abierta con el
visualizador de *Ejecutar → Ver algoritmo en ejecución*.
