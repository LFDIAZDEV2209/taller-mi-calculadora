# Taller «Mi calculadora» — Estructuras de Datos

Calculadora diseñada con **instrucciones y arreglos** en pseudocódigo sobre
**PSeInt**, con sus cuatro operaciones aritméticas, validación de entrada y
manejo del error de división entre cero.

**Autor:** Luis Diaz
**Curso:** Estructuras de Datos — Semana 6
**Entrega:** `docs/diaz_luis_taller.docx`

---

## Qué hace

Pide entre 2 y 5 números, registra 4 operaciones escolhidas por el usuario y
las ejecuta recorriendo los arreglos con un lazo:

| Operación | Regla | Ejemplo (10, 4, 2, 8) |
|---|---|---|
| Suma | n1 + n2 + n3 + n4 | 24 |
| Resta | n1 - n2 - n3 - n4 | -4 |
| Multiplicación | n1 x n2 x n3 x n4 | 640 |
| División | n1 / n2 / n3 / n4 | 0.15625 |

---

## Estructuras de datos empleadas

| Arreglo | Tipo | Función |
|---|---|---|
| `numeros` | Real | Datos ingresados por el usuario (paso I y II) |
| `nombresOperacion` | Cadena | Catálogo de las 4 operaciones (paso III) |
| `simboloOperacion` | Cadena | Símbolo de cada operación (paso III) |
| `operaciones` | Cadena | Operaciones elegidas por el usuario (paso IV) |
| `resultados` | Real | Resultado de cada operación (paso VI y VII) |
| `estadoOperacion` | Cadena | `CORRECTA` o el mensaje de error (paso VII) |

Dos decisiones de diseño que conviene destacar:

1. **Indexación independiente de la configuración.** Todos los arreglos se
   declaran con una posición extra de holgura y se usan los índices del 1 al
   n, de modo que el algoritmo funciona tanto si el intérprete numera desde 0
   como desde 1. Evita una clase completa de errores de ejecución.
2. **Despacho por nombre, no por posición.** El paso VI no asume que el
   usuario elija las operaciones en orden: busca secuencialmente el nombre
   almacenado en el arreglo `operaciones` dentro del catálogo.

---

## Estructura del repositorio

```
taller-mi-calculadora/
├── src/
│   ├── pseint/
│   │   └── MiCalculadora.pseint     Algoritmo del taller (PSeInt)
│   └── python/
│       └── mi_calculadora.py        Implementación de referencia
├── tests/
│   ├── test_mi_calculadora.py       27 pruebas automáticas
│   └── validar_pseint.py            Verificación estructural del pseudocódigo
├── docs/
│   ├── diaz_luis_taller.docx        Entregable del taller
│   ├── generar_documento.py         Genera el .docx
│   └── auditar_documento.py         Auditoría de ortografía
├── evidencias/
│   ├── generar_evidencias.py        Genera las figuras de consola
│   ├── consola/                     Salidas reales de ejecución (.png)
│   └── paso_a_paso/
│       └── CUADRILLA_DE_CAPTURAS.md Registro del procedimiento de capturas
└── README.md
```

---

## Cómo reproducir

Requiere Python 3.10 o superior.

```bash
# Ejecutar la calculadora de forma interactiva
python src/python/mi_calculadora.py

# Reproducir una traza concreta
python src/python/mi_calculadora.py exito
python src/python/mi_calculadora.py division_cero
python src/python/mi_calculadora.py validacion
python src/python/mi_calculadora.py orden_libre

# Ver las cuatro trazas seguidas
python src/python/mi_calculadora.py --demo

# Correr las 27 pruebas
python -m unittest discover -s tests -v

# Verificar la estructura del pseudocódigo antes de digitarlo en PSeInt
python tests/validar_pseint.py

# Regenerar las figuras de consola
pip install pillow
python evidencias/generar_evidencias.py

# Regenerar el documento de entrega
pip install python-docx
python docs/generar_documento.py

# Verificar la ortografía del documento
python docs/auditar_documento.py
```

### Trazas disponibles

| Traza | Qué verifica |
|---|---|
| `exito` | Las cuatro operaciones correctas sobre 4 números |
| `division_cero` | El error de división entre cero se reporta y el programa sigue |
| `validacion` | Se rechazan cantidad fuera de rango, texto y opción inválida |
| `orden_libre` | Las operaciones se despachan bien en cualquier orden |

---

## Estado de la evidencia

| Evidencia | Estado |
|---|---|
| Código PSeInt completo y verificado | Listo |
| 27 pruebas automáticas | 27/27 superadas |
| Verificación estructural del pseudocódigo | Sin hallazgos |
| Figuras de salida real de consola | 10 figuras generadas |
| Documento `.docx` con los 7 pasos | Generado |
| 10 capturas de la interfaz de PSeInt | Insertadas por el estudiante |

Las diez capturas fueron tomadas e insertadas en el documento. El
procedimiento aplicado y la resolución de problemas frecuentes de PSeInt
están en
[`evidencias/paso_a_paso/CUADRILLA_DE_CAPTURAS.md`](evidencias/paso_a_paso/CUADRILLA_DE_CAPTURAS.md).

El documento de entrega contiene únicamente lo que pide la consigna
(portada y desarrollo de los siete pasos). El material adicional —
implementación de referencia, pruebas, verificación estructural, trazas de
consola — está en este repositorio.

### Verificación estructural

Como PSeInt no se puede ejecutar en un entorno automatizado,
`tests/validar_pseint.py` comprueba por análisis estático lo que el intérprete
comprobaría: balance de bloques, declaración previa de los arreglos, rango de
los índices y declaración de las variables leídas. Conviene ejecutarlo antes de
pegar el algoritmo en el editor.

---

## Documento de entrega

El `.docx` se genera con `docs/generar_documento.py` y no se edita a mano,
para que el código y el documento nunca queden desalineados: los bloques de
pseudocódigo de cada paso se extraen automáticamente de
`src/pseint/MiCalculadora.pseint` y el script falla si algún paso no queda
cubierto.

Antes de entregar hay que:

1. Verificar que la portada tenga institución, docente y fecha completos.
2. Confirmar que las diez capturas están insertadas y no queda ninguna caja
   reservada.
3. Confirmar el nombre del archivo: `diaz_luis_taller.docx`.

> **No regenere el documento** con `docs/generar_documento.py` una vez
> insertadas las capturas: sobrescribiría el trabajo manual. El script avisa
> y se detiene si ya existe un archivo de entrega.

---

## Referencias

- Algar, M. y Fernández de Sevilla, M. (2019). Primeros tipos estructurados.
  En *Introducción práctica a la programación con Python* (pp. 205-272).
  Editorial Universidad de Alcalá.
- Joyanes, L. (2020). Las cadenas de caracteres. En *Fundamentos de
  programación: algoritmos, estructuras de datos y objetos* (pp. 293-316).
  McGraw-Hill.
- Romano, F. (2018). Iterating and making decisions: conditional programming.
  En *Learn Python Programming* (pp. 79-105). Packt Publishing.
- Liskov, B. (2008). Data in a program should be structured like the data it
  represents. En B. Lipman y J. Sneakers (Coords.), *Program synthesis*.
  Addison-Wesley.
- PSeInt. (s. f.). *PSeInt: editor de algoritmos en pseudocódigo*.
  <http://pseint.sourceforge.net/>