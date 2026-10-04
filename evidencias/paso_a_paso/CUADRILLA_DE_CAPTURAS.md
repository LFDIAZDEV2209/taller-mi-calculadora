# Cuadrilla de capturas: cómo llenar el documento

Guía paso a paso para tomar las 10 capturas que pide el taller y pegarlas en
`docs/diaz_luis_taller.docx`. Cada captura tiene una caja amarilla marcada
como `PENDIENTE` en el capítulo 4 del documento.

---

## Antes de empezar

1. Descarga e instala **PSeInt** desde <http://pseint.sourceforge.net/>.
2. Abre el archivo `src/pseint/MiCalculadora.pseint` y copia **todo** el
   contenido al portapapeles. La forma más segura es abrir el archivo en un
   editor de texto, seleccionar todo (`Ctrl + A`) y copiar.
3. En PSeInt ve a **Archivo → Nuevo**, borra el contenido por defecto y pega.
4. Guarda con el nombre `MiCalculadora`.
5. Presiona **Ejecutar** (`F5`) una vez, antes de tomar cualquier captura,
   para confirmar que el algoritmo corre completo y sin errores.

> Si PSeInt reporta un error de sintaxis, revisa el capítulo 10 de este
> archivo. No tomés capturas con errores en pantalla.

---

## Datos de entrada de la sesión que vas a documentar

Usá estos valores en **todas** las capturas. Son los que aparecen en las
figuras del capítulo 6 y los que allowan contrastar los resultados:

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

Resultados esperados: **24**, **-4**, **640** y **0.15625**.

---

## Las 10 capturas

### CAPTURA 01 — Declaración de los arreglos
**Qué se ve:** el panel de código con las seis instrucciones `Dimension` y
las declaraciones de variables.
**Cómo:** `Archivo → Nuevo`, pega el bloque del Paso I, no ejecutes todavía.
**Sugerencia:** que la captura incluya la primera línea del algoritmo
(`Algoritmo MiCalculadora`) y la instrucción `FinAlgoritmo` del final.

### CAPTURA 02 — Validación de la cantidad
**Qué se ve:** la consola pidiendo la cantidad de números, el mensaje de
dato inválido y el reintento.
**Qué digitás:** `9` (inválido) → después `4` (válido).
**Por qué importa:** demuestra que el programa no acepta datos fuera de rango.

### CAPTURA 03 — Ingreso y almacenamiento de los números
**Qué se ve:** los cuatro números digitados y el recorrido del arreglo
`numeros` al imprimir su contenido.
**Qué digitás:** `10`, `4`, `2`, `8`.

### CAPTURA 04 — Catálogo de operaciones
**Qué se ve:** el menú impreso a partir del arreglo `nombresOperacion`, con
las cuatro operaciones y sus símbolos.

### CAPTURA 05 — Selección y validación de operaciones
**Qué se ve:** el proceso de selección, incluido el rechazo de una opción
inválida.
**Qué digitás:** `7` (inválido) → después `1`, `2`, `3`, `4`.

### CAPTURA 06 — Lazo y condicionales
**Qué se ve:** el mensaje «Procesando el arreglo operaciones con el lazo
Para...» en la consola, junto con el bloque `Segun` en el panel de código.
**Sugerencia:** tomá dos capturas, una del código y otra de la consola, y
unilas en una sola imagen.

### CAPTURA 07 — Los cuatro resultados  ← **la más importante**
**Qué se ve:** las cuatro operaciones con su resultado: 24, -4, 640 y 0.15625,
y el resumen de operaciones correctas.
**Por qué importa:** es la evidencia central del criterio de diseño de
algoritmos.

### CAPTURA 08 — Rechazo de entrada no numérica
**Qué se ve:** el mensaje «Entrada invalida...» y el reintento posterior.
**Qué digitás:** `abc` (rechazado) → después `10`.
**Por qué importa:** evidencia la función `EsNumeroValido`.

### CAPTURA 09 — La función auxiliar
**Qué se ve:** el código de `EsNumeroValido`, ubicado después de
`FinAlgoritmo`, y la invocación `EsNumeroValido(entrada)` resaltada en el
Paso II.

### CAPTURA 10 — Programa completo y botón Ejecutar
**Qué se ve:** la ventana principal de PSeInt con el programa entero escrito
y el botón Ejecutar resaltado.

---

## Captura opcional: división entre cero

No está en el documento porque no la exige el enunciado, pero conviene
tenerla por si el docente pregunta por el manejo de errores:

1. Reiniciá el programa.
2. Cantidad `3`, números `50`, `0`, `5`.
3. Operaciones `1`, `2`, `3`, `4`.
4. El programa debe mostrar
   `ERROR: NO SE PUEDE DIVIDIR ENTRE CERO` sin detenerse.

Es exactamente la figura 2 del capítulo 6 del documento.

---

## Cómo insertar las capturas en el documento

1. Abrí `docs/diaz_luis_taller.docx` en Microsoft Word.
2. Buscá la caja amarilla `CAPTURA 0X · PENDIENTE` correspondiente.
3. Seleccioná la caja completa y presioná `Supr` para eliminarla.
4. Insertá la imagen: **Insertar → Imágenes → Este dispositivo**.
5. Ajustá el tamaño arrastrando las esquinas. Dejalas en unos **12 cm de
   ancho**; el alto se conserva.
6. Escribí debajo un epígrafe corto, centrado y en cursiva, por ejemplo:
   *Figura 5. Ingreso de los cuatro números en el arreglo.*
7. Repetí para las diez capturas.
8. Revisá que ninguna imagen quedó pegada a un título o al borde de página.

### Ajustes recomendados de Word

- **Imágenes:** clic derecho → *Ajustar texto → Cuadrado*. Evita que el
  texto se parta alrededor de la imagen.
- **Interlineado:** el documento ya viene con 1,5. No lo cambies.
- **Saltos de página:** dejá los que trae. Si una imagen empuja una
  sección, ajustá solo esa página.

---

## Lista de verificación antes de entregar

- [ ] Los datos de la portada están completos: institución, docente y fecha.
- [ ] El nombre del archivo es `diaz_luis_taller.docx`.
- [ ] Las diez capturas están insertadas y numeradas del 01 al 10.
- [ ] No queda ninguna caja amarilla `PENDIENTE`.
- [ ] La figura de la división entre cero se ve si la incluiste.
- [ ] No hay errores de ortografía (correléalo con `python docs/auditar_documento.py`).
- [ ] El documento abre sin errores en Word.

---

## Capítulo 10 — Problemas frecuentes con PSeInt

**PSeInt dice que `ConvertirANumero` no existe.**
Está en el menú *Configurar → Opciones del Lenguaje*. Activá las funciones
de cadena y conversión. Si tu versión es muy antigua, cambiá la línea
`numeroIngresado <- ConvertirANumero(entrada)` por
`numeroIngresado <- ConvertirANumero(entrada) + 0`.

**PSeInt reporta «error de tipo» en la conversión.**
Declaraste `numeroIngresado` como `Entero`. Cambialo a `Real`.

**PSeInt dice «índice fuera de rango».**
Revisá la opción de indexación en *Configurar → Opciones del Lenguaje*. El
algoritmo fue escrito para ser independiente de esa configuración (por eso
los arreglos tienen una posición de holgura), pero si tu versión es muy
antigua y numera desde 0, cambiá los bucles para ir del 0 al n y sumá 1 al
mostrar.

**Las tildes no se ven en la salida.**
Es intencional. Los mensajes están escritos sin tildes por compatibilidad con
la fuente del intérprete. Está explicado en la nota de la sección 4.1 del
documento.

**El programa no pide los datos.**
Verificá que la primera línea sea exactamente
`Algoritmo MiCalculadora` y que exista un solo `FinAlgoritmo`.