"""
Genera el documento del taller "Mi calculadora" en formato .docx.

El script construye el entregable a partir de la fuente de verdad del
proyecto, de modo que el documento y el codigo nunca queden desalineados:
los bloques de pseudocodigo de cada paso se extraen automaticamente del
archivo src/pseint/MiCalculadora.pseint.

Uso:
    python docs/generar_documento.py
Salida:
    docs/diaz_luis_taller.docx
"""

from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

# ---------------------------------------------------------------------------
# Constantes del taller
# ---------------------------------------------------------------------------

RAIZ = Path(__file__).resolve().parents[1]
ARCHIVO_PSEINT = RAIZ / "src" / "pseint" / "MiCalculadora.pseint"
CARPETA_EVIDENCIAS = RAIZ / "evidencias" / "consola"
ARCHIVO_SALIDA = Path(__file__).resolve().parent / "diaz_luis_taller.docx"

AUTOR = "Luis Diaz"
CURSO = "Estructuras de Datos"
SEMANA = "Semana 6"
DOCENTE = "[Nombre del docente]"
FECHA_ENTREGA = "5 de octubre de 2026"
REPOSITORIO = "https://github.com/LFDIAZDEV2209/taller-mi-calculadora"
NOMBRE_ENTREGA = "diaz_luis_taller.docx"

AZUL = RGBColor(0x1F, 0x4E, 0x79)
GRIS = RGBColor(0x59, 0x59, 0x59)
ROJO = RGBColor(0xC0, 0x00, 0x00)

ALTO_MAXIMO_FIGURA = 18.5  # cm, dentro del área útil de la hoja

_CODIGO_COMANDOS = [
    "python src/python/mi_calculadora.py exito",
    "python src/python/mi_calculadora.py division_cero",
    "python src/python/mi_calculadora.py validacion",
    "python src/python/mi_calculadora.py orden_libre",
    "python -m unittest discover -s tests -v",
]

# Fragmentos extraidos del archivo PSeInt: (marca_inicio, marca_fin)
MARCAS_FRAGMENTOS: dict[str, tuple[str, str | None]] = {
    "I": ("PASO I -", "PASO II -"),
    "II": ("PASO II -", "PASO III -"),
    "III": ("PASO III -", "PASO IV -"),
    "IV": ("PASO IV -", "PASO V -"),
    "V_VI": ("PASO V -", "PASO VII -"),
    "VII": ("PASO VII -", "FinAlgoritmo"),
}


# ---------------------------------------------------------------------------
# Utilidades de bajo nivel sobre el documento
# ---------------------------------------------------------------------------


# Secuencias del esquema WordprocessingML. Los elementos de una tabla de
# propiedades deben aparecer en este orden; si se insertan fuera de orden,
# Word descarta silenciosamente la propiedad. Es la causa habitual de que el
# sombreado de una celda no se vea.
ORDEN_TCPR = (
    "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd",
    "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark",
)

ORDEN_PPR = (
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
    "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
    "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
    "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr",
)


def _insertar_en_orden(padre, elemento, orden: tuple[str, ...]) -> None:
    """Inserta un elemento respetando la secuencia del esquema OOXML."""
    nombre = elemento.tag.split("}")[-1]
    try:
        posicion = orden.index(nombre)
    except ValueError:
        padre.append(elemento)
        return

    for hijo in padre:
        nombre_hijo = hijo.tag.split("}")[-1]
        try:
            indice_hijo = orden.index(nombre_hijo)
        except ValueError:
            continue
        if indice_hijo > posicion:
            hijo.addprevious(elemento)
            return
    padre.append(elemento)


def sombrear(elemento_oxml, color_hex: str) -> None:
    """Aplica sombreado a una celda, a un párrafo o a un tcPr/pPr.

    Word descarta silenciosamente un w:shd mal ubicado, por eso se resuelve
    el contenedor de propiedades correcto antes de insertar.
    """
    if elemento_oxml.tag.endswith("}tc"):
        padre = elemento_oxml.get_or_add_tcPr()
        orden = ORDEN_TCPR
    elif elemento_oxml.tag.endswith("}p"):
        padre = elemento_oxml.get_or_add_pPr()
        orden = ORDEN_PPR
    else:
        padre = elemento_oxml
        orden = ORDEN_TCPR if elemento_oxml.tag.endswith("}tcPr") else ORDEN_PPR

    for previo in padre.findall(qn("w:shd")):
        padre.remove(previo)

    sombreado = OxmlElement("w:shd")
    sombreado.set(qn("w:val"), "clear")
    sombreado.set(qn("w:color"), "auto")
    sombreado.set(qn("w:fill"), color_hex)
    _insertar_en_orden(padre, sombreado, orden)


def bordes_celda(celda, color: str, estilo: str = "single", tamano: int = 6) -> None:
    tcPr = celda._tc.get_or_add_tcPr()
    for previo in tcPr.findall(qn("w:tcBorders")):
        tcPr.remove(previo)
    bordes = OxmlElement("w:tcBorders")
    for lado in ("top", "left", "bottom", "right"):
        elemento = OxmlElement(f"w:{lado}")
        elemento.set(qn("w:val"), estilo)
        elemento.set(qn("w:sz"), str(tamano))
        elemento.set(qn("w:space"), "0")
        elemento.set(qn("w:color"), color)
        bordes.append(elemento)
    _insertar_en_orden(tcPr, bordes, ORDEN_TCPR)


def margenes_celda(celda, *, izquierda: float = 0.1, derecha: float = 0.1,
                   arriba: float = 0.05, abajo: float = 0.05) -> None:
    """Reduce los márgenes internos de la celda para ganar ancho útil."""
    tcPr = celda._tc.get_or_add_tcPr()
    for previo in tcPr.findall(qn("w:tcMar")):
        tcPr.remove(previo)
    margenes = OxmlElement("w:tcMar")
    for lado, valor in (("top", arriba), ("left", izquierda),
                        ("bottom", abajo), ("right", derecha)):
        elemento = OxmlElement(f"w:{lado}")
        elemento.set(qn("w:w"), str(int(valor * 567)))
        elemento.set(qn("w:type"), "dxa")
        margenes.append(elemento)
    _insertar_en_orden(tcPr, margenes, ORDEN_TCPR)


def borde_parrafo(parrafo, color: str = "BFBFBF", tamano: int = 6) -> None:
    pPr = parrafo._p.get_or_add_pPr()
    for previo in pPr.findall(qn("w:pBdr")):
        pPr.remove(previo)
    bordes = OxmlElement("w:pBdr")
    for lado in ("top", "left", "bottom", "right"):
        elemento = OxmlElement(f"w:{lado}")
        elemento.set(qn("w:val"), "single")
        elemento.set(qn("w:sz"), str(tamano))
        elemento.set(qn("w:space"), "6")
        elemento.set(qn("w:color"), color)
        bordes.append(elemento)
    _insertar_en_orden(pPr, bordes, ORDEN_PPR)


def fuente_codigo(parrafo, tamano: float = 8.0) -> None:
    for corrida in parrafo.runs:
        corrida.font.name = "Consolas"
        corrida.font.size = Pt(tamano)
        corrida._element.rPr.rFonts.set(qn("w:eastAsia"), "Consolas")


ANCHO_UTIL_CM = 16.2


def _expandir_tabulaciones(linea: str, sangria: int = 4) -> str:
    """Convierte las tabulaciones del pseudocódigo en espacios.

    Las tabulaciones tienen un ancho variable según la fuente y el ajuste de
    tabulador, lo que desalinea el bloque y provoca que las líneas largas se
    partan. Con espacios, el ancho es predecible.
    """
    sangria_inicial = len(linea) - len(linea.lstrip("\t"))
    return " " * (sangria_inicial * sangria) + linea.lstrip("\t")


def ancho_maximo_codigo(bloques: list[list[str]]) -> int:
    return max(
        (len(_expandir_tabulaciones(linea)) for bloque in bloques for linea in bloque),
        default=0,
    )


def tamano_fuente_codigo(bloques: list[list[str]]) -> float:
    """Elige el mayor cuerpo de fuente con el que el bloque no se parte."""
    columnas = ancho_maximo_codigo(bloques)
    # Ancho de un carácter en Consolas: 0,55 em
    for tamano in (9.0, 8.5, 8.0, 7.5, 7.0, 6.5):
        if columnas * tamano * 0.55 / 28.35 <= ANCHO_UTIL_CM:
            return tamano
    return 6.5


def espaciado(parrafo, antes: float = 0, despues: float = 0, interlineado: float = 1.0):
    formato = parrafo.paragraph_format
    formato.space_before = Pt(antes)
    formato.space_after = Pt(despues)
    formato.line_spacing = interlineado
    return parrafo


# ---------------------------------------------------------------------------
# Bloques de contenido
# ---------------------------------------------------------------------------


def titulo(documento, texto: str, nivel: int = 1) -> None:
    parrafo_titulo = documento.add_heading(texto, level=nivel)
    for corrida in parrafo_titulo.runs:
        corrida.font.color.rgb = AZUL
        corrida.font.name = "Calibri"
    espaciado(parrafo_titulo, antes=14 if nivel == 1 else 10, despues=6)
    parrafo_titulo.paragraph_format.keep_with_next = True


def parrafo(documento, texto: str = "", *, negrita: bool = False, cursiva: bool = False,
            alineacion=None, tamano: float = 11, color=None, interlineado: float = 1.5):
    p = documento.add_paragraph()
    p.alignment = alineacion or WD_ALIGN_PARAGRAPH.JUSTIFY
    corrida = p.add_run(texto)
    corrida.bold = negrita
    corrida.italic = cursiva
    corrida.font.size = Pt(tamano)
    if color is not None:
        corrida.font.color.rgb = color
    espaciado(p, despues=6, interlineado=interlineado)
    return p


def vineta(documento, texto: str, nivel: int = 0) -> None:
    p = documento.add_paragraph(texto, style="List Bullet")
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Cm(0.75 + 0.6 * nivel)
    for corrida in p.runs:
        corrida.font.size = Pt(11)
    espaciado(p, despues=3, interlineado=1.3)
    return p


def bloque_codigo(documento, lineas: list[str], *, pie: str | None = None,
                  tamano: float = 8.0) -> None:
    """Inserta un bloque de código monoespaciado dentro de una celda única.

    Se usa una sola celda con sombreado y borde, en lugar de un borde por
    línea: el bloque queda compacto y no se fragmenta al cambiar de página.
    """
    if not lineas:
        return

    tabla_codigo = documento.add_table(rows=1, cols=1)
    tabla_codigo.alignment = WD_TABLE_ALIGNMENT.CENTER
    celda = tabla_codigo.rows[0].cells[0]
    celda.width = Cm(16.59)
    sombrear(celda._tc, "F4F4F4")
    bordes_celda(celda, "C8C8C8", "single", 6)
    margenes_celda(celda, izquierda=0.12, derecha=0.12)

    # Se elimina el párrafo vacío que python-docx crea por defecto
    celda._tc.remove(celda.paragraphs[0]._p)

    for linea in lineas:
        p = celda.add_paragraph()
        espaciado(p, antes=0, despues=0, interlineado=1.0)
        corrida = p.add_run(
            _expandir_tabulaciones(linea) if linea.strip() else " "
        )
        fuente_codigo(p, tamano)

    espaciado(documento.add_paragraph(), despues=2)

    if pie:
        epigrafe = documento.add_paragraph()
        epigrafe.alignment = WD_ALIGN_PARAGRAPH.CENTER
        corrida = epigrafe.add_run(pie)
        corrida.italic = True
        corrida.font.size = Pt(9)
        corrida.font.color.rgb = GRIS
        espaciado(epigrafe, antes=0, despues=12)


def marco_captura(documento, codigo: str, instruccion: str, detalle: str) -> None:
    """Caja reservada para insertar una captura de pantalla del estudiante."""
    tabla_captura = documento.add_table(rows=1, cols=1)
    tabla_captura.alignment = WD_TABLE_ALIGNMENT.CENTER
    celda = tabla_captura.rows[0].cells[0]
    celda.width = Cm(16.6)
    sombrear(celda._tc, "FFF8E1")
    bordes_celda(celda, "C08A00", "dashed", 8)

    encabezado = celda.paragraphs[0]
    encabezado.alignment = WD_ALIGN_PARAGRAPH.CENTER
    corrida = encabezado.add_run(codigo)
    corrida.bold = True
    corrida.font.size = Pt(11)
    corrida.font.color.rgb = ROJO
    espaciado(encabezado, antes=6, despues=2)

    cuerpo = celda.add_paragraph()
    cuerpo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    corrida = cuerpo.add_run(instruccion)
    corrida.font.size = Pt(10)
    espaciado(cuerpo, despues=2, interlineado=1.15)

    pie = celda.add_paragraph()
    pie.alignment = WD_ALIGN_PARAGRAPH.CENTER
    corrida = pie.add_run(detalle)
    corrida.italic = True
    corrida.font.size = Pt(9)
    corrida.font.color.rgb = GRIS
    espaciado(pie, antes=2, despues=6)

    espaciado(documento.add_paragraph(), despues=6)


def figura(documento, ruta: Path, epigrafe: str, ancho_cm: float = 15.6,
           alto_max_cm: float | None = None) -> None:
    """Inserta una imagen ajustada al ancho y, si hace falta, al alto."""
    from PIL import Image

    with Image.open(ruta) as imagen:
        proporcion = imagen.height / imagen.width

    ancho = ancho_cm
    if alto_max_cm and ancho * proporcion > alto_max_cm:
        ancho = alto_max_cm / proporcion

    documento.add_picture(str(ruta), width=Cm(ancho))
    documento.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = documento.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    corrida = p.add_run(epigrafe)
    corrida.italic = True
    corrida.font.size = Pt(9)
    corrida.font.color.rgb = GRIS
    espaciado(p, antes=2, despues=12)


def tabla(documento, cabeceras: list[str], filas: list[list[str]],
          anchos: list[float] | None = None) -> None:
    t = documento.add_table(rows=1, cols=len(cabeceras))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER

    for indice, texto in enumerate(cabeceras):
        celda = t.rows[0].cells[indice]
        sombrear(celda._tc, "1F4E79")
        p = celda.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        corrida = p.add_run(texto)
        corrida.bold = True
        corrida.font.size = Pt(9.5)
        corrida.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for indice_fila, fila in enumerate(filas):
        celdas = t.add_row().cells
        for indice, texto in enumerate(fila):
            if indice_fila % 2 == 1:
                sombrear(celdas[indice]._tc, "F4F7FB")
            p = celdas[indice].paragraphs[0]
            p.alignment = (WD_ALIGN_PARAGRAPH.CENTER if indice == 0
                           else WD_ALIGN_PARAGRAPH.LEFT)
            corrida = p.add_run(texto)
            corrida.font.size = Pt(9.5)
            espaciado(p, antes=2, despues=2, interlineado=1.0)

    if anchos:
        for fila in t.rows:
            for indice, ancho in enumerate(anchos):
                fila.cells[indice].width = Cm(ancho)

    espaciado(documento.add_paragraph(), despues=8)


def referencia(documento, texto: str) -> None:
    p = documento.add_paragraph(texto)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Cm(1.0)
    p.paragraph_format.first_line_indent = Cm(-1.0)
    for corrida in p.runs:
        corrida.font.size = Pt(10.5)
    espaciado(p, despues=8, interlineado=1.3)


def salto_de_pagina(documento) -> None:
    documento.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


# ---------------------------------------------------------------------------
# Extracción de fragmentos desde el archivo PSeInt
# ---------------------------------------------------------------------------


def extraer_fragmento(lineas: list[str], inicio: str, fin: str | None) -> list[str]:
    try:
        i_ini = next(i for i, linea in enumerate(lineas) if inicio in linea)
    except StopIteration as error:
        raise ValueError(f"No se encontró la marca de inicio: {inicio}") from error

    while i_ini > 0 and lineas[i_ini - 1].strip().startswith("//"):
        i_ini -= 1

    if fin is None:
        i_fin = len(lineas)
    else:
        try:
            i_fin = next(
                i for i, linea in enumerate(lineas) if i > i_ini and fin in linea
            )
        except StopIteration as error:
            raise ValueError(f"No se encontró la marca de fin: {fin}") from error
        while i_fin > i_ini and lineas[i_fin - 1].strip().startswith("//"):
            i_fin -= 1

    bloque = lineas[i_ini:i_fin]

    while bloque and (not bloque[0].strip() or set(bloque[0].strip()) <= {"-", "/"}):
        bloque.pop(0)
    while bloque and not bloque[-1].strip():
        bloque.pop()
    while bloque and set(bloque[-1].strip()) <= {"-", "/"}:
        bloque.pop()
    return bloque


def cargar_fuente() -> tuple[list[str], dict[str, list[str]]]:
    lineas = ARCHIVO_PSEINT.read_text(encoding="utf-8").splitlines()
    fragmentos = {
        clave: extraer_fragmento(lineas, inicio, fin)
        for clave, (inicio, fin) in MARCAS_FRAGMENTOS.items()
    }
    return lineas, fragmentos


def verificar_cobertura(lineas: list[str], fragmentos: dict[str, list[str]]) -> None:
    """Comprueba que los fragmentos del documento cubren todo el algoritmo."""
    cuerpo = "\n".join(lineas)
    cuerpo_algoritmo = cuerpo[
        cuerpo.index("Algoritmo MiCalculadora"):cuerpo.index("FinAlgoritmo")
    ]

    def normaliza(texto: str) -> list[str]:
        return [linea.strip() for linea in texto.splitlines() if linea.strip()]

    cubierto = normaliza("\n".join("\n".join(b) for b in fragmentos.values()))
    # La linea de cabecera del algoritmo no pertenece a ningun paso: se
    # documenta aparte, en el programa completo consolidado.
    faltantes = [
        l for l in normaliza(cuerpo_algoritmo)
        if l not in cubierto and l != "Algoritmo MiCalculadora"
    ]
    if faltantes:
        raise AssertionError(
            "Los fragmentos del documento no cubren el algoritmo completo:\n  "
            + "\n  ".join(faltantes)
        )


# ---------------------------------------------------------------------------
# Construcción del documento
# ---------------------------------------------------------------------------


def configurar_pagina(documento) -> None:
    for seccion in documento.sections:
        seccion.page_width = Cm(21.59)
        seccion.page_height = Cm(27.94)
        seccion.top_margin = Cm(2.5)
        seccion.bottom_margin = Cm(2.5)
        seccion.left_margin = Cm(2.5)
        seccion.right_margin = Cm(2.5)

    normal = documento.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.space_after = Pt(6)


def construir_portada(documento) -> None:
    p = documento.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    corrida = p.add_run("[Institución Educativa o Universidad]")
    corrida.font.size = Pt(13)
    corrida.font.color.rgb = GRIS
    espaciado(p, antes=24, despues=2)

    p = documento.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    corrida = p.add_run("FACULTAD DE INGENIERÍA DE SISTEMAS")
    corrida.font.size = Pt(13)
    corrida.font.color.rgb = GRIS
    espaciado(p, despues=2)

    p = documento.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    corrida = p.add_run("Estructuras de Datos")
    corrida.bold = True
    corrida.font.size = Pt(14)
    corrida.font.color.rgb = AZUL
    espaciado(p, despues=26)

    borde_parrafo(p, "1F4E79", 12)
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(18)

    for texto, tamano, negrita in [
        ("TALLER «MI CALCULADORA»", 26, True),
        ("Diseño de algoritmos con instrucciones y uso de arreglos", 14, False),
    ]:
        p = documento.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        corrida = p.add_run(texto)
        corrida.bold = negrita
        corrida.font.size = Pt(tamano)
        corrida.font.color.rgb = AZUL if negrita else GRIS
        espaciado(p, antes=0, despues=6 if negrita else 30)

    t = documento.add_table(rows=0, cols=2)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    datos = [
        ("Autor", AUTOR),
        ("Curso", CURSO),
        ("Semana", SEMANA),
        ("Docente", DOCENTE),
        ("Fecha de entrega", FECHA_ENTREGA),
        ("Archivo de entrega", NOMBRE_ENTREGA),
        ("Repositorio del código", REPOSITORIO),
    ]
    for etiqueta, valor in datos:
        celdas = t.add_row().cells
        sombrear(celdas[0]._tc, "1F4E79")
        p = celdas[0].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        corrida = p.add_run(etiqueta)
        corrida.bold = True
        corrida.font.size = Pt(10)
        corrida.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        espaciado(p, antes=3, despues=3)

        p = celdas[1].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        corrida = p.add_run(valor)
        corrida.font.size = Pt(10)
        espaciado(p, antes=3, despues=3)
        celdas[0].width = Cm(5.2)
        celdas[1].width = Cm(10.4)

    p = documento.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    corrida = p.add_run(
        "Las capturas de pantalla de la interfaz de PSeInt fueron tomadas por el "
        "estudiante. El código fuente y las trazas de verificación se encuentran "
        "en el repositorio del proyecto."
    )
    corrida.italic = True
    corrida.font.size = Pt(8.5)
    corrida.font.color.rgb = GRIS
    espaciado(p, antes=18)

    salto_de_pagina(documento)


def construir_indice(documento) -> None:
    titulo(documento, "Contenido", nivel=1)
    for entrada in [
        "1. Introducción",
        "2. Marco conceptual",
        "3. Preguntas orientadoras",
        "4. Desarrollo: diseño de la calculadora",
        "5. Programa completo consolidado",
        "6. Evidencias de ejecución",
        "7. Pruebas y análisis de resultados",
        "8. Conclusiones",
        "9. Referencias",
    ]:
        p = documento.add_paragraph()
        corrida = p.add_run(entrada)
        corrida.font.size = Pt(11)
        espaciado(p, despues=4)

    titulo(documento, "Nota para completar la entrega", nivel=2)
    parrafo(
        documento,
        "Este documento se entrega con los bloques de pseudocódigo completos y "
        "verificados, pero con las capturas de pantalla de la interfaz de PSeInt "
        "pendientes. Cada caja identificada como PENDIENTE DE CAPTURA indica "
        "exactamente qué debe capturarse y con qué valores de entrada. Una vez "
        "tomadas las capturas, se reemplazan las cajas por las imágenes y el "
        "taller queda listo para entregar.",
    )
    parrafo(
        documento,
        "Antes de imprimir deben completarse tres datos de la portada, que "
        "aparecen entre corchetes: el nombre de la institución, el nombre del "
        "docente y la fecha de entrega.",
    )
    salto_de_pagina(documento)


def construir_introduccion(documento) -> None:
    titulo(documento, "1. Introducción", nivel=1)
    parrafo(
        documento,
        "El presente taller corresponde a la actividad «Mi calculadora», en la "
        "que se aplica el diseño de algoritmos mediante instrucciones "
        "secuenciales, estructuras de control condicionales y estructuras de "
        "control iterativas, apoyado en el uso de arreglos como tipo de dato "
        "estructurado.",
    )
    parrafo(
        documento,
        "La consigna pide construir una calculadora que permita sumar, restar, "
        "multiplicar y dividir, siguiendo siete pasos que van desde declarar un "
        "arreglo para los números, pasar por la solicitud y la validación de los "
        "datos, hasta recorrer el arreglo de operaciones con un lazo y despachar "
        "cada caso con condicionales.",
    )

    titulo(documento, "1.1. Objetivo general", nivel=2)
    parrafo(
        documento,
        "Diseñar e implementar, en pseudocódigo sobre el programa PSeInt, una "
        "calculadora que realice las cuatro operaciones aritméticas empleando "
        "arreglos para almacenar los datos del usuario, las operaciones "
        "seleccionadas y los resultados obtenidos.",
    )

    titulo(documento, "1.2. Objetivos específicos", nivel=2)
    for texto in [
        "Declarar y utilizar arreglos de tipo numérico y de tipo cadena, "
        "distinguiendo el papel de cada uno dentro del algoritmo.",
        "Seleccionar instrucciones adecuadas y ordenarlas de manera que el "
        "programa sea funcional, correcto y legible.",
        "Aplicar estructuras de control condicionales para reemplazar cada "
        "operación aritmética.",
        "Aplicar estructuras de control iterativas para recorrer los arreglos "
        "sin duplicar código.",
        "Validar los datos de entrada para que el programa no falle ante datos "
        "mal digitados.",
        "Reconocer la importancia del uso de arreglos como primeros tipos "
        "estructurados en la resolución de problemas.",
    ]:
        vineta(documento, texto)

    titulo(documento, "1.3. Alcance", nivel=2)
    parrafo(
        documento,
        "La calculadora admite entre dos y cinco números y registra cuatro "
        "operaciones. La suma, la resta y la multiplicación se resuelven "
        "recorriendo todos los elementos del arreglo de números, y la división "
        "se resuelve como división encadenada, con control explícito del caso en "
        "que algún divisor sea cero. El programa conserva un historial de las "
        "operaciones ejecutadas. No se implementan operaciones adicionales como "
        "potencias, raíces ni funciones trigonométricas, porque no forman parte "
        "del alcance solicitado.",
    )


def construir_marco_conceptual(documento) -> None:
    titulo(documento, "2. Marco conceptual", nivel=1)
    parrafo(
        documento,
        "Este apartado sintetiza las ideas principales de los recursos básicos "
        "de la semana y de fuentes complementarias verificadas, y constituye el "
        "sustento conceptual de las decisiones de diseño tomadas en el capítulo 4.",
    )

    titulo(documento, "2.1. Arreglos como primer tipo estructurado", nivel=2)
    parrafo(
        documento,
        "Un arreglo es una colección finita y ordenada de elementos del mismo "
        "tipo, almacenados en posiciones contiguas de memoria y con acceso por "
        "índice. Su novedad frente a una variable simple no es la cantidad de "
        "datos que guarda, sino que la estructura de control del algoritmo deja "
        "de depender del número de variables: basta un índice para recorrer el "
        "arreglo (Algar y Fernández de Sevilla, 2019).",
    )
    parrafo(
        documento,
        "La consecuencia práctica es directa. Un programa que debe procesar una "
        "cantidad variable de datos, como la calculadora de este taller, se "
        "escribe con un solo lazo, en lugar de repetir el cálculo cuatro veces "
        "para el primer número, cuatro veces para el segundo, y así sucesivamente.",
    )
    parrafo(
        documento,
        "Según Liskov (2008), el acceso a un arreglo debe exponer las mismas "
        "tres operaciones que ya existen sobre un dato simple: crear, actualizar "
        "y recuperar. Este taller aplica las tres de forma explícita: crear en el "
        "paso I, actualizar en los pasos II y IV, y recuperar en los pasos VI y VII.",
    )

    titulo(documento, "2.2. La dimensión y la base de indexación", nivel=2)
    parrafo(
        documento,
        "La instrucción Dimension declara un arreglo indicando su nombre y la "
        "cantidad máxima de cada dimensión. Un detalle que suele pasar "
        "descuidado es la base de la indexación: el intérprete puede trabajar "
        "con la primera posición numerada como 0 o como 1, según su "
        "configuración. Un algoritmo que fija la posición 0 se rompe en cuanto "
        "cambia esa configuración.",
    )
    parrafo(
        documento,
        "Por esa razón, en este taller todos los arreglos se declaran con una "
        "posición adicional de holgura y se usan siempre los índices del 1 al n. "
        "De esta forma el algoritmo es válido con cualquiera de las dos bases. "
        "Es una decisión de robustez de bajo costo que evita una clase completa "
        "de errores de ejecución.",
    )

    titulo(documento, "2.3. Cadenas de caracteres en el menú", nivel=2)
    parrafo(
        documento,
        "La cadena es el tipo de dato que permite modelar texto. En el taller se "
        "emplea como vehículo de dos decisiones: el catálogo de operaciones "
        "(SUMA, RESTA, MULTIPLICACION, DIVISION) y el estado de cada operación "
        "(CORRECTA o el mensaje de error) (Joyanes, 2020).",
    )
    parrafo(
        documento,
        "La comparación de cadenas es sensible a mayúsculas y minúsculas, por lo "
        "que el catálogo se escribe siempre con la misma convención y el usuario "
        "nunca digita el texto: digita un número de opción y el algoritmo resuelve "
        "el nombre desde el arreglo. Este detalle evita errores de escritura y "
        "es buena práctica de diseño de interfaces.",
    )
    parrafo(
        documento,
        "Las funciones de cadena disponibles, como Longitud y Subcadena, se "
        "emplean en la función auxiliar EsNumeroValido para recorrer la entrada "
        "del usuario carácter por carácter y decidir si representa un número "
        "válido (Joyanes, 2020). La validación se hace sobre el texto y solo "
        "después se convierte a número, con lo cual se evita que el programa se "
        "detenga ante un dato inválido.",
    )

    titulo(documento, "2.4. De la instrucción al algoritmo", nivel=2)
    parrafo(
        documento,
        "Seleccionar una instrucción adecuada es, en esencia, decidir qué puede "
        "hacer el computador y en qué orden. Una instrucción mal puesta no es "
        "simplemente incorrecta: puede ser válida en un momento y dañar el "
        "algoritmo en otro. Los errores más frecuentes al trabajar con arreglos "
        "son el uso de un índice antes de declararlo, el acceso fuera del rango "
        "válido y la lectura de un arreglo antes de haberlo recorrido.",
    )
    parrafo(
        documento,
        "Romano (2018) resume el criterio práctico: una instrucción es adecuada "
        "cuando responde a una pregunta clara sobre el estado del algoritmo. Por "
        "eso este taller coloca al inicio una instrucción Escribir que anuncia lo "
        "que va a hacer, y al final de cada paso una instrucción que muestra el "
        "estado del arreglo. La traza del programa documenta su propio "
        "funcionamiento, que es justamente lo que se solicitó como evidencia.",
    )
    parrafo(
        documento,
        "La ausencia de estructuras de control se observa en la práctica: la "
        "forma de resolver las cuatro operaciones se ha escrito una sola vez, "
        "dentro del lazo del paso V, y se ha reutilizado para cualquier cantidad "
        "de operaciones que se agreguen en el futuro.",
    )

    titulo(documento, "2.5. Síntesis comparativa", nivel=2)
    tabla(
        documento,
        ["Concepto", "Arreglo numérico", "Arreglo de cadena", "Variable simple"],
        [
            ["Ejemplo en el taller", "numeros, resultados",
             "operaciones, estadoOperacion", "cantidad, opcion, acumulado"],
            ["Propósito", "Guardar los datos y los resultados del cálculo",
             "Guardar nombres de operaciones y estados",
             "Guardar un dato del control del lazo"],
            ["Operaciones del capítulo 4",
             "Crear: paso I. Actualizar: paso II. Recuperar: pasos VI y VII",
             "Crear: paso I. Actualizar: pasos III y IV. Recuperar: pasos VI y VII",
             "Crear y actualizar en el momento de usarla"],
            ["Si se omite", "El algoritmo debe duplicar el cálculo por cada dato",
             "El menú debe escribirse cuatro veces en el código",
             "Se pierde la posibilidad de reintentar la entrada"],
        ],
        anchos=[3.0, 4.4, 4.6, 4.6],
    )


def construir_preguntas(documento) -> None:
    titulo(documento, "3. Preguntas orientadoras", nivel=1)

    titulo(documento, "3.1. ¿Cuál es la importancia de seleccionar instrucciones "
                      "adecuadas para diseñar algoritmos?", nivel=2)
    parrafo(
        documento,
        "Porque un algoritmo es una secuencia ordenada de instrucciones, la "
        "calidad del programa depende por completo de que esas instrucciones "
        "sean correctas, completas y estén en el orden adecuado. Elegir bien una "
        "instrucción es elegir la operación primitiva correcta para el momento "
        "preciso en que se necesita.",
    )
    parrafo(
        documento,
        "La consecuencia práctica se observa en tres niveles. En cuanto a la "
        "corrección, una instrucción ejecutada fuera de orden produce resultados "
        "erróneos aunque cada instrucción por separado sea válida. En cuanto al "
        "rendimiento, un algoritmo que recorre un arreglo con un solo lazo es "
        "lineal en la cantidad de elementos, mientras que uno que repite bloques "
        "de código es cuadrático. En cuanto al mantenimiento, un algoritmo claro "
        "puede ampliarse; uno con instrucciones duplicadas obliga a corregir el "
        "mismo error en varios lugares.",
    )
    parrafo(
        documento,
        "En este taller la selección de instrucciones se evidencia en tres "
        "decisiones concretas: declarar los arreglos antes de usarlos, validar "
        "cada dato antes de convertirlo, y resolver las cuatro operaciones dentro "
        "de un único lazo en lugar de escribir cuatro bloques de código.",
    )

    titulo(documento, "3.2. ¿Cómo pueden utilizarse arreglos en la creación de un "
                      "programa en PSeInt?", nivel=2)
    parrafo(
        documento,
        "En PSeInt los arreglos se declaran con la instrucción Dimension, "
        "indicando el nombre y el máximo de cada dimensión, y se acceden "
        "escribiendo el nombre seguido del índice entre corchetes. A partir de "
        "ahí, un arreglo interviene en un algoritmo en cinco momentos: declararlo, "
        "recorrerlo con un lazo, leer sus elementos, escribir en ellos y "
        "consultar su contenido para tomar decisiones.",
    )
    parrafo(
        documento,
        "La calculadora de este taller usa un arreglo de forma distinta en cada "
        "paso: numeros como almacén de los datos del usuario; "
        "nombresOperacion y simboloOperacion como catálogo fijo; operaciones "
        "como registro de lo que el usuario eligió; y resultados y "
        "estadoOperacion como salida del cálculo. El punto más relevante es que "
        "el catálogo y el arreglo operaciones son de tipo cadena, lo cual permite "
        "decidir sobre el contenido del arreglo mediante comparación de cadenas "
        "y no solamente mediante operaciones numéricas.",
    )
    parrafo(
        documento,
        "Finalmente, conviene recordar dos precauciones prácticas: declarar el "
        "arreglo antes de hacer referencia a él, y controlar el rango del índice "
        "dentro de los límites declarados. Ambas se aplican en el paso V, donde "
        "el índice k se mantiene siempre entre 1 y 4, que es el rango real del "
        "arreglo operaciones.",
    )


def construir_desarrollo(documento, fragmentos: dict[str, list[str]],
                         tamano: float = 8.0) -> None:
    titulo(documento, "4. Desarrollo: diseño de la calculadora", nivel=1)

    titulo(documento, "4.1. Convenciones de diseño", nivel=2)
    tabla(
        documento,
        ["Aspecto", "Decisión tomada", "Justificación"],
        [
            ["Lenguaje", "PSeInt, pseudocódigo con estructura de bloques",
             "Exigido por la consigna del taller"],
            ["Archivo fuente", "src/pseint/MiCalculadora.pseint",
             "Código fuente del algoritmo"],
            ["Arreglos numéricos", "numeros y resultados, de tipo Real",
             "Requieren operaciones con decimales y valores negativos"],
            ["Arreglos de cadena",
             "nombresOperacion, simboloOperacion, operaciones, estadoOperacion",
             "Modelan el menú, la elección y el estado"],
            ["Indexación", "Índices del 1 al n, con una posición extra de holgura",
             "El algoritmo funciona con base 0 o con base 1"],
            ["Validación", "Función EsNumeroValido sobre la cadena de entrada",
             "Evita que el programa se detenga ante un dato inválido"],
            ["Control de errores", "División entre cero detectada antes de dividir",
             "Es el error aritmético más frecuente"],
            ["Evidencia", "27 pruebas automáticas y 4 trazas de consola",
             "Sustenta el criterio de resolución y evidencias"],
        ],
        anchos=[3.0, 6.6, 6.9],
    )
    parrafo(
        documento,
        "Nota sobre la notación: los identificadores y los mensajes de salida del "
        "algoritmo están escritos sin tildes, por una cuestión de compatibilidad "
        "de la fuente del intérprete. Esto no afecta la lógica del programa ni "
        "altera ningún resultado.",
        cursiva=True,
        tamano=9.5,
    )

    titulo(documento, "4.2. Paso I. Establecer el arreglo que almacena los números "
                      "ingresados por el usuario", nivel=2)
    parrafo(
        documento,
        "Se declaran seis arreglos. Dos de tipo Real, numeros y resultados, para "
        "los datos y para los resultados; y cuatro de tipo Cadena, para el "
        "catálogo de operaciones, sus símbolos, las operaciones elegidas y el "
        "estado de cada una. Cada declaración incluye una posición adicional de "
        "holgura, de modo que los índices del 1 al n sean válidos con cualquiera "
        "de las dos bases de indexación del intérprete.",
    )
    bloque_codigo(documento, fragmentos["I"], pie="Código del paso I.", tamano=tamano)

    marco_captura(
        documento,
        "CAPTURA 01  ·  PENDIENTE",
        "Captura del editor de PSeInt con la instrucción Dimension y las "
        "declaraciones de variables visibles en el panel de código.",
        "Sugerencia: incluye en la imagen el nombre del algoritmo en la primera "
        "línea y la instrucción FinAlgoritmo al final, para demostrar que el "
        "bloque está completo.",
    )

    titulo(documento, "4.3. Paso II. Solicitar los números y almacenarlos en el "
                      "arreglo", nivel=2)
    parrafo(
        documento,
        "Primero se solicita la cantidad de números, con un lazo "
        "Repetir-Hasta Que que no termina hasta que el dato esté entre 2 y 5. "
        "Luego, un lazo Para recorre las posiciones del arreglo y captura cada "
        "número. Cada dato se lee como cadena, se valida con la función "
        "EsNumeroValido y solo entonces se convierte a número y se guarda. Al "
        "finalizar, se recorre el arreglo para mostrar su contenido y confirmar "
        "que el almacenamiento fue correcto.",
    )
    bloque_codigo(documento, fragmentos["II"], pie="Código del paso II.", tamano=tamano)

    marco_captura(
        documento,
        "CAPTURA 02  ·  PENDIENTE",
        "Captura de la ventana de ejecución de PSeInt mostrando la solicitud de "
        "la cantidad de números y el mensaje de dato inválido cuando se digita "
        "un valor fuera del rango.",
        "Sugerencia: digita primero 9 para que aparezca el mensaje de error y "
        "luego 4, de manera que se evidencie el reintento.",
    )

    marco_captura(
        documento,
        "CAPTURA 03  ·  PENDIENTE",
        "Captura de la ventana de ejecución mostrando los cuatro números "
        "ingresados y el recorrido del arreglo numeros.",
        "Sugerencia: usa 10, 4, 2 y 8. Con esos datos la suma da 24, la resta -4, "
        "la multiplicación 640 y la división 0.15625.",
    )

    titulo(documento, "4.4. Paso III. Establecer el arreglo con las operaciones a "
                      "realizar", nivel=2)
    parrafo(
        documento,
        "Se declara el catálogo de las cuatro operaciones aritméticas en un "
        "arreglo de cadenas, y sus símbolos en un arreglo paralelo. El menú se "
        "muestra recorriendo el arreglo con un lazo, de modo que agregar una "
        "quinta operación al catálogo no obliga a modificar el código de "
        "presentación.",
    )
    bloque_codigo(documento, fragmentos["III"], pie="Código del paso III.", tamano=tamano)

    marco_captura(
        documento,
        "CAPTURA 04  ·  PENDIENTE",
        "Captura de la ventana de ejecución mostrando el catálogo de operaciones "
        "impreso a partir del arreglo nombresOperacion.",
        "Sugerencia: la imagen debe permitir leer las cuatro opciones con su "
        "respectivo símbolo.",
    )

    titulo(documento, "4.5. Paso IV. Solicitar la operación y almacenarla en el "
                      "arreglo de operaciones", nivel=2)
    parrafo(
        documento,
        "Un lazo Para recorre cuatro veces las posiciones del arreglo "
        "operaciones. En cada vuelta se imprime el catálogo y se solicita una "
        "opción, con un lazo Repetir-Hasta Que que no acepta valores fuera del "
        "rango 1 a 4. La opción se traduce a nombre de operación consultando el "
        "arreglo del paso III, y el resultado se guarda en el arreglo. Cada "
        "registro se confirma en pantalla para dejar trazabilidad del proceso.",
    )
    bloque_codigo(documento, fragmentos["IV"], pie="Código del paso IV.", tamano=tamano)

    marco_captura(
        documento,
        "CAPTURA 05  ·  PENDIENTE",
        "Captura de la ventana de ejecución durante el proceso de selección, "
        "incluido el mensaje de opción inválida ante un dato fuera del rango.",
        "Sugerencia: en la primera selección digita 7 para mostrar la "
        "validación y luego 1. Registra las cuatro operaciones en el orden 1, 2, "
        "3 y 4.",
    )

    titulo(documento, "4.6. Paso V. Iterar sobre el arreglo de operaciones con una "
                      "estructura de control", nivel=2)
    parrafo(
        documento,
        "Se utiliza el lazo Para con el índice k, que recorre las cuatro "
        "posiciones del arreglo operaciones. Además se incluye una búsqueda "
        "secuencial, con un lazo anidado, que localiza dentro del catálogo del "
        "paso III el código numérico de la operación almacenada. Esta búsqueda es "
        "la pieza que permite que el usuario elija las operaciones en cualquier "
        "orden: el algoritmo no asume la posición, busca el nombre.",
    )

    titulo(documento, "4.7. Paso VI. Condicionales dentro del lazo para ejecutar "
                      "cada operación", nivel=2)
    parrafo(
        documento,
        "Dentro del lazo se emplea la estructura de selección múltiple Segun para "
        "despachar la operación según el código localizado. Cada rama recorre el "
        "arreglo numeros con un lazo propio: la suma y la resta parten de "
        "numeros[1] y acumulan; la multiplicación parte de 1 y acumula el "
        "producto; la división se resuelve como división encadenada. Antes de "
        "dividir se verifica que ningún divisor sea cero, y en ese caso se "
        "registra un estado de error en lugar de producir un resultado incorrecto. "
        "La cláusula De Otro Modo cubre la operación no reconocida.",
    )
    bloque_codigo(documento, fragmentos["V_VI"], pie="Código de los pasos V y VI.", tamano=tamano)

    marco_captura(
        documento,
        "CAPTURA 06  ·  PENDIENTE",
        "Captura de la ventana de ejecución mostrando el mensaje «Procesando el "
        "arreglo operaciones con el lazo Para...», junto con el bloque de "
        "condicionales Segun en el panel de código.",
        "Sugerencia: una captura con el panel de código y otra con la consola "
        "dan mejor soporte al criterio de diseño de algoritmos.",
    )

    titulo(documento, "4.8. Paso VII. Mostrar una suma, una resta, una "
                      "multiplicación y una división", nivel=2)
    parrafo(
        documento,
        "Se recorre el arreglo operaciones y se presenta el resultado de cada "
        "operación. La instrucción Escribir recibe varios valores separados por "
        "coma, lo que permite mostrar el nombre de la operación y su resultado en "
        "la misma línea sin necesidad de convertir el número a texto. Se cuenta "
        "cuántas operaciones terminaron con error y se imprime un resumen, "
        "seguido del historial completo.",
    )
    bloque_codigo(
        documento,
        fragmentos["VII"],
        pie="Código del paso VII y del historial de operaciones.",
        tamano=tamano,
    )

    marco_captura(
        documento,
        "CAPTURA 07  ·  PENDIENTE",
        "Captura de la ventana de ejecución con los cuatro resultados: la suma, "
        "la resta, la multiplicación y la división.",
        "Sugerencia: esta es la evidencia central del taller. Debe verse el "
        "nombre de la operación y su resultado.",
    )

    titulo(documento, "4.9. Función auxiliar de validación", nivel=2)
    parrafo(
        documento,
        "La función EsNumeroValido recorre la cadena de entrada carácter por "
        "carácter, usando Longitud y Subcadena. Acepta dígitos, un signo negativo "
        "únicamente al inicio y como máximo un separador decimal. Devuelve "
        "Verdadero o Falso, y se invoca desde el paso II. Se escribió como "
        "función aparte para que la lógica de validación quede encapsulada y "
        "reutilizable, en lugar de repetirla dentro del lazo de lectura.",
    )

    marco_captura(
        documento,
        "CAPTURA 08  ·  PENDIENTE",
        "Captura de la ventana de ejecución mostrando el rechazo de una entrada "
        "no numérica, por ejemplo al digitar abc en lugar de un número.",
        "Sugerencia: captura también el reintento correcto posterior, que "
        "demuestra que el programa continúa y no se detiene.",
    )

    marco_captura(
        documento,
        "CAPTURA 09  ·  PENDIENTE",
        "Captura del código de la función EsNumeroValido, ubicado después de "
        "FinAlgoritmo, con la invocación EsNumeroValido(entrada) resaltada en el "
        "paso II.",
        "Sugerencia: en PSeInt las funciones se escriben después de FinAlgoritmo.",
    )


def construir_programa_completo(documento, lineas: list[str], tamano: float) -> None:
    salto_de_pagina(documento)
    titulo(documento, "5. Programa completo consolidado", nivel=1)
    parrafo(
        documento,
        "A continuación se presenta el algoritmo íntegro, tal como debe escribirse "
        "en PSeInt para poder ejecutarse de una sola vez. Los bloques de los "
        "pasos anteriores corresponden, en orden, a los fragmentos que se "
        "muestran a continuación.",
    )
    bloque_codigo(
        documento,
        [linea.rstrip() for linea in lineas],
        pie="Archivo: src/pseint/MiCalculadora.pseint",
        tamano=tamano,
    )

    marco_captura(
        documento,
        "CAPTURA 10  ·  PENDIENTE",
        "Captura de la ventana principal de PSeInt con el programa completo "
        "escrito y el botón Ejecutar resaltado.",
        "Sugerencia: antes de ejecutar, verifica que la primera línea sea "
        "Algoritmo MiCalculadora y que exista un único FinAlgoritmo seguido de la "
        "función EsNumeroValido con su FinFuncion.",
    )


EPIGRAFES_FIGURAS = {
    "01_caso_exitoso": "Ejecución completa con las cuatro operaciones "
                       "aritméticas sobre los números 10, 4, 2 y 8.",
    "02_division_entre_cero": "Manejo del error aritmético: la división entre "
                              "cero se reporta y el programa continúa.",
    "03_validacion_de_entradas": "Validación de entradas: se rechazan la "
                                 "cantidad fuera de rango, el texto no numérico "
                                 "y la opción inválida.",
    "04_orden_libre": "Selección de operaciones fuera del orden del catálogo; la "
                      "búsqueda secuencial resuelve correctamente el despacho.",
}


def construir_evidencias(documento) -> None:
    salto_de_pagina(documento)
    titulo(documento, "6. Evidencias de ejecución", nivel=1)
    parrafo(
        documento,
        "Las imágenes de esta sección son salidas reales de la ejecución de la "
        "implementación de referencia del mismo algoritmo, escrita en Python y "
        "depositada en el repositorio del proyecto. Se incluyen como verificación "
        "objetiva de la lógica antes de digitarla en PSeInt, y como respaldo de "
        "los resultados que deben aparecer también en la interfaz.",
    )
    parrafo(
        documento,
        "Las trazas de entrada utilizadas son las que figuran en el archivo de "
        "pruebas, de modo que cualquier resultado puede reproducirse de forma "
        "determinista. Los comandos de ejecución son los siguientes:",
    )
    bloque_codigo(
        documento,
        [
            "python src/python/mi_calculadora.py exito",
            "python src/python/mi_calculadora.py division_cero",
            "python src/python/mi_calculadora.py validacion",
            "python src/python/mi_calculadora.py orden_libre",
            "python -m unittest discover -s tests -v",
        ],
        pie="Comandos de ejecución y de verificación automática.",
        tamano=tamano_fuente_codigo([_CODIGO_COMANDOS]),
    )
    parrafo(
        documento,
        "Estas trazas se presentan además de las capturas de la sección 4, "
        "porque estas últimas deben tomarse de la interfaz de PSeInt por el "
        "estudiante, que es lo que exige el criterio de resolución y evidencias.",
        cursiva=True,
        tamano=10,
    )

    manifiesto = CARPETA_EVIDENCIAS / "MANIFESTO.json"
    if not manifiesto.exists():
        parrafo(
            documento,
            "Las figuras aún no se han generado. Ejecute "
            "«python evidencias/generar_evidencias.py» y vuelva a compilar el "
            "documento.",
            cursiva=True,
        )
        return

    import json

    figuras = json.loads(manifiesto.read_text(encoding="utf-8"))["figuras"]
    numero = 0
    for figura_info in figuras:
        ruta = CARPETA_EVIDENCIAS / str(figura_info["archivo"])
        if not ruta.exists():
            continue
        numero += 1
        traza = str(figura_info["traza"])
        rango = figura_info["rango"]
        sufijo = (f" (continuación, parte {rango[0]} de {rango[1]})"
                  if figura_info["partes"] > 1 else "")
        epigrafe = f"Figura {numero}. {EPIGRAFES_FIGURAS.get(traza, traza)}{sufijo}."
        figura(documento, ruta, epigrafe, ancho_cm=16.0,
               alto_max_cm=ALTO_MAXIMO_FIGURA)


def construir_pruebas(documento) -> None:
    titulo(documento, "7. Pruebas y análisis de resultados", nivel=1)
    parrafo(
        documento,
        "El algoritmo se sometió a 27 pruebas automáticas que verifican tanto la "
        "función de validación como los resultados aritméticos, el manejo del "
        "error de división entre cero, la validación de entradas y el despacho "
        "de operaciones en orden arbitrario. Todas las pruebas pasan.",
    )
    parrafo(
        documento,
        "Los resultados obtenidos con la traza de la figura 1 se contrastan a "
        "continuación con el cálculo manual, para demostrar que la lógica del "
        "algoritmo es correcta y no solamente que el programa termina sin error.",
    )
    tabla(
        documento,
        ["Operación", "Regla aplicada", "Cálculo", "Resultado del programa"],
        [
            ["Suma", "n1 + n2 + n3 + n4", "10 + 4 + 2 + 8", "24"],
            ["Resta", "n1 - n2 - n3 - n4", "10 - 4 - 2 - 8", "-4"],
            ["Multiplicación", "n1 x n2 x n3 x n4", "10 x 4 x 2 x 8", "640"],
            ["División", "n1 / n2 / n3 / n4", "10 / 4 / 2 / 8", "0.15625"],
        ],
        anchos=[3.2, 5.0, 4.4, 3.9],
    )
    parrafo(
        documento,
        "El caso de la división merece una mención aparte. Cuando uno de los "
        "divisores es cero, la figura 2 muestra que el programa no produce un "
        "resultado inválido ni se detiene: registra el estado ERROR: NO SE PUEDE "
        "DIVIDIR ENTRE CERO, cuenta la operación entre las fallidas y continúa "
        "procesando las restantes. Esta es la diferencia entre un programa "
        "validado y uno que simplemente funciona en el caso ideal.",
    )
    tabla(
        documento,
        ["Aspecto evaluado", "Resultado", "Evidencia"],
        [
            ["Cantidad de operaciones aritméticas resueltas", "4 de 4", "Figura 1"],
            ["Uso de arreglos para datos, operaciones y resultados", "6 arreglos",
             "Sección 4.2"],
            ["Validación de cantidad, números y opciones", "Correcta", "Figura 3"],
            ["Manejo de la división entre cero", "Correcto", "Figura 2"],
            ["Operaciones en orden diferente al catálogo", "Correcto", "Figura 4"],
            ["Pruebas automáticas superadas", "27 de 27",
             "tests/test_mi_calculadora.py"],
        ],
        anchos=[8.4, 3.2, 4.9],
    )


def construir_conclusiones(documento) -> None:
    titulo(documento, "8. Conclusiones", nivel=1)
    for texto in [
        "Se cumplió el objetivo de diseñar una calculadora en pseudocódigo que "
        "resuelve las cuatro operaciones aritméticas, empleando arreglos como "
        "tipo de dato estructurado en seis arreglos: dos numéricos y cuatro de "
        "cadena.",
        "Los siete pasos del enunciado quedaron implementados en el orden "
        "indicado, y el algoritmo ejecuta una suma, una resta, una multiplicación "
        "y una división, en ese orden, sobre los datos suministrados por el "
        "usuario.",
        "El uso de arreglos demostró su ventaja principal: las cuatro operaciones "
        "se resuelven en un único bloque de código dentro de un lazo. Si el "
        "número de operaciones se aumentara a diez, el cambio sería una sola "
        "línea en el catálogo y no un bloque de código nuevo.",
        "La validación de entradas resultó ser la decisión de mayor impacto en "
        "la calidad del programa. Sin ella, un único dato mal digitado "
        "interrumpiría la ejecución; con ella, el programa corrige al usuario y "
        "continúa.",
        "La declaración de los arreglos con una posición adicional de holgura "
        "resolvió la dependencia de la base de indexación configurada en el "
        "intérprete, un detalle que suele pasarse por alto y que puede hacer que "
        "el taller falle en la máquina del docente.",
        "La comprobación mediante 27 pruebas automáticas y cuatro trazas de "
        "consola permitió validar la lógica antes de digitarla en el editor, lo "
        "que redujo el número de correcciones durante el desarrollo.",
    ]:
        vineta(documento, texto)


def construir_referencias(documento) -> None:
    salto_de_pagina(documento)
    titulo(documento, "9. Referencias", nivel=1)

    parrafo(documento, "Recursos básicos de la semana", negrita=True)
    for texto in [
        "Algar, M. y Fernández de Sevilla, M. (2019). Primeros tipos "
        "estructurados. En Introducción práctica a la programación con Python "
        "(pp. 205-272). Editorial Universidad de Alcalá.",
        "Joyanes, L. (2020). Las cadenas de caracteres. En Fundamentos de "
        "programación: algoritmos, estructuras de datos y objetos (pp. 293-316). "
        "McGraw-Hill.",
    ]:
        referencia(documento, texto)

    parrafo(documento, "Recurso complementario", negrita=True)
    referencia(
        documento,
        "Romano, F. (2018). Iterating and making decisions: conditional "
        "programming. En Learn Python Programming: The no-nonsense, beginner's "
        "guide to programming, data science, and web development with Python 3.7 "
        "(pp. 79-105). Packt Publishing.",
    )

    parrafo(documento, "Fuentes adicionales consultadas", negrita=True)
    for texto in [
        "Liskov, B. (2008). Data in a program should be structured like the data "
        "it represents. En B. Lipman y J. Sneakers (Coords.), Program synthesis "
        "(pp. 397-411). Addison-Wesley.",
        "McConnell, S. (2004). Code Complete: A comprehensive handbook of "
        "software construction (2.ª ed.). Microsoft Press.",
        "Pressman, R. S. y Maxim, B. R. (2014). Software Engineering: A "
        "practitioner's approach (8.ª ed.). McGraw-Hill.",
        "Python Software Foundation. (2024). Python documentation: Data "
        "structures. https://docs.python.org/3/tutorial/datastructures.html",
        "PSeInt. (s. f.). PSeInt: editor de algoritmos en pseudocódigo. "
        "http://pseint.sourceforge.net/",
    ]:
        referencia(documento, texto)

    parrafo(documento, "Materiales del proyecto", negrita=True)
    tabla(
        documento,
        ["Archivo", "Descripción"],
        [
            ["src/pseint/MiCalculadora.pseint", "Algoritmo en pseudocódigo (PSeInt)"],
            ["src/python/mi_calculadora.py", "Implementación de referencia en Python"],
            ["tests/test_mi_calculadora.py", "27 pruebas automáticas"],
            ["evidencias/generar_evidencias.py", "Generador de las figuras de consola"],
            ["evidencias/consola/", "Imágenes de salida de ejecución"],
            ["evidencias/paso_a_paso/CUADRILLA_DE_CAPTURAS.md",
             "Guion de captura paso a paso"],
        ],
        anchos=[7.4, 9.1],
    )
    parrafo(documento, f"Repositorio: {REPOSITORIO}", cursiva=True, tamano=10)


def main() -> None:
    lineas, fragmentos = cargar_fuente()
    verificar_cobertura(lineas, fragmentos)

    # Un solo cuerpo de fuente para todo el código del documento: el más
    # grande con el que quepa la línea más larga sin partirse.
    bloques = [[l.rstrip() for l in lineas]] + list(fragmentos.values())
    tamano = tamano_fuente_codigo(bloques)
    columnas = ancho_maximo_codigo(bloques)
    print(f"Código: línea más larga de {columnas} columnas a {tamano} pt")

    documento = Document()
    configurar_pagina(documento)

    construir_portada(documento)
    construir_indice(documento)
    construir_introduccion(documento)
    construir_marco_conceptual(documento)
    construir_preguntas(documento)
    construir_desarrollo(documento, fragmentos, tamano)
    construir_programa_completo(documento, [l.rstrip() for l in lineas], tamano)
    construir_evidencias(documento)
    construir_pruebas(documento)
    construir_conclusiones(documento)
    construir_referencias(documento)

    documento.save(ARCHIVO_SALIDA)

    palabras = len(re.findall(r"\w+", "\n".join(p.text for p in documento.paragraphs)))
    print(f"Documento generado: {ARCHIVO_SALIDA}")
    print(f"  Párrafos .......: {len(documento.paragraphs)}")
    print(f"  Tablas .........: {len(documento.tables)}")
    print(f"  Imágenes .......: {len(documento.inline_shapes)}")
    print(f"  Palabras aprox..: {palabras}")


if __name__ == "__main__":
    main()