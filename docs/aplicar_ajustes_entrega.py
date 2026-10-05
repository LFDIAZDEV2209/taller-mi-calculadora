"""
Ajusta el documento de entrega antes de subirlo a la plataforma.

El estudiante ya insertó sus capturas de pantalla de la interfaz de PSeInt
en el documento, por lo que este script NO regenera el .docx: opera
directamente sobre el archivo existente para no perder ese trabajo.

Hace tres cosas:

  1. Extrae el programa en pseudocódigo que el estudiante digito en el
     documento y lo guarda como src/pseint/MiCalculadora.pseint, para que
     el repositorio contenga exactamente el algoritmo documentado.
  2. Retira las figuras de consola generadas a partir de la implementación
     de referencia en Python, dejando en el documento únicamente lo que la
     consigna pide, y reoriente el capítulo 6 hacia el repositorio.
  3. Reajusta las referencias del capítulo 7 que apuntaban a esas figuras.

Uso:
    python docs/aplicar_ajustes_entrega.py
"""

from __future__ import annotations

import hashlib
import re
import shutil
import subprocess
import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.table import Table
from docx.text.paragraph import Paragraph

RAIZ = Path(__file__).resolve().parents[1]
DOCUMENTO = RAIZ / "docs" / "diaz_luis_taller.docx"
CONSOLA = RAIZ / "evidencias" / "consola"
ARCHIVO_PSEINT = RAIZ / "src" / "pseint" / "MiCalculadora.pseint"
RESPALDO = RAIZ / "evidencias" / "revision" / "respaldo_con_capturas.docx"

REPOSITORIO = "https://github.com/LFDIAZDEV2209/taller-mi-calculadora"

TITULO_CAPITULO_6 = "6. Material adicional en el repositorio"

TEXTO_CAPITULO_6 = (
    "Las capturas de pantalla de la interfaz de PSeInt que sustentan los siete "
    "pasos del taller se presentan en el capítulo 4, acompañando cada bloque de "
    "pseudocódigo, tal como lo exige la consigna. Para que el documento contenga "
    "únicamente lo solicitado, el material complementario de verificación se "
    f"depositó en el repositorio del proyecto: {REPOSITORIO}"
)

TEXTO_POSTERIOR_COMANDOS = (
    "El archivo tests/validar_pseint.py realiza además una verificación "
    "estructural del pseudocódigo por análisis estático: confirma que cada "
    "bloque abierto tenga su cierre, que todo arreglo esté declarado antes de "
    "usarse, que los índices permanezcan dentro del rango declarado y que toda "
    "variable leída haya sido declarada."
)

REEMPLAZOS_CAPITULO_7 = [
    (
        "Los resultados obtenidos con la traza de la figura 1 se contrastan a "
        "continuación con el cálculo manual, para demostrar que la lógica del "
        "algoritmo es correcta y no solamente que el programa termina sin error.",
        "Los resultados obtenidos en la captura del paso VII, en la sección 4.8, "
        "se contrastan a continuación con el cálculo manual, para demostrar que "
        "la lógica del algoritmo es correcta y no solamente que el programa "
        "termina sin error.",
    ),
    (
        "El caso de la división merece una mención aparte. Cuando uno de los "
        "divisores es cero, la figura 2 muestra que el programa no produce un "
        "resultado inválido ni se detiene: registra el estado ERROR: NO SE PUEDE "
        "DIVIDIR ENTRE CERO, cuenta la operación entre las fallidas y continúa "
        "procesando las restantes. Esta es la diferencia entre un programa "
        "validado y uno que simplemente funciona en el caso ideal.",
        "El caso de la división merece una mención aparte. Cuando uno de los "
        "divisores es cero, el programa no produce un resultado inválido ni se "
        "detiene: registra el estado ERROR: NO SE PUEDE DIVIDIR ENTRE CERO, "
        "cuenta la operación entre las fallidas y continúa procesando las "
        "restantes. El comportamiento queda documentado en la traza "
        "division_cero del repositorio. Esta es la diferencia entre un programa "
        "validado y uno que simplemente funciona en el caso ideal.",
    ),
]

# Referencias de la tabla «Aspecto evaluado / Resultado / Evidencia»
REEMPLAZOS_COLUMNA_EVIDENCIA = {
    "Figura 1": "Captura de la sección 4.8",
    "Figura 2": "Traza division_cero del repositorio",
    "Figura 3": "Capturas de las secciones 4.3, 4.5 y 4.9",
    "Figura 4": "Traza orden_libre del repositorio",
}

# La captura del paso VII se tomó con los números 1, 2, 3 y 4. La tabla de
# contrastación del capítulo 7 debe citar esa misma sesión para que el
# documento y la evidencia coincidan.
REEMPLAZOS_DE_CALCULO = {
    "10 + 4 + 2 + 8": "1 + 2 + 3 + 4",
    "10 - 4 - 2 - 8": "1 - 2 - 3 - 4",
    "10 x 4 x 2 x 8": "1 x 2 x 3 x 4",
    "10 / 4 / 2 / 8": "1 / 2 / 3 / 4",
}

REEMPLAZOS_DE_RESULTADO = {
    "24": "10",
    "-4": "-8",
    "640": "24",
    "0.15625": "0.04166666667",
}

# Marcas del pseudocódigo digitado por el estudiante para extraer el
# fragmento de cada paso del taller.
MARCAS_DE_PASO: dict[str, tuple[str, str | None]] = {
    "I": ("PASO I - DECLARACION", "PASO II - SOLICITAR"),
    "II": ("PASO II - SOLICITAR", "PASO III - ARREGLO"),
    "III": ("PASO III - ARREGLO", "PASO IV - SELECCIONAR"),
    "IV": ("PASO IV - SELECCIONAR", "PASO V - RECORRER"),
    "V_VI": ("PASO V - RECORRER", "PASO VII - MOSTRAR"),
    "VII": ("PASO VII - MOSTRAR", "FinAlgoritmo"),
}

CAPTIONS_DE_PASO = {
    "I": "Código del paso I.",
    "II": "Código del paso II.",
    "III": "Código del paso III.",
    "IV": "Código del paso IV.",
    "V_VI": "Código de los pasos V y VI.",
    "VII": "Código del paso VII y del historial de operaciones.",
}

FUENTE_MONOESPACIADA = "Consolas"
ANCHO_UTIL_CM = 16.2

def re_busca_codigo(texto: str) -> bool:
    """Detecta si una celda contiene pseudocódigo de PSeInt."""
    primeras = texto.strip().splitlines()
    if not primeras:
        return False
    muestra = chr(10).join(primeras[:8])
    marcadores = (
        r"\b(Algoritmo|Dimension|Escribir|Leer|Definir|Repetir"
        r"|Mientras|Segun|FinSi|FinPara|Hasta|Entonces)\b"
    )
    if re.search(marcadores, muestra):
        return True
    # Los banners de comentario encabezan varios bloques sin palabras clave
    return muestra.lstrip().startswith("//")


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------


def iterar_cuerpo(documento):
    """Recorre párrafos y tablas del cuerpo en el orden real del documento."""
    for hijo in documento.element.body.iterchildren():
        etiqueta = hijo.tag.split("}")[-1]
        if etiqueta == "p":
            yield Paragraph(hijo, documento)
        elif etiqueta == "tbl":
            yield Table(hijo, documento)


def rids_de_imagen(elemento) -> list[str]:
    return [
        blip.get(qn("r:embed"))
        for blip in elemento.iter(qn("a:blip"))
        if blip.get(qn("r:embed"))
    ]


def nombres_de_imagen(documento) -> dict[str, str]:
    """Devuelve {rid: nombre del archivo} de las partes de imagen."""
    return {
        rid: str(rel.target_part.partname)
        for rid, rel in documento.part.rels.items()
        if "image" in rel.reltype
    }


# ---------------------------------------------------------------------------
# 1. Extracción del pseudocódigo digitado por el estudiante
# ---------------------------------------------------------------------------


def extraer_pseudocodigo(documento) -> str:
    """Localiza el bloque del programa completo y lo devuelve como texto."""
    for tabla in iterar_cuerpo(documento):
        if not isinstance(tabla, Table):
            continue
        texto = tabla.rows[0].cells[0].text
        if "Algoritmo" not in texto or "FinFuncion" not in texto:
            continue

        lineas = []
        for parrafo in tabla.rows[0].cells[0].paragraphs:
            lineas.append(parrafo.text.rstrip())
        return "\n".join(lineas)

    raise ValueError("No se encontró el bloque del programa completo en el documento.")


def escribir_pseudocodigo(lineas: list[str]) -> None:
    # Se eliminan líneas vacías consecutivas y se recortan los extremos
    limpio: list[str] = []
    for linea in lineas:
        if not linea.strip() and (not limpio or not limpio[-1].strip()):
            continue
        limpio.append(linea.rstrip())
    while limpio and not limpio[0].strip():
        limpio.pop(0)
    while limpio and not limpio[-1].strip():
        limpio.pop()

    ARCHIVO_PSEINT.write_text("\n".join(limpio) + "\n", encoding="utf-8")


# ---------------------------------------------------------------------------
# 4. Sincronización de los fragmentos del capítulo 4 con el programa digitado
# ---------------------------------------------------------------------------


def expandir_tabulaciones(linea: str, sangria: int = 4) -> str:
    sangria_inicial = len(linea) - len(linea.lstrip("\t"))
    return " " * (sangria_inicial * sangria) + linea.lstrip("\t")


def fragmento(lineas: list[str], desde: str, hasta: str | None) -> list[str]:
    i = next(i for i, linea in enumerate(lineas) if desde in linea)

    # Se incluye el banner de comentario que encabeza el bloque
    while i > 0 and lineas[i - 1].strip().startswith("//"):
        i -= 1

    if hasta is None:
        k = len(lineas)
    else:
        k = next(
            j for j, linea in enumerate(lineas) if j > i and hasta in linea
        )
        while k > i and lineas[k - 1].strip().startswith("//"):
            k -= 1
    bloque = lineas[i:k]
    while bloque and not bloque[0].strip():
        bloque.pop(0)
    while bloque and not bloque[-1].strip():
        bloque.pop()
    return bloque


def fragmentos_de_pasos(lineas: list[str]) -> dict[str, list[str]]:
    salida = {}
    for paso, (desde, hasta) in MARCAS_DE_PASO.items():
        try:
            salida[paso] = fragmento(lineas, desde, hasta)
        except StopIteration:
            raise ValueError(f"No se encontró la marca del paso {paso}: {desde}")
    return salida


def cuerpo_de_fuente(bloques: list[list[str]]) -> float:
    columnas = max(
        (len(expandir_tabulaciones(linea)) for bloque in bloques for linea in bloque),
        default=0,
    )
    for tamano in (9.0, 8.5, 8.0, 7.5, 7.0, 6.5):
        if columnas * tamano * 0.55 / 28.35 <= ANCHO_UTIL_CM:
            return tamano
    return 6.5


def _poner_texto_codigo(celda, lineas: list[str], tamano: float) -> None:
    """Reemplaza el contenido de una celda por líneas de código monoespaciado."""
    # Se conservan las propiedades de la celda (sombreado y borde)
    for parrafo in list(celda.paragraphs):
        parrafo._p.getparent().remove(parrafo._p)

    for linea in lineas:
        p = celda.add_paragraph()
        formato = p.paragraph_format
        formato.space_before = Pt(0)
        formato.space_after = Pt(0)
        formato.line_spacing = 1.0
        corrida = p.add_run(expandir_tabulaciones(linea) if linea.strip() else " ")
        corrida.font.name = FUENTE_MONOESPACIADA
        corrida.font.size = Pt(tamano)
        corrida._element.rPr.rFonts.set(qn("w:eastAsia"), FUENTE_MONOESPACIADA)


def sincronizar_bloques_de_paso(documento, lineas_pseint: list[str]) -> int:
    """Sustituye el código de los pasos I a VII por el digitado realmente."""
    fragmentos = fragmentos_de_pasos(lineas_pseint)
    tamano = cuerpo_de_fuente(list(fragmentos.values()))

    # Se relaciona cada bloque de código con el epígrafe que lo sigue
    por_paso: dict[str, object] = {}
    pendiente = None
    for elemento in iterar_cuerpo(documento):
        if isinstance(elemento, Table):
            if len(elemento.rows) == 1 and len(elemento.columns) == 1:
                texto = elemento.rows[0].cells[0].text
                if re_busca_codigo(texto):
                    pendiente = elemento
        elif isinstance(elemento, Paragraph):
            texto = elemento.text.strip()
            for paso, caption in CAPTIONS_DE_PASO.items():
                if texto == caption and pendiente is not None:
                    por_paso[paso] = pendiente
                    pendiente = None

    if len(por_paso) != len(MARCAS_DE_PASO):
        faltan = set(MARCAS_DE_PASO) - set(por_paso)
        raise ValueError(f"No se pudieron ubicar los bloques de los pasos: {faltan}")

    for paso, tabla in por_paso.items():
        _poner_texto_codigo(tabla.rows[0].cells[0], fragmentos[paso], tamano)

    return len(por_paso)


# ---------------------------------------------------------------------------
# 2. Retiro de las figuras de consola
# ---------------------------------------------------------------------------


def retirar_figuras_de_consola(documento) -> int:
    """Elimina los párrafos que contienen las figuras y sus epígrafes."""
    hashes_mios = {
        hashlib.md5(archivo.read_bytes()).hexdigest()
        for archivo in CONSOLA.glob("*.png")
    }
    consola = {
        rid
        for rid, rel in documento.part.rels.items()
        if "image" in rel.reltype
        and hashlib.md5(rel.target_part.blob).hexdigest() in hashes_mios
    }
    if not consola:
        return 0

    a_retirar = []
    for elemento in iterar_cuerpo(documento):
        if not isinstance(elemento, Paragraph):
            continue
        texto = elemento.text.strip()
        es_figura = bool(set(rids_de_imagen(elemento._p)) & consola)
        es_epigrafe = texto.startswith("Figura ")
        if es_figura or es_epigrafe:
            a_retirar.append(elemento)

    for elemento in a_retirar:
        elemento._p.getparent().remove(elemento._p)

    return len(a_retirar)


# ---------------------------------------------------------------------------
# 3. Reorientación del capítulo 6 y del capítulo 7
# ---------------------------------------------------------------------------


def reescribir_parrafo(parrafo, texto: str) -> None:
    for corrida in parrafo.runs:
        corrida._element.getparent().remove(corrida._element)
    corrida = parrafo.add_run(texto)
    corrida.font.name = "Calibri"
    corrida.font.size = parrafo.runs[0].font.size if False else None


def reorientar_capitulos(documento) -> None:
    parrafos = list(iterar_cuerpo(documento))

    # Indice y encabezado del capítulo 6: ambas apariciones se actualizan
    for elemento in parrafos:
        if isinstance(elemento, Paragraph) and elemento.text.strip() == "6. Evidencias de ejecución":
            reescribir_parrafo(elemento, TITULO_CAPITULO_6)

    en_capitulo_6 = False
    retirar: list[Paragraph] = []
    for elemento in parrafos:
        if not isinstance(elemento, Paragraph):
            continue
        texto = elemento.text.strip()
        if texto.startswith("6. "):
            en_capitulo_6 = True
            continue
        if texto.startswith("7. "):
            en_capitulo_6 = False

        if not en_capitulo_6:
            continue

        if texto.startswith("6. "):
            pass
        elif texto.startswith("Las imágenes de esta sección"):
            reescribir_parrafo(elemento, TEXTO_CAPITULO_6)
        elif texto.startswith("Las trazas de entrada utilizadas"):
            reescribir_parrafo(
                elemento,
                "Los comandos de ejecución y de verificación son los siguientes.",
            )
        elif texto.startswith("Estas trazas se presentan además"):
            retirar.append(elemento)
        elif texto.startswith("El archivo tests/validar_pseint.py"):
            reescribir_parrafo(elemento, TEXTO_POSTERIOR_COMANDOS)

    # Título del programa completo consolidado se ajusta al repositorio
    for elemento in parrafos:
        if isinstance(elemento, Paragraph) and elemento.text.strip().startswith("A continuación se presenta el algoritmo íntegro"):
            reescribir_parrafo(
                elemento,
                "A continuación se presenta el algoritmo íntegro, tal como fue "
                "digitado y ejecutado en PSeInt. Los bloques de los pasos "
                "anteriores corresponden, en orden, a los fragmentos de este "
                "listado. El archivo fuente está en el repositorio del proyecto.",
            )
            break

    for elemento in retirar:
        elemento._p.getparent().remove(elemento._p)


def reajustar_tabla_de_contrastacion(documento) -> int:
    """Alinea la tabla del capítulo 7 con la sesión efectivamente capturada."""
    for tabla in iterar_cuerpo(documento):
        if not isinstance(tabla, Table) or len(tabla.rows) < 2:
            continue
        encabezados = [c.text.strip() for c in tabla.rows[0].cells]
        if encabezados[:2] != ["Operación", "Regla aplicada"]:
            continue

        cambios = 0
        for fila in tabla.rows[1:]:
            celda_calculo = fila.cells[2]
            reemplazo = REEMPLAZOS_DE_CALCULO.get(celda_calculo.text.strip())
            if reemplazo:
                _poner_texto_simple(celda_calculo, reemplazo)
                cambios += 1
            celda_resultado = fila.cells[3]
            reemplazo = REEMPLAZOS_DE_RESULTADO.get(celda_resultado.text.strip())
            if reemplazo:
                _poner_texto_simple(celda_resultado, reemplazo)
                cambios += 1
        return cambios
    return 0


def mover_captura_del_paso_v(documento) -> bool:
    """Ubica la captura del recorrido en la sección del paso V.

    La imagen muestra el recorrido del arreglo operaciones, que es el paso V,
    pero estaba colocada en la sección del paso VI.
    """
    objetivo = None
    heading_paso_vi = None
    for elemento in iterar_cuerpo(documento):
        if not isinstance(elemento, Paragraph):
            continue
        texto = elemento.text.strip()
        if texto.startswith("4.6. Paso V."):
            objetivo = elemento
        elif texto.startswith("4.7. Paso VI."):
            heading_paso_vi = elemento
            break

    if objetivo is None or heading_paso_vi is None:
        return False

    # La captura es la tabla que sigue al epígrafe de los pasos V y VI
    epigrafe = None
    for elemento in iterar_cuerpo(documento):
        if isinstance(elemento, Table):
            continue
        if isinstance(elemento, Paragraph) and elemento.text.strip() == "Código de los pasos V y VI.":
            epigrafe = elemento
            break

    if epigrafe is None:
        return False

    # Se recorre hacia adelante hasta encontrar la tabla con imagen
    hermanos = list(documento.element.body.iterchildren())
    indice = hermanos.index(epigrafe._p)
    captura = None
    for candidato in hermanos[indice + 1:]:
        if rids_de_imagen(candidato):
            captura = candidato
            break
        if candidato.tag.split("}")[-1] == "p" and candidato.text and candidato.text.strip():
            # Se detiene ante el primer párrafo con texto (nueva sección)
            if Paragraph(candidato, documento).text.strip().startswith("4.8."):
                break

    if captura is None:
        return False

    heading_paso_vi._p.addprevious(captura)
    return True


def retirar_imagenes_huerfanas(documento) -> int:
    """Elimina las partes de imagen que ya no referencia el documento.

    Al retirar los párrafos de las figuras, python-docx conserva las partes
    binarias en el paquete porque siguen ligadas por su relación, y el
    archivo resultante pesa bastante más de lo necesario.
    """
    body = documento.element.body
    referenciados = {
        nodo.get(qn("r:embed"))
        for nodo in body.iter()
        if nodo.get(qn("r:embed"))
    }
    referenciados |= {
        nodo.get(qn("r:id"))
        for nodo in body.iter()
        if nodo.get(qn("r:id"))
    }

    retiradas = 0
    for rid, rel in list(documento.part.rels.items()):
        if "image" not in rel.reltype:
            continue
        if rid in referenciados:
            continue
        del documento.part.rels[rid]
        retiradas += 1
    return retiradas


def _poner_texto_simple(celda, texto: str) -> None:
    for parrafo in celda.paragraphs:
        for corrida in parrafo.runs:
            corrida._element.getparent().remove(corrida._element)
        parrafo.add_run(texto)
        return


def reajustar_capitulo_7(documento) -> None:
    parrafos = list(iterar_cuerpo(documento))

    for elemento in parrafos:
        if not isinstance(elemento, Paragraph):
            continue
        texto = elemento.text.strip()
        for anterior, nuevo in REEMPLAZOS_CAPITULO_7:
            if texto.startswith(anterior[:60]):
                reescribir_parrafo(elemento, nuevo)
                break

    for elemento in parrafos:
        if not isinstance(elemento, Table):
            continue
        if len(elemento.columns) < 3 or len(elemento.rows) < 2:
            continue
        encabezados = [c.text.strip() for c in elemento.rows[0].cells]
        if encabezados[:2] != ["Aspecto evaluado", "Resultado"]:
            continue
        for fila in elemento.rows[1:]:
            celda = fila.cells[2]
            reemplazo = REEMPLAZOS_COLUMNA_EVIDENCIA.get(celda.text.strip())
            if reemplazo:
                for corrida in celda.paragraphs[0].runs:
                    corrida._element.getparent().remove(corrida._element)
                celda.paragraphs[0].add_run(reemplazo)


# ---------------------------------------------------------------------------


def main() -> int:
    RESPALDO.parent.mkdir(parents=True, exist_ok=True)
    if not RESPALDO.exists():
        shutil.copy2(DOCUMENTO, RESPALDO)
        print(f"Respaldo del documento con capturas: {RESPALDO.relative_to(RAIZ)}")

    documento = Document(str(DOCUMENTO))

    # 1. Se extrae el pseudocódigo digitado por el estudiante
    codigo = extraer_pseudocodigo(documento)
    escribir_pseudocodigo(codigo.splitlines())
    lineas_pseint = ARCHIVO_PSEINT.read_text(encoding="utf-8").splitlines()
    print(f"Pseudocódigo extraído: {ARCHIVO_PSEINT.relative_to(RAIZ)}")

    # 2. Los fragmentos del capítulo 4 se alinean con ese programa
    sincronizados = sincronizar_bloques_de_paso(documento, lineas_pseint)
    print(f"Bloques de los pasos sincronizados: {sincronizados}")

    # 3. Se retiran las figuras de consola
    retiradas = retirar_figuras_de_consola(documento)
    print(f"Figuras de consola retiradas del documento: {retiradas}")

    # 3. Se reorientan los capítulos 6 y 7
    reorientar_capitulos(documento)
    reajustar_capitulo_7(documento)
    print("Capítulo 6 reorientado al repositorio y capítulo 7 reajustado.")

    cambios = reajustar_tabla_de_contrastacion(documento)
    print(f"Celdas de la tabla de contrastación ajustadas: {cambios}")

    if mover_captura_del_paso_v(documento):
        print("Captura del recorrido del arreglo movida a la sección del paso V.")
    else:
        print("La captura del recorrido ya estaba en la sección correcta.")

    # 4. Se depuran las partes de imagen que quedaron sin referencia
    huerfanas = retirar_imagenes_huerfanas(documento)
    print(f"Partes de imagen sin referencia retiradas: {huerfanas}")

    documento.save(str(DOCUMENTO))
    print(f"Documento actualizado: {DOCUMENTO.relative_to(RAIZ)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())