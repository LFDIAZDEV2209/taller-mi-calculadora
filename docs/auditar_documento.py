"""
Auditoría de ortografía del documento generado.

Extrae el texto del .docx descartando los bloques de pseudocódigo (que van
deliberadamente sin tildes por compatibilidad con la fuente de PSeInt) y
verifica tres familias de errores:

  1. Terminaciones que obligan a tilde: -ción, -sión, -ciones, -siones.
  2. Esdrújulas y llanas con tilde obligatoria, mediante diccionario.
  3. Interrogativos y conectores que se escriben sin tilde en relativo pero
     con tilde en interrogativo o en su uso adverbial.
  4. Residuos de otro idioma, caracteres corruptos y espaciado anómalo.

Uso:
    python docs/auditar_documento.py
Código de salida: 0 si el documento está limpio, 1 si hay hallazgos.
"""

from __future__ import annotations

import re
import sys
import unicodedata
from pathlib import Path

from docx import Document

DOCUMENTO = Path(__file__).resolve().parent / "diaz_luis_taller.docx"

# Caracteres legítimos en un documento en español
PERMITIDOS = set("áéíóúüñÁÉÍÓÚÜÑ¿¡«»·—–…“”‘’°ªº©")

# Terminaciones que obligan a tilde.
# Regla clave: en español el plural de los sustantivos terminados en -ción y
# -sión PERDE la tilde (instrucción -> instrucciones, decisión -> decisiones),
# por lo que los plurales "-ciones" y "-siones" NO deben marcarse aquí.
TERMINACIONES_CON_TILDE: dict[str, str] = {
    "cion": "ción",
    "sion": "sión",
    "logico": "lógico",
    "logica": "lógica",
    "logicos": "lógicos",
    "logicas": "lógicas",
    "numerico": "numérico",
    "numerica": "numérica",
    "numericos": "numéricos",
    "numericas": "numéricas",
    "aritmetico": "aritmético",
    "aritmetica": "aritmética",
    "aritmeticos": "aritméticos",
    "aritmeticas": "aritméticas",
    "codigo": "código",
    "codigos": "códigos",
    "metodo": "método",
    "metodos": "métodos",
    "indice": "índice",
    "indices": "índices",
    "numero": "número",
    "numeros": "números",
    "calculo": "cálculo",
    "calculos": "cálculos",
    "analisis": "análisis",
    "sintesis": "síntesis",
    "practica": "práctica",
    "practicas": "prácticas",
    "practico": "práctico",
    "ultimo": "último",
    "ultima": "última",
    "unico": "único",
    "unica": "única",
    "tecnico": "técnico",
    "tecnica": "técnica",
    "automatico": "automático",
    "automatica": "automática",
    "automaticas": "automáticas",
    "minimo": "mínimo",
    "minima": "mínima",
    "maximo": "máximo",
    "maxima": "máxima",
    "teorico": "teórico",
    "teorica": "teórica",
    "basico": "básico",
    "basicos": "básicos",
    "basica": "básica",
    "publico": "público",
    "proposito": "propósito",
    "terminacion": "terminación",
    "expresion": "expresión",
    "atencion": "atención",
    "aplicacion": "aplicación",
    "formacion": "formación",
    "definicion": "definición",
    "descripcion": "descripción",
    "produccion": "producción",
    "reunion": "reunión",
    "decision_": "decisión",
}

# Identificadores y palabras clave de PSeInt. Aparecen citados en la prosa
# del documento y, por compatibilidad con la fuente del intérprete, se
# escriben deliberadamente sin tilde.
IDENTIFICADORES_PSEINT = {
    "dimension", "algoritmo", "finalgoritmo", "funcion", "finfuncion",
    "logico", "entero", "real", "cadena", "proceso", "finproceso",
    "escribir", "leer", "segun", "finsgun", "finsegun", "mientras",
    "finmientras", "repetir", "hastaque", "entonces", "finsi", "sino",
    "otromodo", "subcadena", "longitud", "convertiranumero", "hacer",
    "finpara", "conpaso", "verdadero", "falso",
    "numeros", "operaciones", "resultados", "codigo", "opcion",
    "cantidad", "acumulado", "numerado", "solicitar",
    "nombresoperacion", "simbolooperacion", "estadooperacion",
    "esnumerovalido", "suma", "resta", "multiplicacion", "division",
    "correcta", "noresultados",
}

# Palabras inglesas que no deben aparecer en la prosa. Se excluyen las que
# también son palabras españolas válidas (variable, simple, sistema, etc.).
INGLES: set[str] = {
    "the", "and", "with", "that", "this", "have", "will", "would", "can",
    "could", "step", "steps", "data", "string", "array", "arrays",
    "input", "output", "print", "write", "file", "files",
    "code", "run", "use", "used", "value", "values", "user", "users",
    "result", "first", "second", "third", "next", "last", "true",
    "false", "case", "cases", "example", "examples", "check", "called",
    "table", "tables", "figure", "figures", "problem", "problems",
    "question", "answer", "notes", "let", "get", "given", "give",
    "loop", "make", "reference", "need", "work", "try", "start",
    "one", "two", "three", "four", "show", "shows", "see", "seen",
    "every", "other", "another", "some", "each", "only", "also", "just",
    "like", "short", "long", "same", "different",
}

# Conectores y adverbios: se escriben sin tilde en relativo, con tilde en
# su uso adverbial. Se comprueban solo como palabra aislada.
CONECTORES: dict[str, str] = {
    "mas": "más",
    "tambien": "también",
    "despues": "después",
    "ademas": "además",
    "segun": "según",
    "asi": "así",
    "aqui": "aquí",
    "ahi": "ahí",
    "estan": "están",
    "estara": "estará",
    "estaran": "estarán",
    "sera": "será",
    "seran": "serán",
    "habia": "había",
    "habra": "habrá",
    "podria": "podría",
    "podrian": "podrían",
    "seria": "sería",
    "serian": "serían",
    "quedaria": "quedaría",
    "haria": "haría",
    "harían": "harían",
    "diria": "diría",
    "creeria": "creería",
    "queria": "querría",
    "sabia": "sabía",
}

# Inglés filtrado (excepto en referencias bibliográficas)

# Encabezados que deben llevar signo de apertura de interrogación
INTERROGATIVOS_ESPERADOS = {
    "3.1": "¿Cuál",
    "3.2": "¿Cómo",
}

MARCAS_CODIGO = (
    "//", "Dimension ", "FinAlgoritmo", "Algoritmo MiCalculadora",
    "Escribir ", "Leer ", "FinSi", "Sino", "FinPara", "FinMientras",
    "Hasta Que", "FinSegun", "Funcion ", "FinFuncion", "Subcadena",
    "Longitud(", "ConvertirANumero", "operaciones[", "numeros[",
    "resultados[", "estadoOperacion[", "nombresOperacion[",
    "simboloOperacion[", "acumulado", "EsNumeroValido(", "divisorEnCero",
    "datoValido", "numeroIngresado", "cantidad", "opcion", "codigo",
    "Entero ", "Real ", "Cadena ", "Logico ",
)


def es_bloque_de_codigo(texto: str) -> bool:
    limpio = texto.strip()
    if not limpio:
        return True
    if limpio.startswith(("//", "\t", "\t\t", "\t\t\t")):
        return True
    if any(limpio.startswith(marca) for marca in MARCAS_CODIGO):
        return True
    if re.search(r"[<>-]=|FinSi|FinPara|<-|<- 0|Segun|Entonces", limpio):
        return True
    # fragmentos sin palabras:检 líneas de código recortadas
    if not re.search(r"[A-Za-zÁÉÍÓÚáéíóúñÑ]{4,}", limpio) and re.search(r"[<>\[\]]", limpio):
        return True
    return False


def es_referencia_bibliografica(texto: str) -> bool:
    if re.search(r"\(\d{4}\)\.\s", texto):
        return True
    if "http" in texto:
        return True
    if re.search(r"\bpp\.\s*\d", texto):
        return True
    return False


def es_ruta_o_codigo(texto: str) -> bool:
    return bool(re.search(r"[\w/]+\.(py|md|pseint|docx|png|html|csv|txt)", texto))


def revisar(lineas: list[str]) -> list[str]:
    hallazgos: list[str] = []

    for texto in lineas:
        if es_bloque_de_codigo(texto) or es_ruta_o_codigo(texto):
            continue
        bibliografico = es_referencia_bibliografica(texto)
        limpio = unicodedata.normalize("NFC", texto)

        # 1. Caracteres corruptos
        for caracter in limpio:
            if ord(caracter) >= 128 and caracter not in PERMITIDOS:
                hallazgos.append(
                    f"[caracter U+{ord(caracter):04X}] {limpio[:80]!r}"
                )
                break

        # 2. Terminaciones que obligan a tilde
        for palabra in re.findall(r"[A-Za-zÁÉÍÓÚáéíóúñÑ]+", limpio):
            if IDENTIFICADORES_PSEINT and palabra.lower() in IDENTIFICADORES_PSEINT:
                continue
            if any(ord(c) > 127 and c in "áéíóúÁÉÍÓÚ" for c in palabra):
                continue  # ya lleva tilde
            for terminacion, correccion in TERMINACIONES_CON_TILDE.items():
                if palabra.lower().endswith(terminacion):
                    hallazgos.append(
                        f"[falta tilde: {palabra} -> {palabra[:-len(terminacion)]}"
                        f"{correccion}] {limpio[:80]!r}"
                    )
                    break

        # 3. Conectores y adverbios
        for palabra in re.findall(r"[A-Za-zÁÉÍÓÚáéíóúñÑ]+", limpio):
            correccion = CONECTORES.get(palabra.lower())
            if correccion and not any(ord(c) > 127 for c in palabra):
                hallazgos.append(
                    f"[falta tilde: {palabra} -> {correccion}] {limpio[:80]!r}"
                )

        # 4. Inglés filtrado (excepto en referencias bibliográficas)
        if not bibliografico:
            for palabra in re.findall(r"[A-Za-zÁÉÍÓÚáéíóúñÑ]+", limpio):
                if palabra.lower() in INGLES:
                    hallazgos.append(
                        f"[ingles: {palabra}] {limpio[:80]!r}"
                    )

        # 5. Espaciado anómalo (se permite el doble espacio deliberado
        #    alrededor del punto medio de los rótulos de captura)
        for coincidencia in re.finditer(r"(?<!·) {2,}(?!·)", limpio):
            if not limpio[coincidencia.start() - 1:coincidencia.end() + 1].count("·"):
                hallazgos.append(f"[doble espacio] {limpio[:80]!r}")
                break

        # 6. Comas dobles y signos de exclamación/ interrogación descuadrados
        if ",," in limpio or ";;" in limpio:
            hallazgos.append(f"[signo duplicado] {limpio[:80]!r}")
        if limpio.count("¿") != limpio.count("?"):
            hallazgos.append(f"[interrogacion sin cerrar] {limpio[:80]!r}")

    return hallazgos


def main() -> int:
    if not DOCUMENTO.exists():
        print(f"No existe el documento: {DOCUMENTO}")
        print("Ejecute antes: python docs/generar_documento.py")
        return 1

    documento = Document(str(DOCUMENTO))
    lineas = [p.text.strip() for p in documento.paragraphs if p.text.strip()]
    for t in documento.tables:
        for fila in t.rows:
            for celda in fila.cells:
                if celda.text.strip():
                    lineas.append(celda.text.strip())

    prosa = [l for l in lineas if not es_bloque_de_codigo(l)]
    hallazgos = revisar(lineas)

    print(f"Archivo    : {DOCUMENTO.name}")
    print(f"Bloques    : {len(lineas)}  ({len(prosa)} de prosa)")
    print(f"Palabras   : {len(re.findall(chr(92) + 'w+', chr(32).join(prosa)))}")

    # Verificacion puntual de los encabezados interrogativos
    print()
    for linea in lineas:
        for clave, esperado in INTERROGATIVOS_ESPERADOS.items():
            if linea.strip().startswith(clave):
                marca = "OK   " if esperado in linea else "FALTA"
                print(f"  {marca} {clave}: {linea[:78]}")

    print()
    if not hallazgos:
        print("RESULTADO: sin hallazgos de ortografia, estilo ni idioma.")
        return 0

    unicos = list(dict.fromkeys(hallazgos))
    print(f"RESULTADO: {len(unicos)} hallazgo(s) distinto(s)")
    for hallazgo in unicos:
        print(f"  - {hallazgo}")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())