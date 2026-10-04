"""
Genera las evidencias de ejecución real del taller "Mi calculadora".

Ejecuta la implementación de referencia en Python, captura su salida de
consola tal como se produce y la renderiza en imágenes PNG con el estilo de
una terminal, para incorporarlas como figuras en el documento del taller.

Las trazas largas se dividen en varias páginas de consola para que ninguna
figura desborde el alto útil de la hoja. El listado de archivos generados se
publica en un manifiesto que consume el generador del documento, de modo que
las figuras y sus epígrafes nunca queden desalineados.

Uso:
    python evidencias/generar_evidencias.py
"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

RAIZ = Path(__file__).resolve().parents[1]
SCRIPT = RAIZ / "src" / "python" / "mi_calculadora.py"
DESTINO = Path(__file__).resolve().parent / "consola"
MANIFIESTO = DESTINO / "MANIFESTO.json"

# Alto útil de la hoja con márgenes de 2,5 cm, menos el epígrafe.
ALTO_MAXIMO_FIGURA_CM = 18.5
LINEAS_POR_PARTE = 52

# Paleta inspirada en una terminal de Windows (Campbell / PowerShell)
FONDO = (12, 12, 12)
FONDO_BARRA = (32, 32, 32)
TEXTO = (204, 204, 204)
TEXTO_OK = (106, 205, 106)
TEXTO_ALERTA = (255, 196, 0)
TEXTO_TITULO = (0, 163, 226)
TEXTO_COMENTARIO = (106, 153, 85)

FUENTES_CANDIDATAS = [
    Path("C:/Windows/Fonts/consola.ttf"),
    Path("C:/Windows/Fonts/cour.ttf"),
    Path("C:/Windows/Fonts/lucon.ttf"),
]

TRAZAS = [
    ("01_caso_exitoso", "exito",
     "Las cuatro operaciones aritméticas"),
    ("02_division_entre_cero", "division_cero",
     "Manejo del error: división entre cero"),
    ("03_validacion_de_entradas", "validacion",
     "Validación de entradas inválidas"),
    ("04_orden_libre", "orden_libre",
     "Operaciones en orden distinto al catálogo"),
]


def fuente(tamano: int) -> ImageFont.FreeTypeFont:
    for ruta in FUENTES_CANDIDATAS:
        if ruta.exists():
            return ImageFont.truetype(str(ruta), tamano)
    raise RuntimeError("No se encontró una fuente monoespaciada en el sistema.")


def ejecutar_traza(nombre: str) -> list[str]:
    """Ejecuta el programa y devuelve las líneas reales de salida."""
    proceso = subprocess.run(
        [sys.executable, str(SCRIPT), nombre],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        cwd=str(RAIZ),
    )
    if proceso.returncode != 0:
        raise RuntimeError(f"La traza {nombre} falló:\n{proceso.stderr}")
    return proceso.stdout.splitlines()


def limpiar(lineas: list[str]) -> list[str]:
    """Quita líneas vacías consecutivas y conserva el resto tal cual."""
    salida: list[str] = []
    for linea in lineas:
        if not linea.strip() and (not salida or not salida[-1].strip()):
            continue
        salida.append(linea.rstrip())
    while salida and not salida[0].strip():
        salida.pop(0)
    return salida


def dividir(lineas: list[str], maximo: int) -> list[list[str]]:
    """Divide la salida en páginas de tamaño parejo.

    Se calcula primero cuántas partes hacen falta y luego se reparten los
    cortes de forma uniforme, ajustando cada corte a la línea en blanco más
    cercana al punto ideal. Así ninguna parte queda casi vacía y ninguna
    parte parte un bloque de código por la mitad.
    """
    total = len(lineas)
    if total <= maximo:
        return [list(lineas)]

    partes_needed = -(-total // maximo)
    blancos = [i for i, linea in enumerate(lineas) if not linea.strip()]

    cortes: list[int] = []
    inicio = 0
    for numero in range(1, partes_needed):
        ideal = round(total * numero / partes_needed)
        candidatos = [i for i in blancos if inicio < i < total - 1]
        if candidatos:
            corte = min(candidatos, key=lambda i: abs(i - ideal))
            # Un corte no debe dejar una parte mayor que el máximo
            if corte - inicio > maximo:
                corte = inicio + maximo
            elif inicio + maximo - corte > maximo * 0.75:
                corte = inicio + maximo
        else:
            corte = min(inicio + maximo, total - 1)
        cortes.append(corte)
        inicio = corte + 1

    partes: list[list[str]] = []
    previo = -1
    for corte in cortes + [total]:
        trozo = lineas[previo + 1:corte]
        while trozo and not trozo[0].strip():
            trozo.pop(0)
        if trozo:
            partes.append(trozo)
        previo = corte
    return partes or [list(lineas)]


def color_de(linea: str) -> tuple[int, int, int]:
    if linea.startswith("#"):
        return TEXTO_COMENTARIO
    if "ERROR" in linea:
        return TEXTO_ALERTA
    if any(p in linea for p in ("Resultado =", "RESULTADOS", "HISTORIAL")):
        return TEXTO_OK
    if linea.startswith("===") or linea.startswith("---"):
        return TEXTO_TITULO
    return TEXTO


def renderizar(lineas: list[str], titulo: str, salida: Path) -> None:
    fuente_texto = fuente(17)
    fuente_titulo = fuente(18)

    margen = 18
    alto_barra = 38
    interlinea = 23
    ancho_max = max((len(linea) for linea in lineas), default=40)

    ancho_texto = ancho_max * 10 + margen * 2
    ancho = max(ancho_texto, 780)
    alto = alto_barra + margen * 2 + interlinea * len(lineas)

    lienzo = Image.new("RGB", (ancho, alto), FONDO)
    dibujo = ImageDraw.Draw(lienzo)

    dibujo.rectangle([0, 0, ancho, alto_barra], fill=FONDO_BARRA)
    for i, color_boton in enumerate([(255, 95, 86), (255, 189, 46), (39, 201, 63)]):
        x = 18 + i * 22
        dibujo.ellipse([x, 12, x + 11, 23], fill=color_boton)
    dibujo.text((96, 9), f"PS  \u2014  {titulo}", font=fuente_titulo,
                fill=(180, 180, 180))

    y = alto_barra + margen
    for linea in lineas:
        dibujo.text((margen, y), linea, font=fuente_texto, fill=color_de(linea))
        y += interlinea

    salida.parent.mkdir(parents=True, exist_ok=True)
    lienzo.save(salida)


def main() -> None:
    DESTINO.mkdir(parents=True, exist_ok=True)
    for viejo in DESTINO.glob("*.png"):
        viejo.unlink()

    manifiesto: list[dict[str, object]] = []
    total_figuras = 0

    for nombre_traza, caso, titulo in TRAZAS:
        lineas = limpiar(ejecutar_traza(caso))
        partes = dividir(lineas, LINEAS_POR_PARTE)

        for indice, parte in enumerate(partes, 1):
            sufijo = f"_p{indice}" if len(partes) > 1 else ""
            archivo = f"{nombre_traza}{sufijo}.png"
            encabezado = (f"{titulo} ({indice}/{len(partes)})"
                          if len(partes) > 1 else titulo)
            renderizar(parte, encabezado, DESTINO / archivo)

            total_figuras += 1
            manifiesto.append({
                "archivo": archivo,
                "traza": nombre_traza,
                "descripcion": titulo,
                "parte": indice,
                "partes": len(partes),
                "lineas": len(parte),
                "rango": [indice, len(partes)],
            })
            print(f"[OK] {archivo}  ({len(parte)} líneas, "
                  f"parte {indice}/{len(partes)})")

    MANIFIESTO.write_text(
        json.dumps({"figuras": manifiesto}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"\n{total_figuras} figuras generadas en: {DESTINO}")
    print(f"Manifiesto: {MANIFIESTO.relative_to(RAIZ)}")


if __name__ == "__main__":
    main()