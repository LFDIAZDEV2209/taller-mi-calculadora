"""
Renderiza el documento .docx a imagenes PNG para revision visual.

Convierte el archivo a PDF con Microsoft Word (via docx2pdf) y rasteriza
cada pagina con pypdfium2. Las imagenes quedan en una carpeta temporal
dentro de evidencias/revision y no forman parte de la entrega.

Uso:
    python evidencias/revision_paginas.py
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

import pypdfium2 as pdfium
from docx2pdf import convert

RAIZ = Path(__file__).resolve().parents[1]
DOCUMENTO = RAIZ / "docs" / "diaz_luis_taller.docx"
DESTINO = RAIZ / "evidencias" / "revision"
ESCALA = 2.0


def main() -> int:
    if not DOCUMENTO.exists():
        print(f"No existe: {DOCUMENTO}")
        return 1

    DESTINO.mkdir(parents=True, exist_ok=True)
    for viejo in DESTINO.glob("pagina_*.png"):
        viejo.unlink()

    pdf_temporal = DESTINO / "revision.pdf"
    print("Convirtiendo a PDF con Word...")
    convert(str(DOCUMENTO), str(pdf_temporal))

    documento_pdf = pdfium.PdfDocument(str(pdf_temporal))
    total = len(documento_pdf)
    print(f"Paginas: {total}")

    for indice in range(total):
        pagina = documento_pdf[indice]
        imagen = pagina.render(scale=ESCALA).to_pil()
        salida = DESTINO / f"pagina_{indice + 1:02d}.png"
        imagen.save(salida)
        print(f"  {salida.name}  ({imagen.width}x{imagen.height})")

    documento_pdf.close()
    pdf_temporal.unlink()
    print(f"\nRevision visual en: {DESTINO}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())