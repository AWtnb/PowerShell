"""
convert pdf to image file with PyMuPDF.
encoding: utf8
"""

import sys
from pathlib import Path

import fitz


def main(file_path: str, dpi: int):
    # https://pymupdf.readthedocs.io/en/latest/recipes-images.html
    pdf_path = Path(file_path)

    pdf = fitz.open(str(pdf_path))
    for i, page in enumerate(pdf):
        out_path = str(pdf_path.with_name(
            pdf_path.stem + "_{:04}.png".format(i+1)))
        pix = page.get_pixmap(dpi=dpi)
        pix.save(out_path)


if __name__ == '__main__':
    main(sys.argv[1], int(sys.argv[2]))
