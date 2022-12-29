"""
PyMuPDF で PDF を画像に変換する
encoding: utf8
"""

import sys
from pathlib import Path

import fitz

def main(file_path:str):
    # https://pymupdf.readthedocs.io/en/latest/recipes-images.html
    pdf_path = Path(file_path)

    pdf = fitz.open(str(pdf_path))
    for i, page in enumerate(pdf):
        out_path = str(pdf_path.with_name(pdf_path.stem + "_{:04}.png".format(i)))
        pix = page.get_pixmap(dpi=300)
        pix.save(out_path)


if __name__ == '__main__':
    main(sys.argv[1])
