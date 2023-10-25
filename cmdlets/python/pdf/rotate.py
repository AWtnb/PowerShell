"""
rotate pdf with pdfrw.
encoding: utf8
"""

import sys
from pathlib import Path

from pdfrw import PdfReader, PdfWriter

def main(file_path:str, clockwise:int=90):

    clockwise = int(clockwise)

    pdf = PdfReader(file_path)

    pages = pdf.pages
    for i, _ in enumerate(pages):
        pages[i].Rotate = (int(pages[i].inheritable.Rotate or 0) + clockwise) % 360

    pdf_path = Path(file_path)
    out_path = pdf_path.with_stem("{}_rotate{:03}".format(pdf_path.stem, clockwise))
    writer = PdfWriter(out_path)
    writer.trailer = pdf
    writer.write()

if __name__ == '__main__':
    main(*sys.argv[1:3])

