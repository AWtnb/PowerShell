"""
split pages of pdf with pdfrw.
encoding: utf8
"""

import sys
from pathlib import Path

from pdfrw import PdfReader, PdfWriter

def main(file_path:str):
    pdf_path = Path(file_path)
    pdf = PdfReader(file_path)
    for i, page in enumerate(pdf.pages):
        out_path = pdf_path.with_stem("{}_p{:03}".format(pdf_path.stem, i+1))
        writer = PdfWriter(out_path)
        writer.addPage(page)
        writer.write()


if __name__ == '__main__':
    main(*sys.argv[1:2])

