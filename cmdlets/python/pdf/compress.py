"""
PyPDF2 で PDF を圧縮する
encoding: utf8
"""

import sys
from pathlib import Path

from PyPDF2 import PdfReader, PdfWriter


def main(file_path: str):
    # https://pypdf2.readthedocs.io/en/3.0.0/user/file-size.html
    pdf_path = Path(file_path)
    reader = PdfReader(str(pdf_path))
    writer = PdfWriter()
    out_path = str(pdf_path.with_stem(pdf_path.stem + "_compress"))

    for page in reader.pages:
        page.compress_content_streams()
        writer.add_page(page)

    with open(str(out_path), "wb") as f:
        writer.write(f)

if __name__ == '__main__':
    main(sys.argv[1])
