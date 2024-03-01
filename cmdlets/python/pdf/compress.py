"""
Compress pdf with PyPDF.
encoding: utf8
"""

import sys
from pathlib import Path

from pypdf import PdfReader, PdfWriter


def main(file_path: str):
    # https://pypdf.readthedocs.io/en/latest/user/file-size.html
    pdf_path = Path(file_path)
    reader = PdfReader(str(pdf_path))
    writer = PdfWriter()
    out_path = str(pdf_path.with_stem(pdf_path.stem + "_compress"))

    for page in reader.pages:
        writer.add_page(page)

    for page in writer.pages:
        page.compress_content_streams()

    with open(str(out_path), "wb") as f:
        writer.write(f)

if __name__ == '__main__':
    main(sys.argv[1])
