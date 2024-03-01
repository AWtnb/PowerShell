"""
Set title metadata of pdf with PyPDF.
encoding: utf8
"""

import sys
from pathlib import Path

from pypdf import PdfReader, PdfWriter


def main(file_path: str, title_str: str = "", preserve_untouched_data: bool = False) -> None:
    src_pdf = PdfReader(file_path)
    pdf_path = Path(file_path)
    new_path = pdf_path.with_stem(pdf_path.stem + "_newmatadata")
    new_pdf = PdfWriter()

    for page in src_pdf.pages:
        new_pdf.add_page(page)

    metadata = {}
    if preserve_untouched_data:
        old_metadata = {
            key: src_pdf.metadata[key] for key in src_pdf.metadata.keys()
        }
        old_metadata["/Title"] = title_str
        metadata = old_metadata
    else:
        metadata = {"/Title": title_str}

    new_pdf.add_metadata(metadata)
    new_pdf.write(str(new_path))


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3] == "True")
