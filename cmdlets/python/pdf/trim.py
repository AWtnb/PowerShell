"""
trim margin of B4 size pdf with pdfrw.
encoding: utf8
"""

import sys
from pathlib import Path

from pdfrw import PdfReader, PdfWriter, PageMerge


def main(
    file_path: str, margin_horizontal: float = 0.08, margin_vertical: float = 0.08
) -> None:
    rect = (
        margin_horizontal,
        margin_vertical,
        (1 - margin_horizontal * 2),
        (1 - margin_vertical * 2),
    )

    pdf_path = Path(file_path)
    out_path = str(pdf_path.with_stem(pdf_path.stem + "_trim"))

    writer = PdfWriter(out_path)
    pages = PdfReader(file_path).pages

    for page in pages:
        trimmed = PageMerge().add(page, viewrect=rect).render()
        writer.addpage(trimmed)

    writer.write()


if __name__ == "__main__":
    args = sys.argv[1:]
    main(args[0], float(args[1]), float(args[2]))
