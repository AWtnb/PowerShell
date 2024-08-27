"""
trim margin of B4 size pdf with pdfrw.
encoding: utf8
"""

import sys
from pathlib import Path

from pdfrw import PdfReader, PdfWriter, PageMerge


def main(
    file_path: str,
    margin_horizontal_ratio: float = 0.08,
    margin_vertical_ratio: float = 0.08,
) -> None:
    try:
        msg = "ratio should be float between 0 and 1!"
        assert 0 < margin_horizontal_ratio and margin_horizontal_ratio < 1, msg
        assert 0 < margin_vertical_ratio and margin_vertical_ratio < 1, msg
    except AssertionError as err:
        print(err)
        return

    rect = (
        margin_horizontal_ratio,
        margin_vertical_ratio,
        (1 - margin_horizontal_ratio * 2),
        (1 - margin_vertical_ratio * 2),
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
