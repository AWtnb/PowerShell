"""
pdfrw でB4ゲラの余白を取り除く
encoding: utf8
"""

import sys
from pathlib import Path

from pdfrw import PdfReader, PdfWriter, PageMerge

def main(file_path:str, tombow_percent_h:float=8.0, tombow_percent_v:float=8.0) -> None:

    tombow_ratio_h = tombow_percent_h / 100
    tombow_ratio_v = tombow_percent_v / 100
    rect = (
        tombow_ratio_h,
        tombow_ratio_v,
        (1-tombow_ratio_h*2),
        (1-tombow_ratio_v*2))

    pdf_path = Path(file_path)
    out_path = str(pdf_path.with_stem(pdf_path.stem + "_trim"))

    writer = PdfWriter(out_path)
    pages = PdfReader(file_path).pages

    for page in pages:
        trimmed = PageMerge().add(page, viewrect=rect).render()
        writer.addpage(trimmed)

    writer.write()


if __name__ == '__main__':
    args = sys.argv[1:3]
    main(args[0], float(args[1]))
