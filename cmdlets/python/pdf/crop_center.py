"""
pdfrw でB4ゲラ中央に配置された単ページを半ページサイズに切り出す
encoding: utf8
"""

import sys
from pathlib import Path

from pdfrw import PdfReader, PdfWriter, PageMerge

def crop_center(src):
    # viewrect => x_from_left, y_from_top, width, height of mediabox
    return PageMerge().add(src, viewrect=(0.25, 0, 0.5, 1)).render()

def main(file_path:str, crop_head:bool, crop_tail:bool) -> None:

    pdf_path = Path(file_path)
    out_path = str(pdf_path.with_stem(pdf_path.stem + "_cropped"))

    writer = PdfWriter(out_path)
    pages = PdfReader(file_path).pages
    top_page = pages[0]
    last_page = pages[-1]

    if crop_head:
        writer.addpage(crop_center(top_page))
    else:
        writer.addpage(top_page)

    writer.addpages(pages[1:-1])

    if crop_tail:
        writer.addpage(crop_center(last_page))
    else:
        writer.addpage(last_page)

    writer.write()


if __name__ == '__main__':
    file_path, mode = sys.argv[1:3]
    crop_head, crop_tail = {
        "head": [True, False],
        "tail": [False, True],
        "both": [True, True],
    }[mode]
    main(file_path, crop_head, crop_tail)
