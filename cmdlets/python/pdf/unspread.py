"""
pdfrw で見開きの PDF を単ページに分割する
encoding: utf8
"""

import sys
from pathlib import Path

from pdfrw import PdfReader, PdfWriter, PageMerge

from pdfpagebox import PdfPageBox

# https://github.com/pmaupin/pdfrw/blob/master/examples/unspread.py
def splitpage(src):
    for x_pos in (0, 0.5):
        yield PageMerge().add(src, viewrect=(x_pos, 0, 0.5, 1)).render()

def get_width(page) -> float:
    box = PdfPageBox(page)
    return box.media.width

def main(file_path:str):

    pdf_path = Path(file_path)
    out_path = str(pdf_path.with_stem(pdf_path.stem + "_unspread"))

    writer = PdfWriter(out_path)
    pages = PdfReader(file_path).pages

    page_widths = [get_width(p) for p in pages]

    for i, page in enumerate(pages):
        pw = get_width(page)
        if pw != max(page_widths):
            writer.addpage(page)
            try:
                assert pw == min(page_widths), "page {} is WEIRD size! (skipped unspreading)".format(i+1)
            except AssertionError as err:
                print("ERROR on processing '{}': {}".format(pdf_path.name, err), file=sys.stderr)
        else:
            writer.addpages(splitpage(page))
    writer.write()


if __name__ == '__main__':
    main(sys.argv[1])
