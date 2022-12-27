"""
pdfrw で単ページ PDF を見開きにする
encoding: utf8
"""

import sys
from pathlib import Path

from pdfrw import PdfReader, PdfWriter, PageMerge

# https://github.com/pmaupin/pdfrw/blob/master/examples/4up.py

def fixpage(*pages):
    result = PageMerge() + pages
    if len(result) > 1:
        result[-1].x += result[0].w
    return result.render()


def main(file_path:str):

    pdf_path = Path(file_path)
    out_path = str(pdf_path.with_stem(pdf_path.stem + "_spread"))

    pages = PdfReader(file_path).pages

    out_pages = []
    for idx in range(0, len(pages), 2):
        lpage = pages[idx]
        if idx+1 == len(pages):
            out_pages.append(fixpage(lpage))
        else:
            rpage = pages[idx+1]
            out_pages.append(fixpage(lpage, rpage))

    PdfWriter(out_path).addpages(out_pages).write()

if __name__ == '__main__':
    main(sys.argv[1])
