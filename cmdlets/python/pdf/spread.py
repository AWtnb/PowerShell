"""
pdfrw で単ページ PDF を見開きにする
encoding: utf8
"""

import argparse
from pathlib import Path

from pdfrw import PdfReader, PdfWriter, PageMerge

# https://github.com/pmaupin/pdfrw/blob/master/examples/4up.py

def allocate(*pages):
    result = PageMerge() + pages
    if len(result) > 1:
        result[-1].x += result[0].w
    return result.render()


def main(file_path:str, single_toppage:bool=False):

    pdf_path = Path(file_path)
    out_path = str(pdf_path.with_stem(pdf_path.stem + "_spread"))

    pages = PdfReader(file_path).pages

    out_pages = []
    if single_toppage:
        top_page = pages.pop(0)
        out_pages.append(allocate(top_page))

    for idx in range(0, len(pages), 2):
        lpage = pages[idx]
        if idx+1 == len(pages):
            out_pages.append(allocate(lpage))
        else:
            rpage = pages[idx+1]
            out_pages.append(allocate(lpage, rpage))

    PdfWriter(out_path).addpages(out_pages).write()

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("filePath", type=str)
    parser.add_argument("--singleTopPage", action="store_true")
    args = parser.parse_args()

    main(args.filePath, args.singleTopPage)

