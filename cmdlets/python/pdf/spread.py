"""
spread pair of pdf pages with pdfrw.
encoding: utf8
"""

import argparse
from pathlib import Path

from pdfrw import PdfReader, PdfWriter, PageMerge

# https://github.com/pmaupin/pdfrw/blob/master/examples/4up.py


def allocate(pages, vertical: bool = False):
    if len(pages) < 2:
        return pages[0]
    result = PageMerge() + pages
    if vertical:
        result[0].y += result[-1].h
    else:
        result[-1].x += result[0].w
    return result.render()


def main(
    file_path: str,
    single_toppage: bool = False,
    backwards: bool = False,
    vertical: bool = False,
):

    pdf_path = Path(file_path)
    out_path = str(pdf_path.with_stem(pdf_path.stem + "_spread"))

    pages = PdfReader(file_path).pages

    out_pages = []
    if single_toppage:
        top_page = pages.pop(0)
        out_pages.append(top_page)

    for idx in range(0, len(pages), 2):
        current_page = pages[idx]
        if idx + 1 == len(pages):
            out_pages.append(current_page)
        else:
            next_page = pages[idx + 1]
            if backwards:
                out_pages.append(allocate([next_page, current_page], vertical))
            else:
                out_pages.append(allocate([current_page, next_page], vertical))

    PdfWriter(out_path).addpages(out_pages).write()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("filePath", type=str)
    parser.add_argument("--singleTopPage", action="store_true")
    parser.add_argument("--backwards", action="store_true")
    parser.add_argument("--vertical", action="store_true")
    args = parser.parse_args()

    main(args.filePath, args.singleTopPage, args.backwards, args.vertical)
