"""
pdfrw で見開きの PDF を単ページに分割する
encoding: utf8
"""

import sys
from pathlib import Path
import argparse

from pdfrw import PdfReader, PdfWriter, PageMerge

from package.pdfbox import PdfBox


class Unspreader:
    def __init__(self, file_path: str, vertical: bool = False) -> None:
        pdf_path = Path(file_path)
        self.org_name = pdf_path.name
        self.out_path = str(pdf_path.with_stem(pdf_path.stem + "_unspread"))
        self.pages = PdfReader(file_path).pages
        self.last_page_index = len(self.pages) - 1
        self.vertical = vertical
        if vertical:
            self.size_list = [PdfBox(p).media.height for p in self.pages]
        else:
            self.size_list = [PdfBox(p).media.width for p in self.pages]
        self.max_size = max(self.size_list)
        self.min_size = min(self.size_list)

    def split(self, page):
        for i in (0, 0.5):
            if self.vertical:
                rect = (0, i, 1, 0.5)
            else:
                rect = (i, 0, 0.5, 1)
            yield PageMerge().add(page, viewrect=rect).render()

    def centerize(self, page):
        if self.vertical:
            rect = (0, 0.25, 1, 0.5)
        else:
            rect = (0.25, 0, 0.5, 1)
        return PageMerge().add(page, viewrect=rect).render()

    def test_unspreadable(self, idx: int) -> bool:
        size = self.size_list[idx]
        if size == self.max_size:
            return True
        try:
            assert (
                size == self.min_size
            ), "page {} is neither largest nor smallest size! -> skipped unspreading".format(
                idx + 1
            )
        except AssertionError as err:
            print("[WARNING] '{}': {}".format(self.org_name, err), file=sys.stderr)
        finally:
            return False

    def execute(self, single_top: bool, single_last: bool) -> None:
        writer = PdfWriter(self.out_path)
        for i, page in enumerate(self.pages):
            if single_top and i == 0 and self.test_unspreadable(0):
                writer.addpage(self.centerize(page))
                continue
            if single_last and i == self.last_page_index and self.test_unspreadable(self.last_page_index):
                writer.addpage(self.centerize(page))
                continue
            if self.test_unspreadable(i):
                writer.addpages(self.split(page))
            else:
                writer.addpage(page)
        writer.write()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("filePath", type=str)
    parser.add_argument("--vertical", action="store_true")
    parser.add_argument("--singleTop", action="store_true")
    parser.add_argument("--singleLast", action="store_true")
    args = parser.parse_args()

    uns = Unspreader(args.filePath, args.vertical)
    uns.execute(args.singleTop, args.singleLast)
