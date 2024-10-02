"""
split spreaded pdf pages to single pages with pdfrw.
encoding: utf8
"""

import sys
from pathlib import Path
import argparse

from pdfrw import PdfReader, PdfWriter, PageMerge

from package.pdfbox import PdfBox


class PdfPages:
    def __init__(self, file_path: str, vertical: bool) -> None:
        self._file_path = file_path
        self._pages = PdfReader(self._file_path).pages
        self._vertical = vertical

        if self._vertical:
            self._size_variation = [PdfBox(p).media.height for p in self._pages]
        else:
            self._size_variation = [PdfBox(p).media.width for p in self._pages]

    def is_max(self, idx: int) -> bool:
        return self._size_variation[idx] == max(self._size_variation)

    def is_min(self, idx: int) -> bool:
        return self._size_variation[idx] == min(self._size_variation)

    def is_first(self, idx: int) -> bool:
        return idx == 0

    def is_last(self, idx: int) -> bool:
        return idx == len(self._pages) - 1

    def check_weird_size(self, idx: int):
        if not self.is_min(idx):
            print(
                "skipped unspreading page {}".format(idx + 1),
                "because this page is neither largest nor smallest size.",
                file=sys.stderr,
            )

    def unspread(self, idx: int, backwards: bool):
        page = self._pages[idx]
        for i in (0, 0.5):
            if self._vertical:
                if backwards:
                    rect = (0, (0.5 - i), 1, 0.5)
                else:
                    rect = (0, i, 1, 0.5)
            else:
                if backwards:
                    rect = ((0.5 - i), 0, 0.5, 1)
                else:
                    rect = (i, 0, 0.5, 1)
            yield PageMerge().add(page, viewrect=rect).render()

    def centerize(self, idx: int):
        if self._vertical:
            rect = (0, 0.25, 1, 0.5)
        else:
            rect = (0.25, 0, 0.5, 1)
        page = self._pages[idx]
        return PageMerge().add(page, viewrect=rect).render()

    def unspread_pages(
        self,
        suffix: str = "",
        single_top: bool = False,
        single_last: bool = False,
        to_left: bool = False,
    ):
        p = Path(self._file_path)
        out_path = p.with_stem(p.stem + suffix)
        writer = PdfWriter(out_path)

        for i, page in enumerate(self._pages):
            if self.is_max(i):
                if single_top and self.is_first(i):
                    writer.addpage(self.centerize(i))
                    continue
                if single_last and self.is_last(i):
                    writer.addpage(self.centerize(i))
                    continue
                writer.addpages(self.unspread(i, to_left))
            else:
                self.check_weird_size(i)
                writer.addpage(page)
        writer.write()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("filePath", type=str)
    parser.add_argument("--vertical", action="store_true")
    parser.add_argument("--singleTop", action="store_true")
    parser.add_argument("--singleLast", action="store_true")
    parser.add_argument("--backwards", action="store_true")
    args = parser.parse_args()

    ps = PdfPages(args.filePath, args.vertical)
    ps.unspread_pages(
        "_unspread", args.singleTop, args.singleLast, args.backwards
    )
