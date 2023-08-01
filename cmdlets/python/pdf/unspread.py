"""
pdfrw で見開きの PDF を単ページに分割する
encoding: utf8
"""

import sys
from pathlib import Path

from pdfrw import PdfReader, PdfWriter, PageMerge

from pdfpagebox import PdfPageBox


class Unspreader:
    def __init__(self, file_path: str, vertical: bool = False) -> None:
        pdf_path = Path(file_path)
        self.org_name = pdf_path.name
        self.out_path = str(pdf_path.with_stem(pdf_path.stem + "_unspread"))
        self.pages = PdfReader(file_path).pages
        self.vertical = vertical
        if vertical:
            self.size_list = [PdfPageBox(p).media.height for p in self.pages]
        else:
            self.size_list = [PdfPageBox(p).media.width for p in self.pages]

    def split(self, page):
        if self.vertical:
            for y_pos in (0, 0.5):
                yield PageMerge().add(page, viewrect=(0, y_pos, 1, 0.5)).render()
        else:
            for x_pos in (0, 0.5):
                yield PageMerge().add(page, viewrect=(x_pos, 0, 0.5, 1)).render()

    def execute(self) -> None:
        writer = PdfWriter(self.out_path)
        for i, page in enumerate(self.pages):
            size = self.size_list[i]
            if size != max(self.size_list):
                writer.addpage(page)
                try:
                    assert size == min(self.size_list), "page {} is neither largest nor smallest size! (skipped unspreading)".format(i+1)
                except AssertionError as err:
                    print("[WARNING] '{}': {}".format(self.org_name, err), file=sys.stderr)
            else:
                writer.addpages(self.split(page))
        writer.write()


if __name__ == '__main__':
    file_path = sys.argv[1]
    is_vertical = len(sys.argv) == 3 and sys.argv[2] == "vertical"
    uns = Unspreader(file_path, is_vertical)
    uns.execute()
