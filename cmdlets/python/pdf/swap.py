"""
pdfrw で PDF の指定ページを別ファイルに差し替える
encoding: utf8
"""

import sys
from pathlib import Path

from pdfrw import PdfReader, PdfWriter


def main(base_path:str, insert_path:str, swap_start:int=1):

    base_pdf = PdfReader(base_path)
    insert_pdf = PdfReader(insert_path)

    if swap_start < 0:
        try:
            swap_start = len(base_pdf.pages) + swap_start + 1
            assert 0 < swap_start, "invalid index!"
        except AssertionError as err:
            print(err, file=sys.stderr)
            return

    swap_end = swap_start + len(insert_pdf.pages) - 1

    bp = Path(base_path)
    out_path = bp.with_stem("{}_swap{:03}-{:03}".format(bp.stem, swap_start, swap_end))

    writer = PdfWriter(out_path)
    writer.addpages(base_pdf.pages[:(swap_start - 1)])
    writer.addpages(insert_pdf.pages)
    if swap_end < len(base_pdf.pages):
        writer.addpages(base_pdf.pages[swap_end:])
    writer.write()

if __name__ == '__main__':
    main(*sys.argv[1:3], int(sys.argv[3]))