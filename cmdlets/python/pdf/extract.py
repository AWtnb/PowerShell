"""
pdfrw で PDF からページ範囲を指定して抜き出す
encoding: utf8
"""

import sys
from pathlib import Path

from pdfrw import PdfReader, PdfWriter


def main(file_path: str, p_from: str, p_to: str):
    range_begin = int(p_from)
    range_end = int(p_to)

    pdf = PdfReader(file_path)
    pdf_path = Path(file_path)

    try:
        assert 0 != range_begin, "start-page index starts from 1 (or -1)!"
        if range_begin < 0:
            range_begin = len(pdf.pages) + range_begin + 1
            assert 0 < range_begin, "too small start-page!"
        assert 0 != range_end, "end-page index starts from 1 (or -1)!"
        if range_end < 0:
            range_end = len(pdf.pages) + range_end + 1
            assert 0 < range_end, "too small end-page!"
        assert range_begin <= range_end, "start-page is bigger than end-page!"
        assert range_end <= len(pdf.pages), "too big end-page!"
    except AssertionError as err:
        print(pdf_path.name)
        print("OUT-OF-RANGE-ERROR: {}".format(err), file=sys.stderr)
        return

    out_path = pdf_path.with_stem("{}_{:03}-{:03}".format(pdf_path.stem, range_begin, range_end))

    writer = PdfWriter(out_path)
    writer.addpages(pdf.pages[range_begin - 1 : range_end])
    writer.write()


if __name__ == "__main__":
    main(*sys.argv[1:4])
