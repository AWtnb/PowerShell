"""
insert watermark of filename with pdfrw.
encoding: utf8
"""

import sys
import tempfile
from pathlib import Path

from pdfrw import PdfReader

from package.fnwatermark import FileNameWatermark

def add_watermark(file_path:str, start_idx:int=1):

    pdf = PdfReader(file_path)

    with tempfile.TemporaryDirectory() as tmpdir:

        wm = FileNameWatermark(tmpdir, file_path)
        wm.create(start_idx)
        wm.overlay()

        return len(pdf.pages)


def main(file_list_path:str, start_idx:str, mode:str):
    paths = Path(file_list_path).read_text("utf-8").splitlines()
    idx = int(start_idx)
    for p in paths:
        npages = add_watermark(p, idx)
        if mode == "through":
            idx += npages


if __name__ == '__main__':
    main(*sys.argv[1:4])