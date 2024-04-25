"""
Combine two PDFs alternately and insert watermark of filename.
encoding: utf8
"""

import sys
import tempfile
from pathlib import Path

from pdfrw import PdfReader, PdfWriter

from package.fnwatermark import FileNameWatermark


def new_watermarkred_pdf(file_path:str, work_dir:str) -> str:
    wm = FileNameWatermark(work_dir, file_path)
    wm.create()
    return wm.overlay()


def main(odd_file_path:str, even_file_path:str, out_path:str) -> None:
    writer = PdfWriter()

    with tempfile.TemporaryDirectory() as tmpdir:
        wm_path_odd = new_watermarkred_pdf(odd_file_path, tmpdir)
        wm_path_even = new_watermarkred_pdf(even_file_path, tmpdir)
        if wm_path_odd and wm_path_even:

            odd_pages = PdfReader(wm_path_odd).pages
            even_pages = PdfReader(wm_path_even).pages

            counter = 0
            for i, page in enumerate(odd_pages):
                writer.addpage(page)
                if i < len(even_pages):
                    writer.addpage(even_pages[i])
                counter = i

            rest = even_pages[counter:]
            for page in rest:
                writer.addPage(page)

            writer.write(out_path)
            Path(wm_path_odd).unlink()
            Path(wm_path_even).unlink()


if __name__ == '__main__':
    main(*sys.argv[1:4])