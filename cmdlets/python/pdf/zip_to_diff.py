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

            shorter = min(len(odd_pages), len(even_pages))

            for i in range(shorter):
                writer.addPage(odd_pages[i])
                writer.addPage(even_pages[i])

            # add rest of odd file (if exists)
            for page in odd_pages[shorter:]:
                writer.addPage(page)

            # add rest of even file (if exists)
            for page in even_pages[shorter:]:
                writer.addPage(page)

            writer.write(out_path)
            Path(wm_path_odd).unlink()
            Path(wm_path_even).unlink()


if __name__ == '__main__':
    main(*sys.argv[1:4])