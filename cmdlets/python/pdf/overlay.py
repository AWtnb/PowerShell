"""
pdfrw で PDF 同士を重ねる
encoding: utf8
"""

# https://github.com/pmaupin/pdfrw/blob/master/examples/fancy_watermark.py

import sys
from pathlib import Path

from pdfrw import PdfReader, PdfWriter, PageMerge

def main(file_path:str, watermark_path:str) -> None:

    if Path(watermark_path).suffix != ".pdf":
        return

    out_path = Path(file_path).with_stem("watermarked_" + Path(file_path).stem)
    pdf_reader = PdfReader(file_path)
    watermark_reader = PdfReader(watermark_path)

    for i,page in enumerate(pdf_reader.pages):
        if i < len(watermark_reader.pages):
            watermark = PageMerge().add(watermark_reader.pages[i])[0]
            PageMerge(page).add(watermark, prepend=False).render()

    PdfWriter(out_path, trailer=pdf_reader).write()

if __name__ == '__main__':
    main(*sys.argv[1:3])