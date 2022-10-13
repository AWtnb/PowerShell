"""
pdfrw で PDF の奇数／偶数ページを抽出する
encoding: utf8
"""

import argparse
from pathlib import Path

from pdfrw import PdfReader, PdfWriter

def main(file_path:str, pickup_even:bool):

    pdf = PdfReader(file_path)

    pdf_path = Path(file_path)
    suffix = "even" if pickup_even else "odd"
    out_path = pdf_path.with_stem("{}_{}".format(pdf_path.stem, suffix))

    writer = PdfWriter(out_path)
    start_idx = 1 if pickup_even else 0
    for i in range(start_idx, len(pdf.pages), 2):
        writer.addpage(pdf.pages[i])
    writer.write()

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("filePath", type=str)
    parser.add_argument("--evenPages", action="store_true")
    args = parser.parse_args()
    main(args.filePath, args.evenPages)