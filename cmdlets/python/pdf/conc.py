"""
pdfrw で PDF を結合する
encoding: utf8
"""

# https://github.com/pmaupin/pdfrw/blob/master/examples/cat.py

import sys
from pathlib import Path

from pdfrw import PdfReader, PdfWriter

def main(file_list_path:str, out_path:str):

    path_list = Path(file_list_path).read_text("utf-8").splitlines()

    writer = PdfWriter()

    for file_path in path_list:
        writer.addpages(PdfReader(file_path).pages)

    writer.write(out_path)

if __name__ == '__main__':
    main(*sys.argv[1:3])