"""
Combine two PDFs alternately
encoding: utf8
"""

import sys
from pdfrw import PdfReader, PdfWriter


def main(odd_file_path:str, even_file_path:str, out_path:str) -> None:
    writer = PdfWriter()

    odd_pages = PdfReader(odd_file_path).pages
    even_pages = PdfReader(even_file_path).pages

    counter = 0
    for i, page in enumerate(odd_pages):
        writer.addpage(page)
        if i < len(even_pages):
            writer.addpage(even_pages[i])
        counter = i

    if counter < len(even_pages):
        for j in range(counter+1, len(even_pages)):
            writer.addpage(even_pages[j])

    writer.write(out_path)


if __name__ == '__main__':
    main(*sys.argv[1:4])