"""
pymupdf で PDF からテキスト抽出
encoding: utf8
"""

import sys
from pathlib import Path

import fitz


def main(file_path: str):
    # https://pymupdf.readthedocs.io/en/latest/recipes-text.html
    pdf_path = Path(file_path)
    out_path = pdf_path.with_suffix(".txt")

    pdf = fitz.open(str(pdf_path))
    text = "".join([page.get_text() for page in pdf])
    out_path.write_text(text, encoding="utf-8")

if __name__ == '__main__':
    main(sys.argv[1])
