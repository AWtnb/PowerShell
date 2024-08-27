"""
insert new pdf into existing pdf with pdfrw.
encoding: utf8
"""

import sys
from pathlib import Path

from pdfrw import PdfReader, PdfWriter


def main(
    base_path: str, insert_path: str, insert_after: int = 1, out_name: str = ""
):
    base_pdf = PdfReader(base_path)
    insert_pdf = PdfReader(insert_path)

    base_count = len(base_pdf.pages)
    if insert_after < 0:
        insert_start_idx = base_count + insert_after
    else:
        insert_start_idx = insert_after - 1

    try:
        assert (
            insert_start_idx < base_count
        ), "invalid index: cannot insert after page {} (index {})".format(
            insert_after, insert_start_idx
        )
    except AssertionError as err:
        print(err, file=sys.stderr)
        return

    pdf_path = Path(base_path)
    if 0 < len(out_name):
        out_path = pdf_path.with_stem(out_name)
    else:
        out_path = pdf_path.with_stem("{}_inserted".format(pdf_path.stem))

    writer = PdfWriter(out_path)

    if insert_start_idx < 0:
        writer.addpages(insert_pdf.pages)

    for i, page in enumerate(base_pdf.pages):
        writer.addPage(page)
        if i == insert_start_idx:
            writer.addpages(insert_pdf.pages)

    writer.write()


if __name__ == "__main__":
    args = sys.argv[1:]
    if len(args) < 4:
        args.append("")
    main(*args[:2], int(args[2]), args[3])
