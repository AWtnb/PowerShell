"""
tokenize Japanese with sudachiPy and export to JSON.

- require: rusup (https://rustup.rs/) for initial build.

encoding: utf8
"""

import sys
import re
import json
from pathlib import Path

from sudachipy import tokenizer, dictionary


class ParsedLine:
    def __init__(self, line: str) -> None:
        self.raw_line = line
        self.line = line
        self.tokens = []

    def trim_paren(self) -> None:
        reg_paren = re.compile(
            r"\(.+?\)|\[.+?\]|\uff08.+?\uff09|\uff3b.+?\uff3d"
        )
        self.line = reg_paren.sub("", self.line)

    def trim_noise(self) -> None:
        reg_noise = re.compile(r"　　[^\d]?\d.*$|　→.+$")
        self.line = reg_noise.sub("", self.line)


def main(
    input_file_path: str,
    output_file_path: str,
    ignore_paren: bool = False,
    focus_name: bool = False,
):

    tknzr = dictionary.Dictionary().create()
    lines = Path(input_file_path).read_text("utf-8").splitlines()

    stack = []
    for line in lines:
        pl = None
        if len(line.strip()) < 1:
            pl = ParsedLine("")
        else:
            pl = ParsedLine(line)
            if ignore_paren:
                pl.trim_paren()
            if focus_name:
                pl.trim_noise()
            for t in tknzr.tokenize(pl.line, tokenizer.Tokenizer.SplitMode.C):
                pl.tokens.append(
                    {
                        "surface": t.surface(),
                        "dict_form": t.dictionary_form(),
                        "pos": t.part_of_speech()[0],
                        "reading": t.reading_form(),
                    }
                )
        stack.append(
            {"raw_line": pl.raw_line, "line": pl.line, "tokens": pl.tokens}
        )

    Path(output_file_path).write_text(json.dumps(stack))


if __name__ == "__main__":
    ignore_paren = sys.argv[3] == "IgnoreParen"
    focus_name = sys.argv[4] == "FocusName"
    main(*sys.argv[1:3], ignore_paren, focus_name)
