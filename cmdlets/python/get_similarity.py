"""
類似した行を取得する

encoding: utf8
"""

import sys
import itertools
from difflib import SequenceMatcher
from pathlib import Path

def get_first_elem(s:str) -> str:
    return s.replace("\u3001", ",").replace("\uff0c", ",").replace("\u30fb", ",").split(",")[0]

def main(input_file_path:str, output_file_path:str):
    lines = Path(input_file_path).read_text("utf-8").splitlines()
    out = []
    for pair in itertools.combinations(lines, 2):
        if get_first_elem(pair[0]) == get_first_elem(pair[1]):
            comp = SequenceMatcher(None, *pair)
            out.append({"prox": comp.ratio(), "a": pair[0], "b": pair[1]})
    Path(output_file_path).write_text(str(out), "utf-8")


if __name__ == "__main__":
    main(*sys.argv[1:3])
