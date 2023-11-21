"""
get similality of lines.

encoding: utf8
"""

import sys
import re
import itertools
from difflib import SequenceMatcher
from pathlib import Path


class UnicodeMapper:
    def __init__(self, repl: str) -> None:
        if len(repl):
            self._repl = ord(repl)
        else:
            self._repl = None
        self._mapping = {}

    def register(self, char_code: int) -> None:
        self._mapping[char_code] = self._repl

    def register_range(self, pair: list) -> None:
        start, end = pair
        for i in range(int(start, 16), int(end, 16) + 1):
            self.register(i)

    def register_ranges(self, pairs: list) -> None:
        for pair in pairs:
            self.register_range(pair)

    def get_mapping(self) -> dict:
        return self._mapping


class NoiseMapping:
    def __init__(self, repl: str) -> None:
        _mapper = UnicodeMapper(repl)
        _mapper.register(int("30FB", 16))  # KATAKANA MIDDLE DOT
        _mapper.register_range(["2018", "201F"])  # quotation
        _mapper.register_range(["2E80", "2EF3"])  # kangxi
        _mapper.register_ranges(
            [  # ascii
                ["0021", "002F"],
                ["003A", "0040"],
                ["005B", "0060"],
                ["007B", "007E"],
            ]
        )
        _mapper.register_ranges(
            [  # bars
                ["2010", "2017"],
                ["2500", "2501"],
                ["2E3A", "2E3B"],
            ]
        )
        _mapper.register_ranges(
            [  # fullwidth
                ["25A0", "25EF"],
                ["3000", "3004"],
                ["3008", "3040"],
                ["3097", "30A0"],
                ["3097", "30A0"],
                ["30FD", "30FF"],
                ["FF01", "FF0F"],
                ["FF1A", "FF20"],
                ["FF3B", "FF40"],
                ["FF5B", "FF65"],
            ]
        )

        self._mapping = _mapper.get_mapping()

    def cleanup(self, s: str) -> str:
        return s.translate(str.maketrans(self._mapping))


NOISE_MAPPIING = NoiseMapping(" ")


class LinesMatcher:
    def __init__(self, pair: list[str]) -> None:
        stack = []
        for line in pair[:2]:
            cleaned = NOISE_MAPPIING.cleanup(line)
            stack.append(re.sub(r"\s+", " ", cleaned))
        self._pair = stack

    def _has_common_prefix(self) -> bool:
        return self._pair[0].split(" ")[0] == self._pair[1].split(" ")[0]

    def compare(self) -> dict | None:
        if self._has_common_prefix():
            comp = SequenceMatcher(None, *self._pair)
            return {
                "prox": comp.ratio(),
                "a": self._pair[0],
                "b": self._pair[1],
            }
        return None


def main(input_file_path: str, output_file_path: str):
    lines = Path(input_file_path).read_text("utf-8").splitlines()
    out = []
    for pair in itertools.combinations(lines, 2):
        m = LinesMatcher(pair)
        result = m.compare()
        if result:
            out.append(result)
    Path(output_file_path).write_text(str(out), "utf-8")


if __name__ == "__main__":
    main(*sys.argv[1:3])
