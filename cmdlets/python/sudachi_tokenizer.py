"""
sudachiPy で形態素解析して JSON 形式でファイルに出力する

ビルドに rust を使用するようになったので、初回の pip install 時に rust がインストールされている必要がある。
エラーメッセージで案内される https://rustup.rs/ をインストールして本体を再起動してから実行すれば解決する（はず）。

encoding: utf8
"""

import sys
import re
from pathlib import Path

from sudachipy import tokenizer, dictionary

def main(input_file_path:str, output_file_path:str, ignore_paren:bool=False):

    reg = re.compile(r"\(.+?\)|\[.+?\]|（.+?）|［.+?］")

    tokenizer_obj = dictionary.Dictionary().create()
    lines = Path(input_file_path).read_text("utf-8").splitlines()

    out = []
    for line in lines:
        if not line:
            out.append({"line": "", "tokens": []})
        else:
            target = line
            if ignore_paren:
                target = reg.sub("", line)
            tokens = []
            for t in tokenizer_obj.tokenize(target, tokenizer.Tokenizer.SplitMode.C):
                tokens.append({
                    "surface": t.surface(),
                    "pos": t.part_of_speech()[0],
                    "reading": t.reading_form(),
                    "c_type": t.part_of_speech()[4],
                    "c_form": t.part_of_speech()[5]
                })
            out.append({"line": line, "tokens": tokens})

    Path(output_file_path).write_text(str(out), "utf-8")

if __name__ == "__main__":
    ignore_paren = sys.argv[3] == "IgnoreParen"
    main(*sys.argv[1:3], ignore_paren)
