import argparse
import re
from pathlib import Path

import diff_match_patch
import lxml.html


class PyDiff:
    def __init__(self, from_file: str, to_file: str) -> None:
        from_content = Path(from_file).read_text("utf-8")
        to_content = Path(to_file).read_text("utf-8")

        dmp = diff_match_patch.diff_match_patch()
        diffs = dmp.diff_main(from_content, to_content)
        dmp.diff_cleanupSemantic(diffs)
        self._diff_html = dmp.diff_prettyHtml(diffs)

    def _compress_markup(self) -> str:
        diff_html = lxml.html.fromstring(self._diff_html)
        root = lxml.html.Element("div")
        root.classes.add("diff-container")
        for elem in list(diff_html):
            if elem.tag != "span":
                root.append(elem)
            else:
                text_list = [t for t in elem.xpath("text()")]
                if len(text_list) < 3:
                    root.append(elem)
                else:
                    filler = lxml.html.Element("span")
                    filler.classes.add("filler")
                    new_span = lxml.html.Element("span")
                    new_span.text = text_list[0]
                    new_span.append(filler)
                    new_span.tail = text_list[-1]
                    root.append(new_span)
        return lxml.html.tostring(root, encoding="unicode")

    def get_markup(self, compress: bool = False) -> str:
        if compress:
            return self._compress_markup()
        diff_html = lxml.html.fromstring(self._diff_html)
        diff_html.classes.add("diff-container")
        return lxml.html.tostring(diff_html, encoding="unicode")


def main(
    from_file: str, to_file: str, out_path: str, css_path: str, compress: bool
) -> None:
    css = "<style>{}</style>".format(Path(css_path).read_text("utf-8"))
    pd = PyDiff(from_file, to_file)
    title = Path(out_path).stem
    page_markup = "\n".join(
        [
            "<!DOCTYPE html>",
            '<html lang="ja">',
            "<head>",
            '<meta charset="utf-8" />',
            '<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0" />',
            "<title>{}</title>".format(title),
            css,
            "</head>",
            "<body>",
            '<div class="main">{}</div>'.format(pd.get_markup(compress)),
            "</body>",
            "</html>",
        ]
    )
    Path(out_path).write_text(page_markup, "utf-8")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("fromFile", type=str)
    parser.add_argument("toFile", type=str)
    parser.add_argument("outFile", type=str)
    parser.add_argument("--compress", action="store_true")
    args = parser.parse_args()

    css_path = Path(__file__).with_name("additional.css")
    main(args.fromFile, args.toFile, args.outFile, css_path, args.compress)
