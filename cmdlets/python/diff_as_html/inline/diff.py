import argparse
from pathlib import Path

import diff_match_patch
import lxml.html


class PyDiff:
    def __init__(self, from_file: str, to_file: str) -> None:
        from_path = Path(from_file)
        to_path = Path(to_file)
        from_content = from_path.read_text("utf-8")
        to_content = to_path.read_text("utf-8")

        dmp = diff_match_patch.diff_match_patch()
        dmp.Diff_Timeout = 25 # slowly, take it easy...
        diffs = dmp.diff_main(from_content, to_content)
        dmp.diff_cleanupSemantic(diffs)
        self._heading = "<h1>'{}'→'{}'</h1>".format(
            from_path.name, to_path.name
        )

        diff_html = dmp.diff_prettyHtml(diffs)
        diff_tree = lxml.html.fromstring(diff_html)
        diff_tree.classes.add("diff-container")
        for elem in list(diff_tree):
            if elem.tag != "span":
                elem.attrib.pop("style", None)
                if elem.tag == "del":
                    elem.attrib["inert"] = "true"

        self._diff_tree = diff_tree

    @staticmethod
    def _get_filler() -> lxml.html.Element:
        filler = lxml.html.Element("span")
        filler.classes.add("filler")
        return filler

    def _compress_markup(self) -> lxml.html.Element:
        root = lxml.html.Element("div")
        root.classes.add("diff-container")
        for elem in list(self._diff_tree):
            if elem.tag != "span":
                root.append(elem)
            else:
                text_list = [t for t in elem.xpath("text()")]
                if len(text_list) < 3:
                    root.append(elem)
                else:
                    compressed = lxml.html.Element("span")
                    compressed.classes.add("compressed-lines")
                    compressed.text = text_list[0]
                    filler = self._get_filler()
                    filler.tail = text_list[-1]
                    compressed.append(filler)
                    if text_list[-1].endswith("\u00B6"):
                        br_tag = lxml.html.Element("br")
                        compressed.append(br_tag)
                    root.append(compressed)
        return root

    def get_markup(self, compress: bool = False) -> str:
        if compress:
            elem = self._compress_markup()
        else:
            elem = self._diff_tree
        return self._heading + lxml.html.tostring(elem, encoding="unicode").replace("\u00B6", "\u21B5")


def main(
    from_file: str,
    to_file: str,
    out_path: str,
    css_path: str,
    js_path: str,
    compress: bool,
) -> None:
    css = "<style>{}</style>".format(Path(css_path).read_text("utf-8"))
    js = "<script>{}</script>".format(Path(js_path).read_text("utf-8"))
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
            js,
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
    js_path = Path(__file__).with_name("event.js")
    main(
        args.fromFile,
        args.toFile,
        args.outFile,
        css_path,
        js_path,
        args.compress,
    )
