import argparse
from pathlib import Path
from difflib import HtmlDiff

import lxml.html


class PyDiff:
    def __init__(self, from_path: str, to_path: str, skip_unchanged: bool) -> None:
        f_path = Path(from_path)
        t_path = Path(to_path)
        f = f_path.read_text("utf-8").splitlines()
        t = t_path.read_text("utf-8").splitlines()
        self._markup = HtmlDiff().make_table(f, t, fromdesc=f_path.name, todesc=t_path.name)
        self._html = lxml.html.fromstring(self._markup)
        self._skip_unchanged = skip_unchanged

    def _get_thead(self) -> lxml.html.Element:
        tr = lxml.html.Element("tr")
        for i, elem in enumerate(self._html.xpath("//thead/tr/th")):
            th = lxml.html.Element("th")
            if i in (1, 3):
                th.classes.add("diff_header")
                th.text = elem.text_content()
            tr.append(th)
        thead = lxml.html.Element("thead")
        thead.append(tr)
        return thead

    @staticmethod
    def _get_filler_row() -> lxml.html.Element:
        tr = lxml.html.Element("tr")
        for i in range(4):
            td = lxml.html.Element("td")
            if i in (0, 2):
                td.classes.add("diff_header")
            tr.append(td)
        tr.classes.add("filler")
        return tr

    @classmethod
    def _compress_unchanged_trs(cls, trs: list[lxml.html.Element]) -> list[lxml.html.Element]:
        changed = []
        for tr in trs:
            changed.append(tr)
            if len(changed) and "unchanged" in list(changed[-1].classes):
                changed.pop()
                if not len(changed) or "filler" not in list(changed[-1].classes):
                    changed.append(cls._get_filler_row())
        return changed

    def _get_trs(self) -> list[lxml.html.Element]:
        trs = []
        for table_row in self._html.xpath("//tbody/tr"):
            tr = lxml.html.Element("tr")
            for i, table_cell in enumerate(table_row.xpath("td")):
                if i in (0, 3):
                    continue
                table_cell.attrib.pop("nowrap", None)
                tr.append(table_cell)
            if tr.xpath("//span"):
                tr.classes.add("changed")
            else:
                tr.classes.add("unchanged")
            trs.append(tr)
        if self._skip_unchanged:
            return self._compress_unchanged_trs(trs)
        return trs

    def _get_tbody(self) -> lxml.html.Element:
        tbody = lxml.html.Element("tbody")
        trs = self._get_trs()
        for tr in trs:
            tbody.append(tr)
        return tbody

    def _get_table(self) -> lxml.html.Element:
        table = lxml.html.Element("table")
        table.append(self._get_thead())
        table.append(self._get_tbody())
        return table

    def get_body(self) -> lxml.html.Element:
        main = lxml.html.Element("main")
        main.append(self._get_table())
        body = lxml.html.Element("body")
        body.append(main)
        return lxml.html.tostring(body, encoding="unicode")


def main(from_file: str, to_file: str, out_path: str, css_path: str, skip_unchanged: bool) -> None:
    css = "<style>{}</style>".format(Path(css_path).read_text("utf-8"))
    pd = PyDiff(from_file, to_file, skip_unchanged)
    title = Path(out_path).stem
    html_page = "\n".join(["<!DOCTYPE html>", '<html lang="ja">', "<head>", '<meta charset="utf-8" />', '<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0" />', "<title>{}</title>".format(title), css, "</head>", pd.get_body(), "</html>"])
    Path(out_path).write_text(html_page, "utf-8")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("fromFile", type=str)
    parser.add_argument("toFile", type=str)
    parser.add_argument("outFile", type=str)
    parser.add_argument("--compress", action="store_true")
    args = parser.parse_args()

    css_path = Path(__file__).with_name("wrap.css")
    main(args.fromFile, args.toFile, args.outFile, css_path, args.compress)
