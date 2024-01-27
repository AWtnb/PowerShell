"""
convert markdown to html
encoding: utf8
"""

import argparse
import datetime
import pathlib
import re
import urllib.parse
from pathlib import Path
import webbrowser

import mistletoe

from domtree import DomTree
from custom_renderer import CustomRenderer


class RawMd:

    def __init__(self, file_path:str) -> None:
        self._file_path = file_path
        self.content = Path(self._file_path).read_text("utf-8").strip()

    def get_timestamp(self) -> str:
        date_fmt = r"%Y-%m-%d"
        file_epoch_time = Path(self._file_path).stat().st_mtime
        last_modified = datetime.datetime.fromtimestamp(file_epoch_time).strftime(date_fmt)
        today = datetime.datetime.today().strftime(date_fmt)
        if last_modified == today:
            return "update: {}".format(last_modified)
        return "contents updated: {} / document generated: {}".format(last_modified, today)


class MdHtml:

    def __init__(self, file_path:str) -> None:

        raw_md = RawMd(file_path)
        markup = mistletoe.markdown(raw_md.content, CustomRenderer)

        dom = DomTree(markup)
        self.additional_style = '<style>\n{}\n</style>'.format(dom.trim_leading_css_block())

        dom.adjust_index("//*[contains(@class, 'force-order')]")
        dom.set_heading_id("h2 | h3 | h4 | h5 | h6")
        dom.fix_spacing("h2 | h3 | h4 | h5")
        dom.set_link_target()
        dom.set_timestamp(raw_md.get_timestamp())
        dom.render_pagebreak()
        dom.render_arrowlink()
        dom.render_pdflink()
        dom.render_td()
        dom.render_codeblock_label()

        self.content = dom.get_content()
        self.toc = '<div class="toc">{}</div>'.format(dom.get_toc())
        self.comments = dom.get_comments()

        title_comment = self.grep_comment(r"^title:")
        if len(title_comment) > 0:
            self.title = title_comment[0][6:].strip()
        else:
            top_heading = dom.get_top_heading()
            if top_heading:
                self.title = top_heading
            else:
                self.title = Path(file_path).stem
        self.title = '<title>{}</title>'.format(self.title)

    def grep_comment(self, pattern) -> list[str]:
        reg = re.compile(pattern)
        found = []
        for c in self.comments:
            if reg.search(c):
                found.append(c)
        return found


class HeadElem:
    def __init__(self, title:str, favicon_unicode:str, no_default_css:bool=False) -> None:
        svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><text x="50%" y="50%" style="dominant-baseline:central;text-anchor:middle;font-size:90px;">&#x{};</text></svg>'.format(favicon_unicode)
        self.lines = [
            '<meta charset="utf-8">',
            '<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0">',
            title,
            '<link rel="icon" href="data:image/svg+xml,{}">'.format(urllib.parse.quote(svg))]

        if not no_default_css:
            css_path = Path(__file__).with_name("markdown.css")
            self.append_style(css_path)

        self.append_elem("<style>td.left{text-align:left;}td.center{text-align:center;}td.right{text-align:right;}</style>")


    def append_style(self, css_path:pathlib.Path) -> None:
        elem = "<style>\n{}\n</style>".format(css_path.read_text("utf-8"))
        self.lines.append(elem)

    def append_elem(self, markup:str) -> None:
        self.lines.append(markup)

    def get_markup(self) -> str:
        return "\n".join(["<head>"] + self.lines + ["</head>"])


def main(file_path:str, no_default_css:bool=False, invoke:bool=False, favicon_unicode:str="1F4DD") -> None:

    md_path = Path(file_path)
    if md_path.suffix != ".md":
        return

    md_html = MdHtml(str(md_path))

    head = HeadElem(md_html.title, favicon_unicode, no_default_css)
    extra_css_files = list(Path(md_path.parent).glob("*.css"))
    for css_path in extra_css_files:
        print("  + external style sheet: '{}'".format(css_path.name))
        head.append_elem('<!-- from additional style sheet: "{}" -->'.format(css_path.name))
        head.append_style(css_path)

    head.append_elem(md_html.additional_style)

    full_html = "\n".join([
        '<!DOCTYPE html>',
        '<html lang="ja">',
        head.get_markup(),
        '<body>',
        "\n".join([
            '<div class="container">',
            md_html.toc,
            md_html.content,
            '</div>']),
        '</body>',
        '</html>'])

    ts = datetime.datetime.today().strftime(r"%Y%m%d")
    out_path = md_path.with_name(md_path.stem + "_" + ts + ".html")
    Path(out_path).write_text(full_html, "utf-8")

    if invoke:
        webbrowser.open(out_path)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("filePath", type=str)
    parser.add_argument("--faviconUnicode", default="1F4DD")
    parser.add_argument("--noDefaultCss", action="store_true")
    parser.add_argument("--invoke", action="store_true")
    args = parser.parse_args()

    main(args.filePath, args.noDefaultCss, args.invoke, args.faviconUnicode)