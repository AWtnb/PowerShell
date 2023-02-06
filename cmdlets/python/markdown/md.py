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
    def __init__(self, content:str) -> None:
        self.content = content.strip()
        self.content_lines = self.content.splitlines()

        initial_cmt_end = self._get_initial_comment_end()
        if initial_cmt_end > -1:
            target_lines = self.content_lines[initial_cmt_end+1:]
        else:
            target_lines = self.content_lines

        self.main_lines = []
        for line in target_lines:
            if len(self.main_lines) > 0 or len(line.strip()) > 0:
                self.main_lines.append(line)

    def _get_initial_comment_end(self) -> int:
        lines = self.content_lines
        if lines[0].strip().startswith("<!-"):
            end_of_cmt = 0
            for i, line in enumerate(lines):
                if line.strip().endswith("-->"):
                    end_of_cmt = i
                    break
            return end_of_cmt
        return -1

    def get_additional_css(self)  -> str:
        begin_of_css = -1
        for i, line in enumerate(self.main_lines):
            s = line.strip()
            if s.startswith("```") and s.endswith("css"):
                begin_of_css = i
                break
            else:
                if s and begin_of_css < 0:
                    break
        if 0 <= begin_of_css:
            end_of_css = [line.strip() for line in self.main_lines].index("```")
            return "\n".join(self.main_lines[begin_of_css+1:end_of_css])
        return ""


class ParseMd:

    def __init__(self, file_path:str) -> None:
        self._file_path = file_path
        self.raw_content = Path(self._file_path).read_text("utf-8")

        raw_md = RawMd(self.raw_content)
        self.main_lines = raw_md.main_lines
        self.additional_css = raw_md.get_additional_css()

    def get_main_content(self) -> str:
        css_lines = self.additional_css.splitlines()
        if len(css_lines) > 0:
            return "\n".join(self.main_lines[len(css_lines)+2:])
        return "\n".join(self.main_lines)

    def get_timestamp(self) -> str:
        date_fmt = r"%Y-%m-%d"
        file_epoch_time = Path(self._file_path).stat().st_mtime
        last_modified = datetime.datetime.fromtimestamp(file_epoch_time).strftime(date_fmt)
        today = datetime.datetime.today().strftime(date_fmt)
        if last_modified == today:
            return "update: {}".format(last_modified)
        return "contents updated: {} / document generated: {}".format(last_modified, today)

    def grep_comment(self, pattern:str, capture:int=1) -> list[str]:
        matches = []
        reg = re.compile(pattern)
        for line in self.raw_content.splitlines():
            if line.startswith("<!--"):
                m = reg.search(line.rstrip(" ->"))
                if m:
                    matches.append(m.group(capture))
        return matches


class Markup:

    def __init__(self, file_path:str) -> None:

        md = ParseMd(file_path)
        markup = mistletoe.markdown(md.get_main_content(), CustomRenderer)

        self.additional_style = '<style>\n{}\n</style>'.format(md.additional_css)

        dom = DomTree(markup)
        dom.adjust_index("//*[contains(@class, 'force-order')]")
        dom.set_heading_id("h2 | h3 | h4 | h5 | h6")
        dom.fix_spacing("h2 | h3 | h4 | h5")
        dom.set_link_target()
        dom.set_timestamp(md.get_timestamp())

        self.content = dom.get_content()
        self.toc = '<div class="toc">{}</div>'.format(dom.get_toc())

        title_comment = md.grep_comment(r"title: ?(.+)", 1)
        if len(title_comment) > 0:
            self.title = title_comment[0]
        else:
            top_heading = dom.get_top_heading()
            if top_heading:
                self.title = top_heading
            else:
                self.title = Path(file_path).stem
        self.title = '<title>{}</title>'.format(self.title)


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

    def append_elem(self, markup:str)  -> None:
        self.lines.append(markup)

    def get_markup(self) -> str:
        return "\n".join(["<head>"] + self.lines + ["</head>"])


def main(file_path:str, no_default_css:bool=False, invoke:bool=False, favicon_unicode:str="1F4DD") -> None:

    md_path = Path(file_path)
    if md_path.suffix != ".md":
        return

    html = Markup(str(md_path))

    head = HeadElem(html.title, favicon_unicode, no_default_css)
    extra_css_files = list(Path(md_path.parent).glob("*.css"))
    for css_path in extra_css_files:
        print("  + external style sheet: '{}'".format(css_path.name))
        head.append_elem('<!-- from additional style sheet: "{}" -->'.format(css_path.name))
        head.append_style(css_path)

    head.append_elem(html.additional_style)

    full_html = "\n".join([
        '<!DOCTYPE html>',
        '<html lang="ja">',
        head.get_markup(),
        '<body>',
        "\n".join([
            '<div class="container">',
            html.toc,
            html.content,
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