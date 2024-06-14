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
import yaml
import webbrowser

import mistletoe

from domtree import DomTree
from custom_renderer import CustomRenderer


class MdContent:
    frontmatter = []
    main = []


class RawMd:
    def __init__(self, file_path: str) -> None:
        self._file_path = file_path
        self.content = Path(self._file_path).read_text("utf-8").strip()

    def get_timestamp(self) -> str:
        date_fmt = r"%Y-%m-%d"
        file_epoch_time = Path(self._file_path).stat().st_mtime
        last_modified = datetime.datetime.fromtimestamp(
            file_epoch_time
        ).strftime(date_fmt)
        today = datetime.datetime.today().strftime(date_fmt)
        if last_modified == today:
            return "update: {}".format(last_modified)
        return "contents updated: {} / document generated: {}".format(
            last_modified, today
        )

    def get_lines(self) -> MdContent:
        idxes = []
        lines = self.content.splitlines()
        reg = re.compile(r"^-{3,}$")
        for i, line in enumerate(lines):
            if reg.search(line):
                idxes.append(i)
        mc = MdContent()
        if 1 < len(idxes) and idxes[0] == 0:
            idx = idxes[1]
            mc.main = lines[idx:]
            mc.frontmatter = lines[1:idx]
        else:
            mc.main = lines
        return mc


"""
---
title: title of document
styles:
  - p:
      color: red
      border: 1px solid yellow
  - h1:
      color: blue
css-vars:
  width-container: 42rem
---
"""


class Frontmatter:
    def __init__(self, lines) -> None:
        s = "\n".join(lines)
        try:
            self.yaml_data = yaml.safe_load(s)
        except:
            self.yaml_data = None

    def get_title(self) -> str:
        try:
            return self.yaml_data["title"]
        except:
            return ""

    def get_styles(self) -> str:
        try:
            lines = []
            styles = self.yaml_data.get("styles") or self.yaml_data.get("style")
            for style in styles:
                for sel, rules in style.items():
                    val = ";".join([f"{prop}:{v}" for prop, v in rules.items()])
                    line = sel + "{" + val + "}"
                    lines.append(line)
            return "\n".join(lines)
        except:
            return ""

    def get_css_variables(self) -> str:
        try:
            lines = [":root {"]
            vs = self.yaml_data.get("css-vars") or self.yaml_data.get("css-var")
            for k, v in vs.items():
                line = f"--{k}:{v};"
                lines.append(line)
            lines.append("}")
            return "\n".join(lines)
        except:
            return ""


class MdHtml:
    def __init__(self, file_path: str) -> None:
        raw_md = RawMd(file_path)
        md_lines = raw_md.get_lines()
        markup = mistletoe.markdown(md_lines.main, CustomRenderer)

        self._frontmatter = Frontmatter(md_lines.frontmatter)

        tree = DomTree(markup)
        tree.adjust_index("//*[contains(@class, 'force-order')]")
        tree.set_heading_id("h2 | h3 | h4 | h5 | h6")
        tree.fix_spacing("h2 | h3 | h4 | h5")
        tree.set_link_target()
        tree.set_timestamp(raw_md.get_timestamp())
        tree.render_pagebreak()
        tree.render_arrow_list()
        tree.render_blank_list()
        tree.render_pdflink()
        tree.render_td()
        tree.render_codeblock_label()
        tree.set_image_container()

        self._tree = tree

    @property
    def additional_style(self) -> str:
        return "<style>\n{}\n{}\n</style>".format(
            self._frontmatter.get_styles(),
            self._frontmatter.get_css_variables(),
        )

    @property
    def content(self) -> str:
        return self._tree.get_content()

    @property
    def toc(self) -> str:
        return '<div class="toc">{}</div>'.format(self._tree.get_toc())

    @property
    def title(self) -> str:
        t = self._frontmatter.get_title()
        if len(t) < 1:
            return self._tree.get_top_heading()
        return t


class HeadElem:
    def __init__(
        self, title: str, favicon_unicode: str, no_default_css: bool = False
    ) -> None:
        svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><text x="50%" y="50%" style="dominant-baseline:central;text-anchor:middle;font-size:90px;">&#x{};</text></svg>'.format(
            favicon_unicode
        )
        self.lines = [
            '<meta charset="utf-8">',
            '<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0">',
            "<title>{}</title>".format(title),
            '<link rel="icon" href="data:image/svg+xml,{}">'.format(
                urllib.parse.quote(svg)
            ),
        ]

        if not no_default_css:
            css_path = Path(__file__).with_name("markdown.css")
            self.append_style(css_path)

        self.append_elem(
            "<style>td.left{text-align:left;}td.center{text-align:center;}td.right{text-align:right;}</style>"
        )

    def append_style(self, css_path: pathlib.Path) -> None:
        elem = "<style>\n{}\n</style>".format(css_path.read_text("utf-8"))
        self.lines.append(elem)

    def append_elem(self, markup: str) -> None:
        self.lines.append(markup)

    def get_markup(self) -> str:
        return "\n".join(["<head>"] + self.lines + ["</head>"])


def main(
    file_path: str,
    no_default_css: bool = False,
    invoke: bool = False,
    favicon_unicode: str = "1F4DD",
) -> None:
    md_path = Path(file_path)
    if md_path.suffix != ".md":
        return

    md_html = MdHtml(str(md_path))

    head = HeadElem(
        (md_html.title or md_path.stem), favicon_unicode, no_default_css
    )
    extra_css_files = list(Path(md_path.parent).glob("*.css"))
    for css_path in extra_css_files:
        print("  + external style sheet: '{}'".format(css_path.name))
        head.append_elem(
            '<!-- from additional style sheet: "{}" -->'.format(css_path.name)
        )
        head.append_style(css_path)

    head.append_elem(md_html.additional_style)

    full_html = "\n".join(
        [
            "<!DOCTYPE html>",
            '<html lang="ja">',
            head.get_markup(),
            "<body>",
            "\n".join(
                [
                    '<div class="container">',
                    md_html.toc,
                    md_html.content,
                    "</div>",
                ]
            ),
            "</body>",
            "</html>",
        ]
    )

    ts = datetime.datetime.today().strftime(r"%Y%m%d")
    out_path = md_path.with_name(md_path.stem + "_" + ts + ".html")
    Path(out_path).write_text(full_html, "utf-8")

    if invoke:
        webbrowser.open(out_path)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("filePath", type=str)
    parser.add_argument("--faviconUnicode", default="1F4DD")
    parser.add_argument("--noDefaultCss", action="store_true")
    parser.add_argument("--invoke", action="store_true")
    args = parser.parse_args()
    main(args.filePath, args.noDefaultCss, args.invoke, args.faviconUnicode)
