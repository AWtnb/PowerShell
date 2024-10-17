"""
https://lxml.de/apidoc/lxml.html
"""

import lxml.html


def decode(elem) -> str:
    return lxml.html.tostring(elem, encoding="unicode")


class DomTree:
    def __init__(self, markup: str = "") -> None:
        self._root = lxml.html.fromstring(markup)

    def adjust_index(self, x_path: str) -> None:
        for elem in self._root.xpath(x_path):
            start_idx = elem.get("start") or 1
            counter = int(start_idx)
            for ol in elem.xpath("ol"):
                ol.set("start", str(counter))
                counter += len(ol.xpath("li"))

    def set_heading_id(self, x_path: str) -> None:
        elems = self._root.xpath(x_path)
        for i, hd in enumerate(elems):
            hd.set("id", "section-{}".format(i))

    def render_arrow_list(self) -> None:
        elems = self._root.xpath("//li")
        for l in elems:
            if str(l.text).startswith("=>"):
                l.text = l.text[3:]
                l.classes.add("sub")

    def render_blank_list(self) -> None:
        elems = self._root.xpath("//li")
        for l in elems:
            if len(l.text_content()) < 1:
                l.classes.add("empty")

    def render_pagebreak(self) -> None:
        elems = self._root.xpath("p")
        for p in elems:
            if str(p.text).startswith("==="):
                p.clear()
                p.classes.add("page-separator")

    def render_pdflink(self) -> None:
        elems = self._root.xpath("//a")
        for a in elems:
            if str(a.get("href")).endswith(".pdf"):
                a.set("filetype", "pdf")

    def render_td(self) -> None:
        elems = self._root.xpath("//td")
        for td in elems:
            if (al := td.get("align")) is not None:
                td.classes.add(al)

    def render_codeblock_label(self) -> None:
        elems = self._root.xpath("//pre/code")
        for bl in elems:
            if (cl := bl.classes) and len(cl):
                n = list(cl)[0][len("language-") :]
                if (p := bl.getparent()) is not None:
                    p.classes.add("codeblock-header")
                    p.set("data-label", n)

    def fix_spacing(self, x_path: str) -> None:
        for elem in self._root.xpath(x_path):
            t = elem.text_content()
            if t:
                l = len(t)
                if l >= 2 and l <= 4:
                    elem.classes.add("spacing-{}".format(l))

    def set_image_container(self) -> None:
        for elem in self._root.xpath("//p"):
            if elem.xpath("img"):
                alt = elem.xpath("img")[0].get("alt", "center")
                container = lxml.html.Element("div")
                container.classes.add("img-container")
                container.set("pos", alt)
                wrapper = lxml.html.Element("div")
                wrapper.classes.add("img-wrapper")
                for c in elem.getchildren():
                    wrapper.append(c)
                container.append(wrapper)
                self._root.replace(elem, container)

    def set_link_target(self) -> None:
        for elem in self._root.xpath("//a"):
            if not str(elem.get("href")).startswith("#"):
                elem.set("target", "_blank")
                elem.set("rel", "noopener noreferrer")

    def set_timestamp(self, ts: str) -> None:
        div = lxml.html.Element("div")
        div.classes.add("timestamp")
        div.text = ts
        self._root.insert(0, div)

    def get_toc(self) -> str:
        toc = ""
        headers = self._root.xpath("h2 | h3 | h4 | h5 | h6")
        if len(headers) > 0:
            ul = lxml.html.Element("ul")
            for hd in headers:
                a = lxml.html.Element("a")
                a.set("href", "#{}".format(hd.get("id") or ""))
                a.text = hd.text_content()
                li = lxml.html.Element("li")
                li.classes.add("toc-{}".format(hd.tag))
                li.append(a)
                ul.append(li)
            toc = decode(ul)
        return toc

    def get_comments(self) -> list[str]:
        comments = self._root.xpath("//comment()")
        return [c.text.strip() for c in comments]

    def get_content(self) -> str:
        self._root.classes.add("main")
        return decode(self._root)

    def get_top_heading(self) -> str:
        headers = self._root.xpath("h1 | h2")
        if len(headers) > 0:
            t = headers[0].text_content().strip()
            return t
        return ""

    def trim_leading_css_block(self) -> str:
        headers = 0
        css = ""
        for elem in self._root.xpath("pre[@data-label='css']/code | h1 | h2"):
            if elem.tag in ("h1", "h2"):
                headers += 1
                continue
            if headers < 1 and elem.tag == "code":
                css = elem.text_content()
                elem.getparent().drop_tree()
                break
        return css
