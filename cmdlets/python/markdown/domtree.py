"""
https://lxml.de/apidoc/lxml.html
"""

import lxml.html

def decode(elem) -> str:
    return lxml.html.tostring(elem, encoding="unicode")

class DomTree:
    def __init__(self, markup:str="") -> None:
        self._root = lxml.html.fromstring(markup)

    def adjust_index(self, x_path:str) -> None:
        for elem in self._root.xpath(x_path):
            start_idx = elem.get("start") or 1
            counter = int(start_idx)
            for ol in elem.xpath("ol"):
                ol.set("start", str(counter))
                counter += len(ol.xpath("li"))

    def set_heading_id(self, x_path:str) -> None:
        elems = self._root.xpath(x_path)
        for i, hd in enumerate(elems):
            hd.set("id", "section-{}".format(i))

    def fix_spacing(self, x_path:str) -> None:
        for elem in self._root.xpath(x_path):
            t = elem.text_content()
            if t:
                l = len(t)
                if l >= 2 and l <= 4:
                    elem.classes.add("spacing-{}".format(l))

    def set_link_target(self) -> None:
        for elem in self._root.xpath("//a"):
            if not str(elem.get("href")).startswith("#"):
                elem.set("target", "_blank")
                elem.set("rel", "noopener noreferrer")

    def set_timestamp(self, ts:str) -> None:
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

    def get_content(self) -> str:
        self._root.classes.add("main")
        return decode(self._root)

    def get_top_heading(self) -> str:
        headers = self._root.xpath("h1 | h2")
        if len(headers) > 0:
            t = headers[0].text_content().strip()
            return t
        return ""
