"""
https://lxml.de/apidoc/lxml.html
"""

import lxml.html

def decode(elem) -> str:
    return lxml.html.tostring(elem, encoding="unicode")

class DomTree:
    def __init__(self, markup:str="", timestamp:str="") -> None:
        self._root = lxml.html.fromstring(markup)
        self._adjust_index("//*[contains(@class, 'force-order')]")
        self._set_heading_id("h2 | h3 | h4 | h5 | h6")
        self._fix_spacing("h2 | h3 | h4 | h5")
        self._set_link_target()
        self._set_timestamp(timestamp)

    def _adjust_index(self, x_path:str) -> None:
        for elem in self._root.xpath(x_path):
            start_idx = elem.get("start") or 1
            counter = int(start_idx)
            for ol in elem.xpath("ol"):
                ol.set("start", str(counter))
                counter += len(ol.xpath("li"))

    def _set_heading_id(self, x_path:str) -> None:
        elems = self._root.xpath(x_path)
        for i, hd in enumerate(elems):
            hd.set("id", "section-{}".format(i))

    def _fix_spacing(self, x_path:str) -> None:
        for elem in self._root.xpath(x_path):
            t = elem.text_content()
            if t:
                l = len(t)
                if l >= 2 and l <= 4:
                    elem.classes.add("spacing-{}".format(l))

    def _set_link_target(self) -> None:
        for elem in self._root.xpath("//a"):
            if not str(elem.get("href")).startswith("#"):
                elem.set("target", "_blank")
                elem.set("rel", "noopener noreferrer")

    def _set_timestamp(self, ts:str) -> None:
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
                max_width = 16 - int(hd.tag[1:]) - 2
                if max_width < len(a.text):
                    a.text = a.text[:(max_width - 1)] + "..."
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
