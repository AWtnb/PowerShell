import re
import html

from mistletoe.span_token import SpanToken
from mistletoe.block_token import BlockToken
from mistletoe.html_renderer import HTMLRenderer


"""

Renderer for block quote content

"""

class Tag(SpanToken):
    pattern = re.compile(r" #([^ ]+)")
    parse_group = 1

class QuoteSource(SpanToken):
    pattern = re.compile(r"~{(.+?)}")
    parse_group = 1

class QuoteRenderer(HTMLRenderer):

    def __init__(self):
        super().__init__(Tag, QuoteSource)

    def render_tag(self, token):
        template = '<span class="tag">{}</span>'
        return template.format(self.render_inner(token))

    def render_quote_source(self, token):
        template = '<div class="src">{}</div>'
        return template.format(self.render_inner(token))

    def render_line_break(self, _):
        """
        In block quote, all line break is rendered to '<br />'
        """
        return "<br />"


"""

Render with custom token

"""

class CheckBox(SpanToken):
    pattern = re.compile(r"(\[ *\]|\[x\])")
    parse_inner = False
    parse_group = 1
    def __init__(self, mo:re.Match):
        if mo.group(1) == "[x]":
            self.stat = "checked"
        else:
            self.stat = ""

class LineBreak(SpanToken):
    """
    For Japanese document, simple line break is removed through compile.
    If there is a single space at the end of line, lines are connected with single space to a single line.
    2 or more spaces at the end of line is rendered to line break.
    """
    pattern = re.compile(r"( *|\\)\n")
    parse_inner = False
    parse_group = 0

    def __init__(self, match):
        content = match.group(1)
        self.soft = not content.startswith(("  ", "\\"))
        self.has_space = content.startswith(" ")
        self.content = ''

class PageBreak(BlockToken):
    pattern = re.compile(r'^={3,}$')
    def __init__(self, _):
        pass

    @classmethod
    def start(cls, line):
        return cls.pattern.match(line)


"""

Main renderer for markdown text

"""

class CustomRenderer(HTMLRenderer):

    quote_renderer = QuoteRenderer()

    def __init__(self):
        super().__init__(CheckBox, LineBreak, PageBreak)
        self.page_break = '<div class="page-separator"></div>'

    def render_check_box(self, token):
        template = '<input type="checkbox" disabled {stat}>'
        return template.format(stat=token.stat)

    def render_quote(self, token):
        elements = ['<blockquote>']
        self._suppress_ptag_stack.append(False)
        elements.extend([self.quote_renderer.render(child) for child in token.children])
        self._suppress_ptag_stack.pop()
        elements.append('</blockquote>')
        return '\n'.join(elements)

    def render_page_break(self, _) -> str:
        return self.page_break

    def render_heading(self, token):
        inner = self.render_inner(token)
        if inner.strip()[:3] in ["===", "---", "///"]:
            return self.page_break
        template = '<h{level}>{inner}</h{level}>'
        return template.format(level=token.level, inner=inner)

    def render_list_item(self, token):
        if len(token.children) == 0:
            return '<li class="empty"></li>'
        inner = '\n'.join([self.render(child) for child in token.children])
        inner_template = '\n{}\n'
        if self._suppress_ptag_stack[-1]:
            if token.children[0].__class__.__name__ == 'Paragraph':
                inner_template = inner_template[1:]
            if token.children[-1].__class__.__name__ == 'Paragraph':
                inner_template = inner_template[:-1]
        inner_html = inner_template.format(inner)
        if inner_html.startswith("=&gt;"):
            attr = ' class="sub"'
            inner_html = inner_html[5:]
        else:
            attr = ""
        return '<li{}>{}</li>'.format(attr, inner_html)

    def render_table_cell(self, token, in_header=False):
        template = '<{tag}{attr}>{inner}</{tag}>\n'
        tag = 'th' if in_header else 'td'
        if token.align is None:
            align = 'left'
        elif token.align == 0:
            align = 'center'
        elif token.align == 1:
            align = 'right'
        attr = ' class="{}"'.format(align)
        inner = self.render_inner(token)
        return template.format(tag=tag, attr=attr, inner=inner)

    def render_block_code(self, token):
        template = '<pre{attr}><code>{inner}</code></pre>'
        if token.language:
            attr = ' class="codeblock-header" data-label="{}"'.format(html.escape(token.language))
        else:
            attr = ''
        inner = html.escape(token.children[0].content)
        return template.format(attr=attr, inner=inner)

    def render_link(self, token):
        template = '<a href="{dest}"{title}{attr}>{inner}</a>'
        dest = self.escape_url(token.target)
        if dest.lower().endswith(".pdf"):
            attr = ' filetype="pdf" '
        else:
            attr = ''
        if token.title:
            title = ' title="{}"'.format(html.escape(token.title))
        else:
            title = ''
        inner = self.render_inner(token)
        return template.format(dest=dest, title=title, attr=attr, inner=inner)

    def render_line_break(self, token):
        if token.soft:
            return "\n" if token.has_space else ""
        return "<br />"