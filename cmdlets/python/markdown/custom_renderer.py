import re

from mistletoe.span_token import SpanToken
from mistletoe.html_renderer import HTMLRenderer


"""

Renderer for block quote content

"""

# class Tag(SpanToken):
#     pattern = re.compile(r" #([^ ]+)")
#     parse_group = 1

class QuoteRenderer(HTMLRenderer):

    def __init__(self):
        # super().__init__(Tag)
        super().__init__()

    # def render_tag(self, token):
    #     template = '<span class="tag">{}</span>'
    #     return template.format(self.render_inner(token))

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


"""

Main renderer for markdown text

"""

class CustomRenderer(HTMLRenderer):

    quote_renderer = QuoteRenderer()

    def __init__(self):
        super().__init__(CheckBox, LineBreak)

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

    def render_line_break(self, token):
        if token.soft:
            return "\n" if token.has_space else ""
        return "<br />"