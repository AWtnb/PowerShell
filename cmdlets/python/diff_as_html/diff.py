import argparse
from pathlib import Path
from difflib import HtmlDiff


class PyDiff:

    def __init__(self, from_path:str, to_path:str) -> None:
        f_path = Path(from_path)
        t_path = Path(to_path)
        f = f_path.read_text("utf-8").splitlines()
        t = t_path.read_text("utf-8").splitlines()
        df = HtmlDiff()
        self.markup = df.make_table(f, t, fromdesc=f_path.name, todesc=t_path.name)


    def compress_markup(self) -> None:
        markup_lines = self.markup.splitlines()
        pre = markup_lines[:7]
        trs = markup_lines[7:-2]
        post = markup_lines[-2:]
        filler = '<tr class="filler"><td class="diff_next"></td><td class="diff_header"></td><td nowrap="nowrap"></td><td class="diff_next"></td><td class="diff_header"></td><td nowrap="nowrap"></td></tr>'
        minimal_lines = []
        for i, line in enumerate(trs):
            if ('class="diff_chg"' in line) or ('class="diff_sub"' in line) or ('class="diff_add"' in line):
                minimal_lines.append([-1, filler])
                minimal_lines.append([i, line])
                if i == 0:
                    minimal_lines.pop(-2)
                if len(minimal_lines) > 2 and minimal_lines[-3][0] + 1 == i:
                    minimal_lines.pop(-2)
        self.markup = "\n".join(pre + [x[1] for x in minimal_lines] + post)

def main(from_file:str, to_file:str, out_path:str, template_path:str, css_path:str, skip_unchanged:bool) -> None:
    pd = PyDiff(from_file, to_file)
    if skip_unchanged:
        pd.compress_markup()
    style_sheet = Path(css_path).read_text("utf-8")
    template = Path(template_path).read_text("utf-8")
    html_page = template.replace(
        "<style></style>", "<style>\n{}\n</style>".format(style_sheet)
    ).replace(
        '<div class="main"></div>', '<div class="main">\n{}\n</div>'.format(pd.markup)
    )
    Path(out_path).write_text(html_page, "utf-8")
    print("compared '{}' -> '{}' as '{}'.".format(
        Path(from_file).name,
        Path(to_file).name,
        Path(out_path).name))

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("fromFile", type=str)
    parser.add_argument("toFile", type=str)
    parser.add_argument("outFile", type=str)
    parser.add_argument("--compress", action="store_true")
    args = parser.parse_args()

    css_path = Path(__file__).with_name("wrap.css")
    template_path = Path(__file__).with_name("template.html")
    main(args.fromFile, args.toFile, args.outFile, template_path, css_path, args.compress)