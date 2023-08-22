from pathlib import Path
import sys

from pdfrw import PdfReader, PdfWriter, PageMerge

from pdfbox import PdfBox

from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

class FileNameWatermark:

    def __init__(self, work_dir:str, file_path:str) -> None:
        self.target_path = Path(file_path)
        self.target_pdf = PdfReader(str(self.target_path))
        self.watermark_path = Path(work_dir, "wm_for_"+self.target_path.name)
        self.watermark_pdf = None

    def create(self, start_idx:int=1) -> None:
        try:
            c = canvas.Canvas(str(self.watermark_path))
            font_size = 12
            pdfmetrics.registerFont(TTFont("localfont", r"C:\Windows\Fonts\msgothic.ttc"))
            for i, page in enumerate(self.target_pdf.pages):
                box = PdfBox(page)
                visible_box = box.visible
                c.setPageSize(tuple([box.media.width, box.media.height]))
                watermark_text = "  {}(p.{})".format(self.target_path.stem, i + start_idx)
                c.setFont('localfont', font_size)
                c.setFillColor(Color(red=35/255, green=90/255, blue=150/255))
                c.rotate(90)
                c.drawString(visible_box.bottom_left_y, 0 - visible_box.bottom_left_x - font_size, watermark_text*30)
                c.showPage()
            c.save()
            self.watermark_pdf = PdfReader(str(self.watermark_path))
        except Exception as e:
            print(e, file=sys.stderr)

    def overlay(self) -> str:
        if self.watermark_pdf:
            out_path = str(self.target_path.with_stem("wm_" + self.target_path.stem))
            for i, page in enumerate(self.target_pdf.pages):
                watermark = PageMerge().add(self.watermark_pdf.pages[i])[0]
                PageMerge(page).add(watermark, prepend=False).render()
            PdfWriter(out_path, trailer=self.target_pdf).write()
            return out_path
        else:
            print("ERROR: failed to watermarking on '{}'".format(self.target_path.name), file=sys.stderr)
            return ""
