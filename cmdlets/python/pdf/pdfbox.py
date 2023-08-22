from pdfrw import PdfDict

class BoxParser:
    def __init__(self, rect:list[str]) -> None:
        self.bottom_left_x, self.bottom_left_y, self.top_right_x, self.top_right_y = [ float(x) for x in rect ]
        self.width = self.top_right_x - self.bottom_left_x
        self.height = self.top_right_y - self.bottom_left_y

class PdfBox:
    def __init__(self, page:PdfDict) -> None:
        mbox = page.inheritable.MediaBox
        self.media = BoxParser(mbox)
        if (cbox := page.inheritable.CropBox) is not None:
            self.visible = BoxParser(cbox)
        else:
            self.visible = BoxParser(mbox)
