from pdfrw import PdfDict

class BoxParser:
    def __init__(self, rect:list[str]) -> None:
        self.rect = [ float(x) for x in rect ]
        self.bottom_left_x = self.rect[0]
        self.bottom_left_y = self.rect[1]
        self.top_right_x = self.rect[2]
        self.top_right_y = self.rect[3]
        self.width = self.top_right_x - self.bottom_left_x
        self.height = self.top_right_y - self.bottom_left_y

class PdfPageBox:
    def __init__(self, page:PdfDict) -> None:
        mbox = page.inheritable.MediaBox
        self.media = BoxParser(mbox)
        if (cbox := page.inheritable.CropBox) is not None:
            self.visible = BoxParser(cbox)
        else:
            self.visible = BoxParser(mbox)
