from Color import Color


class Sheet:
    ID: str
    Title: str
    Rows: int
    Columns: []
    TabColor = dict()
    FR: int
    HeaderColor = Color.Cyan
    LineColor = {0: {"fore": Color.Black, "back": Color.White},
                 1: {"fore": Color.Black, "back": Color.LightGray}}

    def __init__(self, title, cl, columns=[], rows=50, fr=1):
        self.Title = title
        self.TabColor = cl
        self.Rows = rows
        self.Columns = columns
        self.ID = None
        self.FR = fr
