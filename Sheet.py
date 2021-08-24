class Sheet:
    ID: str
    Title: str
    Rows: int
    Columns: []
    Color = dict()
    FR: int

    def __init__(self, title, cl, columns=[], rows=50, fr=1):
        self.Title = title
        self.Color = cl
        self.Rows = rows
        self.Columns = columns
        self.ID = None
        self.FR = fr