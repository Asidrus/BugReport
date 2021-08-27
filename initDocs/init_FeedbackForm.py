from Column import Column
from Sheet import Sheet
from Color import Color


def initDoc():
    return initSheets()


def initSheets():
    Sheets = []
    Columns = initColumns()
    Sheets.append(Sheet("ФОС", Color.Red))
    for sheet in Sheets:
        sheet.Columns = Columns
    return Sheets


def initColumns():
    Columns = [Column(0, "ID", 30), Column(1, "", 30), Column(2, "Description", 230), Column(3, "Link", 120),
               Column(4, "Name", 90), Column(5, "Phone", 90), Column(6, "Email", 90),
               Column(7, "Platform", 90, ["Desktop", "IOS", "Android"]), Column(8, "Request time", 100),
               Column(9, "Status client", 100, "Checkbox"), Column(10, "Status CRM", 100, "Checkbox"),
               Column(11, "Comment", 230)]
    return Columns
