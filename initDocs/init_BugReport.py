from Column import Column
from Sheet import Sheet
from Color import Color


def initDoc():
    return initSheets()


def initSheets():
    Sheets = []
    Columns = initColumns()
    Sheets.append(Sheet("Func", Color.Red))
    Sheets.append(Sheet("Layout", Color.Yellow))
    Sheets.append(Sheet("Design/Content", Color.Green))
    for sheet in Sheets:
        sheet.Columns = Columns
    return Sheets


def initColumns():
    Columns = [Column(0, "ID", 30), Column(1, "Description", 350), Column(2, "STR", 230),
               Column(3, "Platform", 90, ["ALL", "Desktop", "Adaptive", "Windows", "MacOS", "IOS", "Android"]),
               Column(4, "Browser", 90, ["ALL", "Safari", "Google Chrome", "Yandex Chrome", "MI", "Mozila", "Opera"]),
               Column(5, "Severity", 100,
                      ["S1 (blocker)", "S2 (critical)", "S3 (major)", "S4 (minor)", "S5 (trivial)"]),
               Column(6, "Status", 90, ["New", "Rejected", "Fixed", "Verified"]), Column(7, "Comment", 230),
               Column(8, "Feedback", 230)]
    return Columns
