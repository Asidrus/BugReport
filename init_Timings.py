from BugReport import *


def initDoc():
    return initSheets()


def initSheets():
    Sheets = []
    Columns = initColumns()
    for i in range(31):
        Sheets.append(Sheet(f"{i + 1}", color.White, rows=int(2 + 60 / 10 * 24), fr=2))
        Sheets.Columns = Columns
    return Sheets


def initColumns():
    Columns = [Column(0, "ID", 30),
               Column(1, "Description", 350),
               Column(2, "STR", 230),
               Column(3, "Platform", 90, ["ALL", "Desktop", "Adaptive", "Windows", "MacOS", "IOS", "Android"]),
               Column(4, "Browser", 90, ["ALL", "Safari", "Google Chrome", "Yandex Chrome", "MI", "Mozila", "Opera"]),
               Column(5, "Severity", 100,
                      ["S1 (blocker)", "S2 (critical)", "S3 (major)", "S4 (minor)", "S5 (trivial)"]),
               Column(6, "Status", 90, ["New", "Rejected", "Fixed", "Verified"]),
               Column(7, "Comment", 230),
               Column(8, "Feedback", 230)]
    return Columns

