from Column import Column
from Sheet import Sheet
from Color import Color


def initDoc():
    return initSheets()


def initSheets():
    Sheets = []
    Columns = initColumns()
    for i in range(31):
        Sheets.append(Sheet(f"{i + 1}", Color.White,
                            columns=Columns,
                            rows=int(2 + 60 / 10 * 24),
                            fr=2))
    return Sheets


def initColumns():
    Columns = [Column(0, "DateTime", 150), Column(1, "Страница логина", 150), Column(2, "Редирект", 150),
               Column(3, "Страница ЛК", 150), Column(4, "Войти в модуль", 150), Column(5, "Лекция", 150),
               Column(6, "Видео", 150), Column(7, "Тест вопросы", 150), Column(8, "Практическое", 150),
               Column(9, "Итоговое тест", 150), Column(10, "Страница ЛК", 150), Column(11, "Календарь", 150)]
    return Columns
