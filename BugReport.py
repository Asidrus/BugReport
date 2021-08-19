import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
from time import sleep
import os
from datetime import datetime


class color:
    White = {"red": 1, "green": 1, "blue": 1}
    Black = {"red": 0, "green": 0, "blue": 0}
    LightGray = {"red": 0.95, "green": 0.95, "blue": 0.95}
    Red = {"red": 1, "green": 0, "blue": 0}
    Green = {"red": 0, "green": 1, "blue": 0}
    Blue = {"red": 0, "green": 0, "blue": 1}
    Cyan = {"red": 0, "green": 1, "blue": 1}
    Yellow = {"red": 1, "green": 1, "blue": 0}


class Column:
    ID: int
    Title: str
    Width: int
    DataValidation = None

    def __init__(self, id, title, width, datavalidation=None):
        self.ID = id
        self.Title = title
        self.Width = width
        if datavalidation is not None:
            if datavalidation == "Checkbox":
                self.DataValidation = datavalidation
            else:
                self.DataValidation = list()
                for data in datavalidation:
                    self.DataValidation.append({"userEnteredValue": data})


class Sheet:
    ID: str
    Title: str
    Color = dict()
    FR: int

    def __init__(self, title, cl, fr=1):
        self.Title = title
        self.Color = cl
        self.ID = None
        self.FR = fr


class BugReport:
    __MaxBugs__ = 50
    __Columns__ = list()
    __Sheets__ = list()
    __HeaderColor__ = color.Cyan
    __LinesColor__ = {0: {"fore": color.Black, "back": color.White},
                      1: {"fore": color.Black, "back": color.LightGray}}

    SSID: str
    CREDENTIALS_FILE: str
    ServiceAccount = 'acc-340@stable-ring-316114.iam.gserviceaccount.com'

    def initColumns(self, type):
        if type == "br":
            self.__Columns__.append(Column(0, "ID", 30))
            self.__Columns__.append(Column(1, "Description", 350))
            self.__Columns__.append(Column(2, "STR", 230))
            self.__Columns__.append(
                Column(3, "Platform", 90, ["ALL", "Desktop", "Adaptive", "Windows", "MacOS", "IOS", "Android"]))
            self.__Columns__.append(
                Column(4, "Browser", 90, ["ALL", "Safari", "Google Chrome", "Yandex Chrome", "MI", "Mozila", "Opera"]))
            self.__Columns__.append(
                Column(5, "Severity", 100,
                       ["S1 (blocker)", "S2 (critical)", "S3 (major)", "S4 (minor)", "S5 (trivial)"]))
            self.__Columns__.append(Column(6, "Status", 90, ["New", "Rejected", "Fixed", "Verified"]))
            self.__Columns__.append(Column(7, "Comment", 230))
            self.__Columns__.append(Column(8, "Feedback", 230))
        elif type == "fbf":
            self.__Columns__.append(Column(0, "ID", 30))
            self.__Columns__.append(Column(1, "", 30))
            self.__Columns__.append(Column(2, "Description", 230))
            self.__Columns__.append(Column(3, "Link", 120))
            self.__Columns__.append(Column(4, "Name", 90))
            self.__Columns__.append(Column(5, "Phone", 90))
            self.__Columns__.append(Column(6, "Email", 90))
            self.__Columns__.append(Column(7, "Platform", 90, ["Desktop", "IOS", "Android"]))
            self.__Columns__.append(Column(8, "Request time", 100))
            self.__Columns__.append(Column(9, "Status client", 50, "Checkbox"))
            self.__Columns__.append(Column(10, "Status CRM", 50, "Checkbox"))
            self.__Columns__.append(Column(11, "Comment", 230))
        elif type == "times":
            self.__Columns__.append(Column(0, "DateTime", 30))
            self.__Columns__.append(Column(1, "Страница логина", 100))
            self.__Columns__.append(Column(2, "Страница ЛК", 100))
            self.__Columns__.append(Column(3, "Войти в модуль", 100))
            self.__Columns__.append(Column(4, "Лекция", 100))
            self.__Columns__.append(Column(5, "Видео", 100))
            self.__Columns__.append(Column(6, "Тест вопросы", 100))
            self.__Columns__.append(Column(7, "Практическое", 100))
            self.__Columns__.append(Column(8, "Итоговое тест", 100))
            self.__Columns__.append(Column(9, "Страница ЛК", 100))
            self.__Columns__.append(Column(10, "Календарь", 100))

    def initSheets(self, type):
        if type == "br":
            self.__Sheets__.append(Sheet("Func", color.Red))
            self.__Sheets__.append(Sheet("Layout", color.Yellow))
            self.__Sheets__.append(Sheet("Design/Content", color.Green))
        elif type == "fbf":
            self.__Sheets__.append(Sheet("Niidpo", color.Red))
        elif type == "time":
            for i in range(30):
                self.__Sheets__.append(f"{i + 1}", color.White, FR=2)

    def RomanovskayaChanges(self):
        self.__Columns__[5] = Column(5, "Severity", 100,
                                     ["S0 (Simple)", "S1 (blocker)", "S2 (critical)", "S3 (major)", "S4 (minor)",
                                      "S5 (trivial)"])
        self.__Columns__[6] = Column(6, "Status", 90, ["New", "Rejected", "Fixed", "Verified", "Check"])

    def __init__(self, spreadsheet_id, CREDENTIALS_FILE='stable-ring-316114-8acf36454762.json'):
        self.SSID = spreadsheet_id
        self.CREDENTIALS_FILE = CREDENTIALS_FILE
        credentials = ServiceAccountCredentials.from_json_keyfile_name(self.CREDENTIALS_FILE,
                                                                       ['https://www.googleapis.com/auth/spreadsheets',
                                                                        'https://www.googleapis.com/auth/drive'])
        httpAuth = credentials.authorize(httplib2.Http())
        self.service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)

    def getSheets(self):
        for sheet in self.service.spreadsheets().get(spreadsheetId=self.SSID).execute()["sheets"]:
            for _Sheet in self.__Sheets__:
                if sheet["properties"]["title"] == _Sheet.Title:
                    _Sheet.ID = sheet["properties"]["sheetId"]

    def setSheets(self, type):
        self.getSheets()
        for sheet in self.__Sheets__:
            if sheet.ID is None:
                res = self.SendRequest({
                    "requests": [
                        {
                            "addSheet": {
                                "properties": {
                                    "title": sheet.Title,
                                    "gridProperties": {
                                        "rowCount": self.__MaxBugs__ + 1,
                                        "columnCount": len(self.__Columns__)
                                    },
                                    "tabColor": sheet.Color
                                }
                            }
                        }
                    ]
                })
        self.getSheets()

    def setHeaderStyle(self, sheet, type):
        # titles
        res = self.service.spreadsheets().values().update(spreadsheetId=self.SSID,
                                                          range=f"{sheet.Title}!A1:{ind2str(len(self.__Columns__))}1",
                                                          valueInputOption="USER_ENTERED",
                                                          body={"values": [
                                                              [column.Title for column in self.__Columns__]]}).execute()

        # backgroundColor, textFormat, horizontalAlignment
        res = self.SendRequest({
            "requests": [
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet.ID,
                            "startRowIndex": 0,
                            "endRowIndex": 1
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": self.__HeaderColor__,
                                "horizontalAlignment": "CENTER",
                                "textFormat": {
                                    "foregroundColor": color.Black,
                                    "fontSize": 12,
                                    "bold": True
                                }
                            }
                        },
                        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
                    }
                },
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": sheet.ID,
                            "gridProperties": {"frozenRowCount": sheet.FR}
                        },
                        "fields": "gridProperties.frozenRowCount"
                    }
                }
            ]
        })
        # width, datavalidation
        for column in self.__Columns__:
            body = {
                "requests": [
                    {
                        "updateDimensionProperties": {
                            "range": {
                                "sheetId": sheet.ID,
                                "dimension": "COLUMNS",
                                "startIndex": column.ID,
                                "endIndex": column.ID + 1
                            },
                            "properties": {
                                "pixelSize": column.Width
                            },
                            "fields": "pixelSize"
                        }
                    }
                ]
            }
            if column.DataValidation is not None:
                if column.DataValidation == "Checkbox":
                    body["requests"].append({
                        "setDataValidation": {
                            "range": {
                                "sheetId": sheet.ID,
                                "startRowIndex": 1,
                                "endRowIndex": self.__MaxBugs__ + 1,
                                "startColumnIndex": column.ID,
                                "endColumnIndex": column.ID + 1
                            },
                            "rule": {
                                "condition": {
                                    "type": "BOOLEAN"
                                },
                                "showCustomUi": True,
                                "strict": True
                            }
                        }
                    })
                else:
                    body["requests"].append({
                        "setDataValidation": {
                            "range": {
                                "sheetId": sheet.ID,
                                "startRowIndex": 1,
                                "endRowIndex": self.__MaxBugs__ + 1,
                                "startColumnIndex": column.ID,
                                "endColumnIndex": column.ID + 1
                            },
                            "rule": {
                                "condition": {
                                    "type": "ONE_OF_LIST",
                                    "values": column.DataValidation
                                },
                                "showCustomUi": True,
                                "strict": True
                            }
                        }
                    })
            self.SendRequest(body)

    def setLineStyle(self, sheet, type):
        # lines
        for i in range(self.__MaxBugs__):
            color = self.__LinesColor__[i % 2]
            res = self.SendRequest({
                "requests": [
                    {
                        "repeatCell": {
                            "range": {
                                "sheetId": sheet.ID,
                                "startRowIndex": i + 1,
                                "endRowIndex": i + 2
                            },
                            "cell": {
                                "userEnteredFormat": {
                                    "backgroundColor": color["back"],
                                    "verticalAlignment": "TOP",
                                    "horizontalAlignment": "LEFT",
                                    "wrapStrategy": "WRAP",
                                    "textFormat": {
                                        "foregroundColor": color["fore"],
                                        "fontSize": 10,
                                    }
                                }
                            },
                            "fields": "userEnteredFormat(backgroundColor,wrapStrategy,textFormat,horizontalAlignment,"
                                      "verticalAlignment)"
                        }
                    },
                ]
            })
        if type == "times":
            res = self.SendRequest({
                "requests": [{
                    "setDataValidation": {
                        "range": {
                            "sheetId": sheet.ID,
                            "startRowIndex": 2,
                            "endRowIndex": 3,
                        },
                        "rule": {
                            "condition": {
                                "type": "BOOLEAN"
                            },
                            "showCustomUi": True,
                            "strict": True
                        }
                    }
                }]})

    def doDoc(self, type, for_Romanovskaya=False):
        self.initColumns(type)
        self.initSheets(type)
        if for_Romanovskaya:
            self.RomanovskayaChanges()
        self.setSheets(type)
        for sheet in self.__Sheets__:
            self.setHeaderStyle(sheet, type)
            self.setLineStyle(sheet, type)

    def SendRequest(self, body):
        while True:
            try:
                res = self.service.spreadsheets().batchUpdate(spreadsheetId=self.SSID, body=body).execute()
                break
            except Exception as e:
                for i in range(100, 0, -1):
                    os.system("cls")
                    print(e)
                    print("Превышен лимит ожидания")
                    print(f"Повторная отправка через {i} сек")
                    sleep(1)
        return res


def STRdict2str(STR: dict):
    res = """"""
    for index in STR.keys():
        res = res + """
        """ + index + ". " + STR[index]
    return res


def ind2str(ind: int):
    return chr(65 + ind)
