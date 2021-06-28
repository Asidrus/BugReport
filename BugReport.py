import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
from time import sleep
import os


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
            self.DataValidation = list()
            for data in datavalidation:
                self.DataValidation.append({"userEnteredValue": data})


class Sheet:
    ID: str
    Title: str
    Color = dict()

    def __init__(self, title, cl):
        self.Title = title
        self.Color = cl
        self.ID = None


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

    def initColumns(self):
        self.__Columns__.append(Column(0, "ID", 30))
        self.__Columns__.append(Column(1, "Description", 350))
        self.__Columns__.append(Column(2, "STR", 230))
        self.__Columns__.append(Column(3, "Platform", 90, ["ALL", "Desktop", "Adaptive", "Windows", "MacOS", "IOS", "Android"]))
        self.__Columns__.append(
            Column(4, "Browser", 90, ["ALL", "Safari", "Google Chrome", "Yandex Chrome", "MI", "Mozila", "Opera"]))
        self.__Columns__.append(
            Column(5, "Severity", 100, ["S1 (blocker)", "S2 (critical)", "S3 (major)", "S4 (minor)", "S5 (trivial)"]))
        self.__Columns__.append(Column(6, "Status", 90, ["New", "Rejected", "Fixed", "Verified"]))
        self.__Columns__.append(Column(7, "Comment", 230))
        self.__Columns__.append(Column(8, "Feedback", 230))

    def initSheets(self):
        self.__Sheets__.append(Sheet("Func", color.Red))
        self.__Sheets__.append(Sheet("Layout", color.Yellow))
        self.__Sheets__.append(Sheet("Design/Content", color.Green))

    def RomanovskayaChanges(self):
        self.__Columns__[5] = Column(5, "Severity", 100, ["S0 (Simple)", "S1 (blocker)", "S2 (critical)", "S3 (major)", "S4 (minor)", "S5 (trivial)"])
        self.__Columns__[6] = Column(6, "Status", 90, ["New", "Rejected", "Fixed", "Verified", "Check"])


    def __init__(self, spreadsheet_id, CREDENTIALS_FILE='stable-ring-316114-8acf36454762.json', for_Romanovskaya = False):
        self.SSID = spreadsheet_id
        self.CREDENTIALS_FILE = CREDENTIALS_FILE
        credentials = ServiceAccountCredentials.from_json_keyfile_name(self.CREDENTIALS_FILE,
                                                                       ['https://www.googleapis.com/auth/spreadsheets',
                                                                        'https://www.googleapis.com/auth/drive'])
        httpAuth = credentials.authorize(httplib2.Http())
        self.service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)
        self.initColumns()
        self.initSheets()
        if for_Romanovskaya:
            self.RomanovskayaChanges()

    def getSheets(self):
        for sheet in self.service.spreadsheets().get(spreadsheetId=self.SSID).execute()["sheets"]:
            for _Sheet in self.__Sheets__:
                if sheet["properties"]["title"] == _Sheet.Title:
                    _Sheet.ID = sheet["properties"]["sheetId"]

    def getBug(self, ID):
        range = f"A{ID + 1}:I{ID + 1}"
        p = self.service.spreadsheets().values().get(spreadsheetId=self.SSID, range=range,
                                                     majorDimension='ROWS').execute()["values"][0]
        bug = Bug(p[0], p[1], p[2], p[3], p[4], p[5], p[6], p[7], p[8])
        return bug

    def getLength(self, sheet):
        values = self.service.spreadsheets().values().get(spreadsheetId=self.SSID,
                                                          range=f"{sheet['Title']}!A2:A{self.__MaxBugs__ + 1}",
                                                          majorDimension='COLUMNS').execute()
        try:
            return len(values['values'][0])
        except:
            return 0

    def addBug(self):
        pass

    def setSheets(self):
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

    def setHeaderStyle(self, sheet):
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
                            "gridProperties": {"frozenRowCount": 1}
                        },
                        "fields": "gridProperties.frozenRowCount"
                    }
                }
            ]
        })
        # width,datavalidation
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

    def setLineStyle(self, sheet):
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

    def doDoc(self):
        self.setSheets()
        for sheet in self.__Sheets__:
            self.setHeaderStyle(sheet)
            self.setLineStyle(sheet)

    def SendRequest(self, body):
        while True:
            try:
                res = self.service.spreadsheets().batchUpdate(spreadsheetId=self.SSID, body=body).execute()
                break
            except:
                for i in range(100, 0, -1):
                    os.system("cls")
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


class Bug:
    ID: int
    Description: str
    STR: dict
    Platform: str
    Browser: str
    Priority: str
    Status: str
    Comment: str
    Feedback: str

    def __init__(self):
        pass

    def __init__(self, _ID, _Description, _STR, _Platform, _Browser, _Priority, _Status, _Comment, _Feedback):
        self.ID = _ID
        self.Description = _Description
        self.STR = _STR
        self.Platform = _Platform
        self.Browser = _Browser
        self.Priority = _Priority
        self.Status = _Status
        self.Comment = _Comment
        self.Feedback = _Feedback

    def __str__(self):
        return f"ID = {self.ID}; Des = {self.Description}; Status: {self.Status}"

    def push(self, BR: BugReport, SheetID):
        data = [
            [self.id, self.Description, STRdict2str(self.STR), self.Platform, self.Browser, self.Status, self.Comment]]
        range = f"A{self.ID + 1}:H{self.ID + 1}"
        values = BR.service.spreadsheets().values().batchUpdate(
            spreadsheetId=SheetID,
            body={
                "valueInputOption": "USER_ENTERED",
                "data": [
                    {"range": range,
                     "majorDimension": "ROWS",
                     "values": [["This is B3", "This is C3"], ["This is B4", "This is C4"]]},
                    {"range": "D5:E6",
                     "majorDimension": "COLUMNS",
                     "values": [["This is D5", "This is D6"], ["This is E5", "=5+5"]]}
                ]
            }
        ).execute()
