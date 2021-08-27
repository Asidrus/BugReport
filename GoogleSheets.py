import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
from time import sleep
import os
import initDocs as Doc
from Color import Color
from GenRequests import *

ind2str = lambda ind: chr(65 + ind)
path = os.path.abspath(os.getcwd())


class BugReport:
    __Sheets__ = list()
    SSID: str
    CREDENTIALS_FILE: str
    TypeOfDoc: str

    def __init__(self, SSID, typeOfDoc, CREDENTIALS_FILE=path + 'stable-ring-316114-8acf36454762.json'):
        self.SSID = SSID
        credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE,
                                                                       ['https://www.googleapis.com/auth/spreadsheets',
                                                                        'https://www.googleapis.com/auth/drive'])
        httpAuth = credentials.authorize(httplib2.Http())
        self.service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)
        self.TypeOfDoc = typeOfDoc
        self.__Sheets__ = Doc.init(typeOfDoc)

    def getSheets(self):
        for sheet in self.service.spreadsheets().get(spreadsheetId=self.SSID).execute()["sheets"]:
            for _Sheet in self.__Sheets__:
                if sheet["properties"]["title"] == _Sheet.Title:
                    _Sheet.ID = sheet["properties"]["sheetId"]

    def setSheets(self):
        self.getSheets()
        for sheet in self.__Sheets__:
            if sheet.ID is None:
                self.SendRequest({
                    "requests": [
                        {
                            "addSheet": {
                                "properties": {
                                    "title": sheet.Title,
                                    "gridProperties": {
                                        "rowCount": sheet.Rows,
                                        "columnCount": len(sheet.Columns)
                                    },
                                    "tabColor": sheet.TabColor
                                }
                            }
                        }
                    ]
                })
        self.getSheets()

    def addData(self, sheet, data):
        raw = self.service.spreadsheets().values().get(spreadsheetId=self.SSID,
                                                       range=f"{sheet.Title}!A1:{ind2str(len(self.__Columns__))}{self.__MaxBugs__}",
                                                       majorDimension='ROWS').execute()["values"]
        ind = len(raw) + 1
        cover = lambda t, l: f"=if({l}2,1,0)*\"{t}\""
        lines = []
        for dat in data:
            line = []
            for i in range(len(dat)):
                if i == 0:
                    line.append(dat[i])
                    continue
                line.append(cover(dat[i], ind2str(i)))
            lines.append(line)

        res = self.service.spreadsheets().values().update(spreadsheetId=self.SSID,
                                                          range=f"{sheet.Title}!A{ind}:{ind2str(len(lines[0]))}{ind + len(lines)}",
                                                          valueInputOption="USER_ENTERED",
                                                          body={"values": lines}).execute()

    def doDoc(self):
        self.getSheets()
        self.setSheets()
        for sheet in self.__Sheets__:
            # Update Column Title
            self.UpdateValue(range=f"{sheet.Title}!A1:{ind2str(len(sheet.Columns))}1",
                             values=[[column.Title for column in sheet.Columns]])
            requests = []
            setHeaderStyle(requests, sheet)
            setColumnWidth(requests, sheet)
            DataValidation(requests, sheet)
            if self.TypeOfDoc == "Timings":
                setCheckBoxOnSecondLine(requests, sheet)
                addGraph(requests, sheet)
                DurationFormat(requests, sheet)
                DateTimeFormat(requests, sheet)
            else:
                TextFormat(requests, sheet)
            self.SendRequest(body={"requests": requests})

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

    def GetValue(self):
        pass

    def UpdateValue(self, range, values):
        while True:
            try:
                res = self.service.spreadsheets().values().update(spreadsheetId=self.SSID,
                                                                  range=range,
                                                                  valueInputOption="USER_ENTERED",
                                                                  body={"values": values}).execute()
                break
            except Exception as e:
                for i in range(100, 0, -1):
                    os.system("clear")
                    print(e)
                    print("Превышен лимит ожидания")
                    print(f"Повторная отправка через {i} сек")
                    sleep(1)
        return res