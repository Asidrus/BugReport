from BugReport import *
import json
import sys
import argparse


def main():
    parser = argparse.ArgumentParser(description='Script for creating templates in google sheets')
    parser.add_argument("--SSID", type=str, help="Spreadsheet ID of sheet")
    parser.add_argument("--dd", choices=["br", "fbf", "times"], required=False, type=str, help="--dd Do doc with name 'br'-BugReport, 'fbf'-Feedback From, 'times' - for graphics)")
    args = parser.parse_args()
    br = BugReport(spreadsheet_id=args.SSID)
    if args.dd:
        br.doDoc(type=args.dd, for_Romanovskaya=False)

def test():
    from  datetime import datetime
    br = BugReport(spreadsheet_id='1u88yKDi46j1AjpSxVr2tp1sdt1oKyCzoLkSXZ99cGh4')

    data = [[str(datetime.now()), '00:05:01', '00:05:01', '00:05:01', '00:05:01', '00:05:01', '00:05:01', '00:05:01', '00:05:01', '00:05:01', '00:05:01', '00:05:01', '00:05:01'],
            [str(datetime.now()), '00:06:01', '00:05:01', '00:05:01', '00:05:01', '00:02:01', '00:05:01', '00:05:01', '00:05:01', '00:05:01', '00:05:01']]

    br.initColumns("times")
    br.initSheets("times")
    br.getSheets()
    br.SendRequest({
        "requests": [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": br.__Sheets__[0].ID,
                        "startRowIndex": 2,
                        "endRowIndex": br.__MaxBugs__,
                        "startColumnIndex": 1,
                        "endColumnIndex": len(br.__Columns__)
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {"type": "TIME", "pattern": "[hh]:mm:ss"}
                        }
                    },
                    "fields": "userEnteredFormat(numberFormat)"
                }}]})
    br.SendRequest({
        "requests": [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": br.__Sheets__[0].ID,
                        "startColumnIndex": 0,
                        "endColumnIndex": 1
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {"type": "DATE_TIME"}
                        }
                    },
                    "fields": "userEnteredFormat(numberFormat)"
                }}]})
    br.addData(br.__Sheets__[0], data)

if __name__ == '__main__':
    main()
    # test()
