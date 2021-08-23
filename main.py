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
        print("start doing")
        br.doDoc(type=args.dd, for_Romanovskaya=False)

def test():
    from  datetime import datetime
    br = BugReport(spreadsheet_id='1cwMf4h8dANH6_sTIIfVR3cV04jWOt0lYZw2Nx9FgxoE')
    mm = lambda t, l: f"=if({l}2,1,0)*\"{t}\""
    data = [[str(datetime.now()), '00:05:01', '00:05:01', '00:05:01', '00:05:01', '00:05:01', '00:05:01', '00:05:01', '00:05:01', '00:05:01', '00:05:01', '00:05:01', '00:05:01'],
            [str(datetime.now()), '00:06:01', '00:05:01', '00:05:01', '00:05:01', '00:02:01', '00:05:01', '00:05:01', '00:05:01', '00:05:01', '00:05:01']]

    br.initColumns("times")
    br.initSheets("times")
    br.getSheets()
    br.addData(br.__Sheets__[0], data)

if __name__ == '__main__':
    main()
    # test()
