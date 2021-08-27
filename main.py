from GoogleSheets import *
import argparse
import os

path = os.path.abspath(os.getcwd())


def main():
    parser = argparse.ArgumentParser(description='Script for creating templates in google sheets')
    parser.add_argument("--SSID", type=str, help="Spreadsheet ID of sheet")
    parser.add_argument("--dd", choices=["BugReport", "FeedbackForm", "Timings"], required=False, type=str,
                        help="--dd Do doc")
    args = parser.parse_args()
    GSh = GoogleSheets(SSID=args.SSID,
                      typeOfDoc=args.dd,
                      CREDENTIALS_FILE=path + "/stable-ring-316114-8acf36454762.json")
    if args.dd:
        GSh.doDoc()


if __name__ == '__main__':
    main()
