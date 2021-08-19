from BugReport import *
import json
import sys
import argparse


class doDoc(argparse.Action):

    def __init__(self, option_strings, dest, nargs=None, **kwargs):
        print('work')
        if nargs is not None:
            raise ValueError("nargs not allowed")
        super().__init__(option_strings, dest, **kwargs)

    def __call__(self, parser, namespace, values, option_string=None):
        print('%r %r %r' % (namespace, values, option_string))
        setattr(namespace, self.dest, values)
        print("this work")

def main():
    parser = argparse.ArgumentParser(description='Script for creating templates in google sheets')
    parser.add_argument("--SSID", type=str, help="Spreadsheet ID of sheet")
    parser.add_argument("--dd", choices=["br", "fbf", "times"], required=False, type=str, help="--dd Do doc with name 'br'-BugReport, 'fbf'-Feedback From)")
    # parser.add_argument("--br", action='store_true')
    # parser.add_argument("--dd", default=False, type=bool, required=False, help="Create full document", action=doDoc)
    # parser.add_argument("--a", default=1, type=int, required=False, help="This is the 'a' variable")


    args = parser.parse_args()
    print(args)
    return args


if __name__ == '__main__':
    args = main()
    print(args.SSID)
    print(args.dd)

    br = BugReport(spreadsheet_id=args.SSID)
    if args.dd:
        br.doDoc(type=args.dd, for_Romanovskaya=False)
