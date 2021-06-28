from BugReport import *
import json
import sys
import argparse


def readDocs(fname):
    with open(fname, "r") as read_file:
        data = json.load(read_file)
    return data


def deleteDoc(Name):
    pass


def addDoc(data, Name, spreadsheet_id):
    data["Docs"].append({"Name": Name, "spreadsheet_id": spreadsheet_id})
    return data


def writeDocs(fname, data):
    with open(fname, "w") as write_file:
        json.dumps(data, write_file)


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
    fname = "Docs.json"
    Docs = readDocs(fname)
    if not True:
        Docs = addDoc(Docs, "", "")
        writeDocs(fname, Docs)

    parser = argparse.ArgumentParser(description='Parser for bug reports in google sheets.')
    parser.add_argument("--a", default=1, type=int, required=False, help="This is the 'a' variable")
    parser.add_argument("--SSID", type=str, help="Spreadsheet ID of sheet")
    parser.add_argument("--dd", default=False, type=bool, required=False, help="Create full document", action=doDoc)
    # parser.add_argument("--education",
    #                     choices=["highschool", "college", "university", "other"],
    #                     required=True, type=str, help="Your name")

    args = parser.parse_args()
    print(args)


if __name__ == '__main__':
    #main()
    #1tA-dQswk0dbggyuhJ_C9NMjIjEIRV1TaNFIbF8IxpPs
    #adpo
    #1oO27abVe4v1UJjO_RF2PxbIp_aKNZ78Nj7wNAsC2YmY
    br = BugReport(spreadsheet_id='1Ys5m0ehZHl4HQpl9ZRF3-TQ95S5--Ocp90DmsC4TGRo', for_Romanovskaya=False)

    br.doDoc()
