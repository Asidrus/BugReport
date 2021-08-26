import initDocs.init_BugReport as BR

def __init__(typeOfDoc):
    # cases = {"BugReport": init_BugReport.initSheets,
    #          "Timings": init_Timings.initDoc}
    # cases = {"BugReport": BR.initSheets}

    return BR.initSheets()