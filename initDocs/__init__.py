import initDocs.init_BugReport as BR
import initDocs.init_Timings as TIMINGS
import initDocs.init_FeedbackForm as FBF


def init(typeOfDoc):
    cases = {"BugReport": BR,
             "Timings": TIMINGS,
             "FeedbackForm": FBF}

    return cases[typeOfDoc].initDoc()