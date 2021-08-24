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