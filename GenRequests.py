from Color import Color

ind2str = lambda ind: chr(65 + ind)
FontHeader = {"foregroundColor": Color.Black, "fontSize": 12, "bold": True}
HeaderStyle = {"backgroundColor": Color.Cyan, "horizontalAlignment": "CENTER", "textFormat": FontHeader}


def setHeaderStyle(requests, sheet):
    # backgroundColor, textFormat, horizontalAlignment
    requests.append({
        "repeatCell": {
            "range": {
                "sheetId": sheet.ID,
                "startRowIndex": 0,
                "endRowIndex": sheet.FR
            },
            "cell": {
                "userEnteredFormat": HeaderStyle
            },
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
        }
    })
    requests.append({
        "updateSheetProperties": {
            "properties": {
                "sheetId": sheet.ID,
                "gridProperties": {"frozenRowCount": sheet.FR}
            },
            "fields": "gridProperties.frozenRowCount"
        }
    })


def setColumnWidth(requests, sheet):
    for column in sheet.Columns:
        requests.append({
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
        })


def setCheckBoxOnSecondLine(requests, sheet):
    requests.append({
        "setDataValidation": {
            "range": {
                "sheetId": sheet.ID,
                "startRowIndex": 1,
                "endRowIndex": 2
            },
            "rule": {
                "condition": {
                    "type": "BOOLEAN"
                },
                "showCustomUi": True,
                "strict": True
            }
        }
    })


def addGraph(requests, sheet):
    series = []
    for i in range(1, len(sheet.Columns)):
        series.append({
            "series": {
                "sourceRange": {
                    "sources": [
                        {
                            "sheetId": sheet.ID,
                            "startRowIndex": 0,
                            "endRowIndex": sheet.Rows,
                            "startColumnIndex": i,
                            "endColumnIndex": i + 1
                        }
                    ]
                }
            },
            "targetAxis": "LEFT_AXIS"
        })
    requests.append({
        "addChart": {
            "chart": {
                "spec": {
                    "title": f"Graphics of {sheet.Title}'th day",
                    "basicChart": {
                        "chartType": "LINE",
                        "legendPosition": "BOTTOM_LEGEND",
                        "axis": [
                            {
                                "position": "BOTTOM_AXIS",
                                "title": "DateTime"
                            },
                            {
                                "position": "LEFT_AXIS",
                                "title": "Time"
                            }
                        ],
                        "domains": [
                            {
                                "domain": {
                                    "sourceRange": {
                                        "sources": [
                                            {
                                                "sheetId": sheet.ID,
                                                "startRowIndex": 0,
                                                "endRowIndex": sheet.Rows,
                                                "startColumnIndex": 0,
                                                "endColumnIndex": 1
                                            }
                                        ]
                                    }
                                }
                            }
                        ],
                        "series": series,
                        "headerCount": 1
                    }
                },
                "position": {
                    "overlayPosition": {
                        "anchorCell": {
                            "sheetId": sheet.ID,
                            "rowIndex": 0,
                            "columnIndex": 0
                        },
                        "offsetXPixels": 0,
                        "offsetYPixels": 0,
                        "widthPixels": 1200,
                        "heightPixels": 640
                    }
                }
            }
        }
    })


def TextFormat(requests, sheet):
    for i in range(sheet.FR, sheet.Rows):
        color = sheet.LineColor[i % 2]
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet.ID,
                    "startRowIndex": i,
                    "endRowIndex": i+1
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
        })


def DurationFormat(requests, sheet):
    requests.append({
        "repeatCell": {
            "range": {
                "sheetId": sheet.ID,
                "startRowIndex": 2,
                "endRowIndex": sheet.Rows,
                "startColumnIndex": 1,
                "endColumnIndex": len(sheet.Columns)
            },
            "cell": {
                "userEnteredFormat": {
                    "numberFormat": {"type": "TIME", "pattern": "[hh]:mm:ss"}
                }
            },
            "fields": "userEnteredFormat(numberFormat)"
        }})


def DateTimeFormat(requests, sheet):
    requests.append({
        "repeatCell": {
            "range": {
                "sheetId": sheet.ID,
                "startColumnIndex": 0,
                "endColumnIndex": 1
            },
            "cell": {
                "userEnteredFormat": {
                    "numberFormat": {"type": "DATE_TIME"}
                }
            },
            "fields": "userEnteredFormat(numberFormat)"
        }})


def DataValidation(requsets, sheet):
    for column in sheet.Columns:
        if column.DataValidation is not None:
            if column.DataValidation == "Checkbox":
                requsets.append({
                    "setDataValidation": {
                        "range": {
                            "sheetId": sheet.ID,
                            "startRowIndex": 1,
                            "endRowIndex": sheet.Rows,
                            "startColumnIndex": column.ID,
                            "endColumnIndex": column.ID + 1
                        },
                        "rule": {
                            "condition": {
                                "type": "BOOLEAN"
                            },
                            "showCustomUi": True,
                            "strict": True
                        }
                    }
                })
            else:
                requsets.append({
                    "setDataValidation": {
                        "range": {
                            "sheetId": sheet.ID,
                            "startRowIndex": 1,
                            "endRowIndex": sheet.Rows,
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
