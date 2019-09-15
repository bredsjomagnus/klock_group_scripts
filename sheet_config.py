SPREADSHEET_TITLE = "TEST diamant sheet"
NUMBER_OF_TESTES = 3


# Template for compiled results
trailing_widths = [40] * 5
template_complied_results = [75] * (3 + NUMBER_OF_TESTES)
template_complied_results.extend(trailing_widths)


# Tempalte for diagnostics sheets
tests = [3, 15, 17]
test_spacer = [10]
test_result = [50]

template_diagnostics = [75] * 3
for x in tests:
    test = []
    test = [25] * x
    test.extend(test_result)
    test.extend(test_spacer)
    template_diagnostics.extend(test)

template_dict = {
    "template_1": template_complied_results,
    "template_2": template_diagnostics
}


# print("column_widths:", column_widths)

# sheet_names = ["Klass 7A", "Klass 7B", "Klass 7C", "Klass 8A", "Klass 8B", "Klass 8C"]
sheet_names = ["Klass 7A", "Klass 7A - Diagnoser"]
sheet_template = [template_dict['template_1'], template_dict['template_2']]

tabcolor = {
    "Klass 7A": {
        "red": 0.2,
        "green": 0.3,
        "blue": 1.0
    },
    "Klass 7A - Diagnoser": {
        "red": 0.6,
        "green": 0.3,
        "blue": 1.0
    }
}


sheet_objects = {}
sheet_objects[0] = { 
    "updateSheetProperties": {
            "properties": {
                "sheetId": 0,
                "title": sheet_names[0],
            },
            "fields": "title"
        }
}
sheet_objects[1] = {
    "updateSheetProperties": {
        "properties": {
            "sheetId": 0,
            "tabColor": tabcolor[sheet_names[0]]
        },
        "fields": "tabColor"
    }
}

for i, sheet in enumerate(sheet_names[1:]):
    # Add new sheet tabs
    sheet_objects[i+2] = {
        "addSheet": {
            "properties": {
                "title": sheet,
                "tabColor": tabcolor[sheet_names[i+1]]
            }
        }
    }

def generate_add_column_object(sheet_template, sheetIds):
    add_column_objects = {}
    for i, template in enumerate(sheet_template):
        add = len(template) - 24 if len(template) > 25 else 0
        if add > 0:
            add_column_objects[i] = {
                "insertDimension": {
                    "range": {
                    "sheetId": sheetIds[i],
                    "dimension": "COLUMNS",
                    "startIndex": 1,
                    "endIndex": add
                    },
                    "inheritFromBefore": True
                }
            }
    return add_column_objects



def generate_columns_update_object(template_complied_results, sheetIds):
    columns = {}
    list_index = 0
    for j, sheetId in enumerate(sheetIds):
        template = sheet_template[j]
        for i, width in enumerate(template):
            columns[list_index] = {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheetId,
                        "dimension": "COLUMNS",
                        "startIndex": i,
                        "endIndex": i+1
                    },
                    "properties": {
                        "pixelSize": width
                    },
                    "fields": "pixelSize"
                }
            }
            list_index += 1

    return columns



# spreadsheet_settings = {
#     # Set title on spreadsheet
#     "spreadsheet_title": {
#         'updateSpreadsheetProperties': {
#             'properties': {
#                 'title': SPREADSHEET_TITLE
#             },
#             'fields': 'title'
#         }
#     },

#     # Update sheet name of first sheet.
#     # "first_sheet_name": {
#     #     "updateSheetProperties": {
#     #         "properties": {
#     #             "sheetId": 0,
#     #             "title": sheet_names[0],
#     #         },
#     #         "fields": "title"
#     #     }
#     },

#     # Change tab color on first sheet
#     # "first_sheet_tabcolor": {
#     #     "updateSheetProperties": {
#     #         "properties": {
#     #             "sheetId": 0,
#     #             "tabColor": tabcolor[sheet_names[0]]
#     #         },
#     #         "fields": "tabColor"
#     #     }
#     # },

#     # Adjust column width
#     "column_width": {
#         "updateDimensionProperties": {
#             "range": {
#                 "dimension": "COLUMNS",
#                 "startIndex": 0,
#                 "endIndex": 1
#             },
#             "properties": {
#                 "pixelSize": 50
#             },
#             "fields": "pixelSize"
#         }
#     },
#     # Add new sheet tab
#     # "new_sheet": {
#     #     "addSheet": {
#     #         "properties": {
#     #             "title": sheet_names[1],
#     #             "tabColor": tabcolor[sheet_names[1]]
#     #         }
#     #     }
#     # }
# }

