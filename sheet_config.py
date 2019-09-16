SPREADSHEET_TITLE = "API genereated DIAMANT DOC"
NUMBER_OF_TESTES = 3
HEADER_1 =[
            ["Klass", "", "19/20", "HT", "", "Diagram"],
            ["", "", "", "test1", "test2", "test3"],
            ["Elev", "SVA", "Kön", "test1", "test2", "test3"],
        ]

# Template for compiled results
trailing_widths = [40] * 5
template_complied_results = [200]
template_complied_results.extend([75] * (2 + NUMBER_OF_TESTES))
template_complied_results.extend(trailing_widths)


# Tempalte for diagnostics sheets
tests = [3, 15, 17]
test_spacer = [10]
test_result = [50]

# Building template_diagnostics, starting with first three columns.
template_diagnostics = [200]
template_diagnostics.extend([75] * 2)
for x in tests:
    test = []
    test = [25] * x
    test.extend(test_result)
    test.extend(test_spacer)
    template_diagnostics.extend(test)

# Collecting the different templates in template_dict:dictionary
template_dict = {
    "template_1": {
        "sheets": ["Klass 7A", "Klass 7B", "Klass 7C"],
        "header": HEADER_1,
        "columns": template_complied_results
    },
    "template_2": {
        "sheets": ["Klass 7A - Diagnoser", "Klass 7B - Diagnoser", "Klass 7C - Diagnoser"],
        "header": [
            ["Klass", "19/20", "HT"],
            ["", "", ""],
            ["Elev", "SVA", "Kön"],
        ],
        "columns": template_diagnostics
    }
}

def generate_temp_dict():
    template_dict = {
        "template_1": {
            "sheets": ["Klass 7A", "Klass 7B", "Klass 7C"],
            "header": HEADER_1,
            "columns": template_complied_results
        },
        "template_2": {
            "sheets": ["Klass 7A - Diagnoser", "Klass 7B - Diagnoser", "Klass 7C - Diagnoser"],
            "header": [
                ["Klass", "19/20", "HT"],
                ["", "", ""],
                ["Elev", "SVA", "Kön"],
            ],
            "columns": template_diagnostics
        }
    }

    return template_dict

sheet_names = ["Klass 7A", "Klass 7A - Diagnoser", "Klass 7B", "Klass 7B - Diagnoser", "Klass 7C", "Klass 7C - Diagnoser"]
sheet_template = [
    template_dict['template_1']['columns'], 
    template_dict['template_2']['columns'], 
    template_dict['template_1']['columns'], 
    template_dict['template_2']['columns'],
    template_dict['template_1']['columns'], 
    template_dict['template_2']['columns'],
]

tabcolor = {
    "Klass 7A": {
        "red": 0.4,
        "green": 0.3,
        "blue": 1.0
    },
    "Klass 7A - Diagnoser": {
        "red": 0.4,
        "green": 0.3,
        "blue": 1.0
    },
    "Klass 7B": {
        "red": 0.4,
        "green": 0.8,
        "blue": 0.4
    },
    "Klass 7B - Diagnoser": {
        "red": 0.4,
        "green": 0.8,
        "blue": 0.4
    },
    "Klass 7C": {
        "red": 1.0,
        "green": 0.2,
        "blue": 0.2
    },
    "Klass 7C - Diagnoser": {
        "red": 1.0,
        "green": 0.2,
        "blue": 0.2
    }
}



# BUILDING SHEETS
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


# ADDING COLUMNS TO THOSE SHEETS THAT NEEDS IT
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


# SETTING COLUMN WIDTH TO THE DIFFERENT SHEETS
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

