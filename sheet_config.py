from diamant_test_config import *
import pprint
pp = pprint.PrettyPrinter(indent=2)
SPREADSHEET_TITLE = "API genereated DIAMANT DOC"
sheet_names = [
    "Klass 7A", 
    "Klass 7A - Diagnoser", 
    "Klass 7B", 
    "Klass 7B - Diagnoser", 
    "Klass 7C", 
    "Klass 7C - Diagnoser",
    "Klass 8A", 
    "Klass 8A - Diagnoser", 
    "Klass 8B", 
    "Klass 8B - Diagnoser", 
    "Klass 8C", 
    "Klass 8C - Diagnoser",
    ]
TESTS_1 = ["tae1", "rb4", "rb5"]
TESTS_2 = ["rb4", "rb5"]
HEADER_1 = [
            ["Klass", "", "19/20", "HT", "", "Diagram"],
            ["", "", ""],
            ["Elev", "SVA", "Kön"],
        ]
HEADER_2 = [
            ["Klass", "19/20", "HT"],
            ["", "", ""],
            ["Elev", "SVA", "Kön"],
        ]
HEADER_3 = [
            ["Klass", "", "19/20", "HT", "", "Diagram"],
            ["", "", ""],
            ["Elev", "SVA", "Kön"],
        ]
HEADER_4 = [
            ["Klass", "19/20", "HT"],
            ["", "", ""],
            ["Elev", "SVA", "Kön"],
        ]
for i, test in enumerate(TESTS_1):
    test_name = [diamant_tests[test]['name']]
    test_number = str(i +1)
    HEADER_1[1].extend(test_name)
    HEADER_1[2].extend(["Test " + test_number])

for i, test in enumerate(TESTS_2):
    test_name = [diamant_tests[test]['name']]
    test_number = str(i +1)
    HEADER_3[1].extend(test_name)
    HEADER_3[2].extend(["Test " + test_number])

# print(HEADER_1)

def get_columns_list(test_tasks, test_list):
    columns_list = []
    if not test_tasks:
        # Template for compiled results
        trailing_widths = [40] * 5
        columns_list = [200]
        columns_list.extend([75] * (2 + len(test_list)))
        columns_list.extend(trailing_widths)
    else:
        tests = []
        for test_name in test_list:
            tests.append(len(diamant_tests[test_name]['tasks']))
            
            test_spacer = [diamant_tests['settings']['spacer_column']]
            test_result = [diamant_tests['settings']['result_column']]

        # Building template_diagnostics, starting with first three columns.
        columns_list = [200]
        columns_list.extend([75] * 2)
        for x in tests:
            test = []
            test = [diamant_tests['settings']['task_column_width']] * x
            test.extend(test_result)
            test.extend(test_spacer)
            columns_list.extend(test)
        
    return columns_list

# Collecting the different templates in template_dict:dictionary
template_dict = {
    "template_1": {
        "sheets": ["Klass 7A", "Klass 7B", "Klass 7C"],
        "header": HEADER_1,
        "tests": TESTS_1,
        "columns": get_columns_list(False, TESTS_1),
        "generateClass": True,
        "testTasks": False
    },
    "template_2": {
        "sheets": ["Klass 7A - Diagnoser", "Klass 7B - Diagnoser", "Klass 7C - Diagnoser"],
        "header": HEADER_2,
        "tests": TESTS_1,
        "columns": get_columns_list(True, TESTS_1),
        "generateClass": True,
        "testTasks": True
    },
    "template_3": {
        "sheets": ["Klass 8A", "Klass 8B", "Klass 8C"],
        "header": HEADER_3,
        "tests": TESTS_2,
        "columns": get_columns_list(False, TESTS_2),
        "generateClass": True,
        "testTasks": False
    },
    "template_4": {
        "sheets": ["Klass 8A - Diagnoser", "Klass 8B - Diagnoser", "Klass 8C - Diagnoser"],
        "header": HEADER_4,
        "tests": TESTS_2,
        "columns": get_columns_list(True, TESTS_2),
        "generateClass": True,
        "testTasks": True
    }
}

sheet_template = {}
for template in template_dict.keys():
    # print(template)
    for sheet in template_dict[template]["sheets"]:
        sheet_template[sheet] = template

# pp.pprint(sheet_template)



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
    },
    "Klass 8A": {
        "red": 0.4,
        "green": 0.3,
        "blue": 1.0
    },
    "Klass 8A - Diagnoser": {
        "red": 0.4,
        "green": 0.3,
        "blue": 1.0
    },
    "Klass 8B": {
        "red": 0.4,
        "green": 0.8,
        "blue": 0.4
    },
    "Klass 8B - Diagnoser": {
        "red": 0.4,
        "green": 0.8,
        "blue": 0.4
    },
    "Klass 8C": {
        "red": 1.0,
        "green": 0.2,
        "blue": 0.2
    },
    "Klass 8C - Diagnoser": {
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
def generate_add_column_object(sheet_dict):
    add_column_objects = {}
    i = 0
    for key, value in sheet_dict.items():
        sheet_name = key
        sheet_id = value
        template = sheet_template[sheet_name]
        columns_template = template_dict[template]['columns']
        add = len(columns_template) - 24 if len(columns_template) > 25 else 0
        if add > 0:
            add_column_objects[i] = {
                "insertDimension": {
                    "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": 1,
                    "endIndex": add
                    },
                    "inheritFromBefore": True
                }
            }
        i += 1

    return add_column_objects


# SETTING COLUMN WIDTH TO THE DIFFERENT SHEETS
def generate_columns_update_object(sheet_dict):
    columns = {}
    list_index = 0
    for key, value in sheet_dict.items():
        sheet_name = key
        sheet_id = value
        template = sheet_template[sheet_name]
        columns_template = template_dict[template]['columns']
        # if sheet_name in template_dict['template_1']['sheets']:
        #     columns_template = template_dict['template_1']['columns']
        # else:
        #     columns_template = template_dict['template_2']['columns']
        for i, width in enumerate(columns_template):
            columns[list_index] = {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
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

