from diamant_test_config import *
import pprint
pp = pprint.PrettyPrinter(indent=2)

SPREADSHEET_TITLE = "not set"
YEAR = "not set"
TERMIN = "not set"
indata = {}

datafile = open("data.txt", "r")
indata_index = 0
for line in datafile:
    print(line)
    linelist = line.split('=')
    linelist = list(map(str.strip, linelist))
    print(linelist)
    if linelist[0] == 'TITLE':
        print("SPREADSHEET_TITLE IS SET")
        SPREADSHEET_TITLE = linelist[1]
    elif linelist[0] == 'YEAR':
        YEAR = linelist[1]
    elif linelist[0] == 'TERMIN':
        TERMIN = linelist[1]
    elif linelist[0] == 'INDATA':
        klasslist = linelist[1].split(';')
        klasser = klasslist[0].split(",")
        diagnoser = klasslist[1].split(",")
        indata[indata_index] = [klasser, diagnoser]
        indata_index += 1

print(indata)
# SPREADSHEET_TITLE = "API generarad diamantdiagnos"
# YEAR = "19/20"
# TERMIN = "VT"
# HEADER_TEMPLATE = [
#             ["Klass", YEAR, TERMIN],
#             ["", "", ""],
#             ["Elev", "SVA", "Kön"],
#         ]
# indata = {
#     0: [
#         ["4A", "4B", "4C", "4D"],
#         ["ag4", "as1", "as2"]
#     ],
#     1: [
#         ["5A", "5B", "5C", "5D"],
#         ["ag6", "ag8", "ag9", "mti1", "mti2"]
#     ],
#     2: [
#         ["6A", "6B", "6C", "6D"],
#         ["ag6", "as9", "rd1", "rd2"]
#     ]
# }
# indata = {
#     0: [
#         ["7A"],
#         ["mti1"]
#     ]
# }



def get_sheet_names(indata):
    sheet_names = []
    for key, value in indata.items():
        for klass in value[0]:
            sheet_names.append("Klass " + klass)
            sheet_names.append("Klass " + klass + " - Diagnoser")
    print(sheet_names)
    return sheet_names

def generate_header(indata, diamant_tests):
    headers = {}
    ####
    # header = {
    #    0 : {
    #        0 : [[]],
    #        1 : [[]]
    #    },
    #    1 : {
    #        0 : [[]],
    #        1 : [[]] 
    #    }
    # }
    ####

    for key, value in indata.items():
        ht_template = [
            ["Klass", YEAR, TERMIN],
            ["", "", ""],
            ["Elev", "SVA", "Kön"],
        ]
        h0 = {}
        headers[key] = h0
        headers[key][1] = [
            ["Klass", YEAR, TERMIN],
            ["", "", ""],
            ["Elev", "SVA", "Kön"],
        ]
        for i, test in enumerate(value[1]):
            print("     test: ", test)
            test_name = diamant_tests[test]['name']
            test_number = str(i + 1)
            ht_template[1].extend([test_name])
            ht_template[2].extend(["Test " + test_number])
        headers[key][0] = ht_template

    return headers

sheet_names = get_sheet_names(indata)

headers = generate_header(indata, diamant_tests)

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

def generate_template_dict(indata, headers):
    template_dict = {}
   
    for key, value in indata.items():
        template = {}
        template_number = (key+1)*2-1           # Calculate correct template_dict index (0,1,2,3,...)
        template_name = "template_" + str(template_number)
        template_dict[template_name] = template # dictionariet får ett tomt dictionary
        
        # FIRST TEMPLATE
        sheet_list = []
        for klass in indata[key][0]:        # Append "Klass " + XX to sheet name
            k = "Klass " + klass
            sheet_list.append(k)

        template_dict[template_name]['sheets'] = sheet_list
        template_dict[template_name]['header'] = headers[key][0]
        template_dict[template_name]['tests'] = indata[key][1]
        template_dict[template_name]['columns'] = get_columns_list(False, indata[key][1])
        template_dict[template_name]['generateClass'] = True
        template_dict[template_name]['testTasks'] = False
        
        # SECOND TEMPLATE
        template = {}
        template_number = (key+1)*2             # Calculate correct template_dict index (0,1,2,3,...)
        template_name = "template_" + str(template_number)
        template_dict[template_name] = template # dictionariet får ett tomt dictionary

        sheet_list = []
        for klass in indata[key][0]:        # Append "Klass " + XX + " - Diagnoser" to sheet name
            k = "Klass " + klass + " - Diagnoser"
            sheet_list.append(k)

        template_dict[template_name]['sheets'] = sheet_list
        template_dict[template_name]['header'] = headers[key][1]
        template_dict[template_name]['tests'] = indata[key][1]
        template_dict[template_name]['columns'] = get_columns_list(True, indata[key][1])
        template_dict[template_name]['generateClass'] = True
        template_dict[template_name]['testTasks'] = True

    return template_dict

####################################
#   TEMPLATE_DICT
####################################
# Collecting the different templates in template_dict:dictionary
# template_dict = {
#     "template_1": {
#         "sheets": ["Klass 7A", "Klass 7B", "Klass 7C"],
#         "header": HEADER_1,
#         "tests": TESTS_1,
#         "columns": get_columns_list(False, TESTS_1),
#         "generateClass": True,
#         "testTasks": False
#     },
#     "template_2": {
#         "sheets": ["Klass 7A - Diagnoser", "Klass 7B - Diagnoser", "Klass 7C - Diagnoser"],
#         "header": HEADER_2,
#         "tests": TESTS_1,
#         "columns": get_columns_list(True, TESTS_1),
#         "generateClass": True,
#         "testTasks": True
#     }
####################################
template_dict = generate_template_dict(indata, headers)

sheet_template = {}
for template in template_dict.keys():
    for sheet in template_dict[template]["sheets"]:
        sheet_template[sheet] = template



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
    },
    "Klass 9A": {
        "red": 0.4,
        "green": 0.3,
        "blue": 1.0
    },
    "Klass 9A - Diagnoser": {
        "red": 0.4,
        "green": 0.3,
        "blue": 1.0
    },
    "Klass 9B": {
        "red": 0.4,
        "green": 0.8,
        "blue": 0.4
    },
    "Klass 9B - Diagnoser": {
        "red": 0.4,
        "green": 0.8,
        "blue": 0.4
    },
    "Klass 9C": {
        "red": 1.0,
        "green": 0.2,
        "blue": 0.2
    },
    "Klass 9C - Diagnoser": {
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
# sheet_objects[1] = {
#     "updateSheetProperties": {
#         "properties": {
#             "sheetId": 0,
#             "tabColor": tabcolor[sheet_names[0]]
#         },
#         "fields": "tabColor"
#     }
# }

for i, sheet in enumerate(sheet_names[1:]):
    # Add new sheet tabs
    sheet_objects[i+2] = {
        "addSheet": {
            "properties": {
                "title": sheet
                # "tabColor": tabcolor[sheet_names[i+1]]
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

