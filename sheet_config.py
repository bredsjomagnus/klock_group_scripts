SPREADSHEET_TITLE = "TEST diamant sheet"
NUMBER_OF_TESTES = 3
trailing_widths = [40] * 5
column_widths = [75] * ( 3 + NUMBER_OF_TESTES)
column_widths.extend(trailing_widths)

print("column_widths:", column_widths)

sheets = ["Klass 7A", "Klass 7B", "Klass 7C", "Klass 8A", "Klass 8B", "Klass 8C"]
tabcolor = {
    "Klass 7A": {
        "red": 0.2,
        "green": 0.3,
        "blue": 1.0
    },
    "Klass 7B": {
        "red": 0.6,
        "green": 0.3,
        "blue": 1.0
    }
}
columns = {}
for i, width in enumerate(column_widths):
    columns[i] = {
        "updateDimensionProperties": {
            "range": {
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



spreadsheet_settings = {
    # Set title on spreadsheet
    "spreadsheet_title": {
        'updateSpreadsheetProperties': {
            'properties': {
                'title': SPREADSHEET_TITLE
            },
            'fields': 'title'
        }
    },

    # Update sheet name of first sheet.
    "first_sheet_name": {
        "updateSheetProperties": {
            "properties": {
                "sheetId": 0,
                "title": sheets[0],
            },
            "fields": "title"
        }
    },

    # Change tab color on first sheet
    "first_sheet_tabcolor": {
        "updateSheetProperties": {
            "properties": {
                "sheetId": 0,
                "tabColor": tabcolor[sheets[0]]
            },
            "fields": "tabColor"
        }
    },

    # Adjust column width
    "column_width": {
        "updateDimensionProperties": {
            "range": {
                "dimension": "COLUMNS",
                "startIndex": 0,
                "endIndex": 1
            },
            "properties": {
                "pixelSize": 50
            },
            "fields": "pixelSize"
        }
    },
    # Add new sheet tab
    "new_sheet": {
        "addSheet": {
            "properties": {
                "title": sheets[1],
                "tabColor": tabcolor[sheets[1]]
            }
        }
    }
}

