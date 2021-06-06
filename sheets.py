"""
This module will print the names of all the sheets in the Excel File
"""
import active_excel
WORKBOOK = active_excel.WORKBOOK
def print_sheets():
    """ This will return the list of all the sheet names """
    return WORKBOOK.get_sheet_names()
