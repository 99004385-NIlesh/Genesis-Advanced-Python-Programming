"""
This module will open the Excel File and makes it active
"""

import openpyxl
PATH = "sample.xlsx"
WORKBOOK = openpyxl.load_workbook(PATH)
SHEET = WORKBOOK.active
"""
End of the module

"""
