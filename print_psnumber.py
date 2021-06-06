"""
This module will print the list of all the PS Numbers in the Excel File
"""
import active_excel
SHEET = active_excel.SHEET

def print_psnumber_list():
    psnumber_list = []
    for i in range(2, SHEET.max_row+1):
        psnumber = SHEET.cell(row=i, column=1)
        if psnumber.value is not None:
            psnumber_list.append(psnumber.value)
    return psnumber_list
