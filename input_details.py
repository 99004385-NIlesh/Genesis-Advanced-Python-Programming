"""
This module will take the input from the user
"""

def input_details():
    """ This will return the Selected PS Number and Sheet Name """
    ps_number_input = int(input("Enter the PS Number: "))
    sheet_input = input("Enter the Sheet Name: ")
    return [ps_number_input, sheet_input]
