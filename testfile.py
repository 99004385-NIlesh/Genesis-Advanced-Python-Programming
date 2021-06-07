"""
This module will test all the functions using Pytest
"""
from main import *

def test_print_psnumber_list():
    assert print_psnumber_list() == [99004319, 99004320, 99004322, 99004324, 99004356, 99004357, 99004358, 99004359, 99004360, 99004361, 99004362, 99004363, 99004364, 99004365, 99004366]

def test_print_sheets():
    assert print_sheets() == ['Basic Details', 'Cities Travelled', 'Test Marks', 'Hobbies', 'Expertise']

