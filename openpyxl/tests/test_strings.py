# file openpyxl/tests/test_strings.py

# Python stdlib imports
from __future__ import with_statement
import os.path

# 3rd party imports
from nose.tools import eq_

# package imports
from openpyxl.tests.helper import DATADIR
from openpyxl.workbook import Workbook
from openpyxl.writer.strings import create_string_table
from openpyxl.reader.strings import read_string_table


def test_create_string_table():
    wb = Workbook()
    ws = wb.create_sheet()
    ws.cell('B12').value = 'hello'
    ws.cell('B13').value = 'world'
    ws.cell('D28').value = 'hello'
    table = create_string_table(wb)
    eq_({'hello': 1, 'world': 0}, table)


def test_read_string_table():
    with open(os.path.join(DATADIR, 'reader', 'sharedStrings.xml')) as handle:
        content = handle.read()
    string_table = read_string_table(content)
    eq_({0: 'This is cell A1 in Sheet 1', 1: 'This is cell G5'}, string_table)


def test_formatted_string_table():
    with open(os.path.join(DATADIR, 'reader', 'shared-strings-rich.xml')) \
            as handle:
        content = handle.read()
    string_table = read_string_table(content)
    eq_({0: 'Welcome', 1: 'to the best shop in town',
            2: "     let's play "}, string_table)
