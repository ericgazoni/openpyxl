# coding=UTF-8
'''
Copyright (c) 2010 openpyxl

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.

@license: http://www.opensource.org/licenses/mit-license.php
@author: Eric Gazoni
'''

from openpyxl.tests.helper import BaseTestCase

from openpyxl.worksheet import Worksheet
from openpyxl.workbook import Workbook
from openpyxl.cell import column_index_from_string, coordinate_from_string, get_column_letter, Cell, absolute_coordinate
from datetime import time, datetime

class TestCell(BaseTestCase):

    def test_coordinates(self):

        column, row = coordinate_from_string(coord_string = "ZF46")

        self.assertEqual("ZF", column)
        self.assertEqual(46, row)

    def test_invalid_coordinate(self):

        self.assertRaises(Exception, coordinate_from_string, "AAA")

    def test_absolute(self):

        self.assertEqual('$ZF$51', absolute_coordinate(coord_string = 'ZF51'))

    def test_column_index(self):

        self.assertEqual(10, column_index_from_string(column = 'J'))

        self.assertEqual(270, column_index_from_string(column = 'JJ'))

        self.assertEqual(7030, column_index_from_string(column = 'JJJ'))

        self.assertRaises(Exception, column_index_from_string, 'JJJJ')

        self.assertRaises(Exception, column_index_from_string, '')

    def test_column_letter(self):

        self.assertEqual('ZZZ', get_column_letter(col_idx = 18278))

        self.assertEqual('AA', get_column_letter(col_idx = 27))

        self.assertEqual('Z', get_column_letter(col_idx = 26))

    def test_value(self):

        c = Cell(worksheet = None, column = 'A', row = 1)

        self.assertEqual(c.TYPE_NULL, c.data_type)

        c.value = 42
        self.assertEqual(c.TYPE_NUMERIC, c.data_type)

        c.value = 'hello'
        self.assertEqual(c.TYPE_STRING, c.data_type)

        c.value = '=42'
        self.assertEqual(c.TYPE_FORMULA, c.data_type)

        c.value = '4.2'
        self.assertEqual(c.TYPE_NUMERIC, c.data_type)

        c.value = '-42.00'
        self.assertEqual(c.TYPE_NUMERIC, c.data_type)

        c.value = '0'
        self.assertEqual(c.TYPE_NUMERIC, c.data_type)

        c.value = 0
        self.assertEqual(c.TYPE_NUMERIC, c.data_type)

        c.value = 0.0001
        self.assertEqual(c.TYPE_NUMERIC, c.data_type)

        c.value = '0.9999'
        self.assertEqual(c.TYPE_NUMERIC, c.data_type)

    def test_time_value(self):

        wb = Workbook()
        ws = Worksheet(parent_workbook = wb)

        c = Cell(worksheet = ws, column = 'A', row = 1)

        c.value = '03:40:16'
        self.assertEqual(c.TYPE_NUMERIC, c.data_type)

        self.assertEqual(time(3, 40, 16), c.value)

    def test_date_format_applied_on_non_dates(self):

        wb = Workbook()
        ws = Worksheet(parent_workbook = wb)

        c = Cell(worksheet = ws, column = 'A', row = 1)

        c.value = datetime.now()

        c.value = 'testme'

        self.assertEqual('testme', c.value)
