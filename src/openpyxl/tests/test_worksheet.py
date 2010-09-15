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

from openpyxl.workbook import Workbook
from openpyxl.worksheet import Worksheet
from openpyxl.cell import Cell

from openpyxl.shared.exc import CellCoordinatesException

class TestWorksheet(BaseTestCase):

    def setUp(self):

        self.wb = Workbook()

    def test_new_worksheet(self):

        ws = Worksheet(parent_workbook = self.wb)

        self.assertEqual(self.wb, ws._parent)

    def test_get_cell(self):

        ws = Worksheet(parent_workbook = self.wb)

        c = ws.cell(coordinate = 'A1')

        self.assertEqual(c.get_coordinate(), 'A1')

    def test_set_wrong_title(self):

        self.assertRaises(Exception, Worksheet, self.wb, 'X' * 50)

    def test_worksheet_dimension(self):

        ws = Worksheet(parent_workbook = self.wb)

        self.assertEqual('A1:A1', ws.calculate_dimension())

        ws.cell('B12').value = 'AAA'

        self.assertEqual('A1:B12', ws.calculate_dimension())

    def test_worksheet_range(self):

        ws = Worksheet(parent_workbook = self.wb)

        rng = ws.range('A1:C4')

        self.assertTrue(isinstance(rng, tuple))

        self.assertEqual(4, len(rng))

        self.assertEqual(3, len(rng[0]))

    def test_worksheet_range_named_range(self):

        ws = Worksheet(parent_workbook = self.wb)

        self.wb.create_named_range(name = 'test_range', worksheet = ws, range = 'C5')

        rng = ws.range("test_range")

        self.assertTrue(isinstance(rng, Cell))

        self.assertEqual(5, rng.row) #pylint: disable-msg=E1103

    def test_cell_offset(self):

        ws = Worksheet(parent_workbook = self.wb)

        self.assertEqual('C17', ws.cell('B15').offset(row = 2, column = 1).get_coordinate())

    def test_range_offset(self):

        ws = Worksheet(parent_workbook = self.wb)

        rng = ws.range('A1:C4', row = 1, column = 3)

        self.assertTrue(isinstance(rng, tuple))

        self.assertEqual(4, len(rng))

        self.assertEqual(3, len(rng[0]))

        self.assertEqual('D2', rng[0][0].get_coordinate())

    def test_cell_alternate_coordinates(self):

        ws = Worksheet(parent_workbook = self.wb)

        c = ws.cell(row = 8, column = 4)

        self.assertEqual('D8', c.get_coordinate())

    def test_cell_range_name(self):

        ws = Worksheet(parent_workbook = self.wb)

        self.wb.create_named_range(name = 'test_range_single', worksheet = ws, range = 'B12')

        self.assertRaises(CellCoordinatesException, ws.cell, 'test_range_single')

        c_range_name = ws.range('test_range_single')
        c_range_coord = ws.range('B12')
        c_cell = ws.cell('B12')

        self.assertEqual(c_range_coord, c_range_name)
        self.assertEqual(c_range_coord, c_cell)

    def test_garbage_collect(self):

        ws = Worksheet(parent_workbook = self.wb)

        ws.cell('A1').value = ''
        ws.cell('B2').value = '0'
        ws.cell('C4').value = 0

        ws.garbage_collect()

        self.assertEqual(ws.get_cell_collection(), [ws.cell('B2'), ws.cell('C4')])


    def test_hyperlink_relationships(self):
        ws = Worksheet(parent_workbook = self.wb)
        self.assertEqual(len(ws.relationships), 0)

        ws.cell('A1').hyperlink = "http://test.com"
        self.assertEqual(len(ws.relationships), 1)
        self.assertEqual("rId1", ws.cell('A1').hyperlink_rel_id)
        self.assertEqual("rId1", ws.relationships[0].id)
        self.assertEqual("http://test.com", ws.relationships[0].target)
        self.assertEqual("External", ws.relationships[0].target_mode)

        ws.cell('A2').hyperlink = "http://test2.com"
        self.assertEqual(len(ws.relationships), 2)
        self.assertEqual("rId2", ws.cell('A2').hyperlink_rel_id)
        self.assertEqual("rId2", ws.relationships[1].id)
        self.assertEqual("http://test2.com", ws.relationships[1].target)
        self.assertEqual("External", ws.relationships[1].target_mode)
