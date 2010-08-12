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
from openpyxl.namedrange import NamedRange

class TestWorkbook(BaseTestCase):

    def test_new_workbook(self):

        wb = Workbook()

    def test_get_active_sheet(self):

        wb = Workbook()

        ash = wb.get_active_sheet()

        self.assertEqual(ash, wb.worksheets[0])

    def test_create_sheet(self):

        wb = Workbook()

        nsh = wb.create_sheet(index = 0)

        self.assertEqual(nsh, wb.worksheets[0])

    def test_remove_sheet(self):

        wb = Workbook()

        nsh = wb.create_sheet(index = 0)

        wb.remove_sheet(nsh)

        self.assertFalse(nsh in wb.worksheets)

    def test_get_sheet_by_name(self):

        wb = Workbook()

        nsh = wb.create_sheet()

        title = 'my sheet'

        nsh.title = title

        fsh = wb.get_sheet_by_name(name = title)

        self.assertEqual(nsh, fsh)

    def test_get_index(self):

        wb = Workbook()

        nsh = wb.create_sheet(index = 0)

        sidx = wb.get_index(nsh)

        self.assertEqual(sidx, 0)

    def test_get_sheet_names(self):

        wb = Workbook()

        names = ['Sheet', 'Sheet1', 'Sheet2', 'Sheet3', 'Sheet4', 'Sheet5']
        for i in xrange(5):
            wb.create_sheet(index = 0)

        actual_names = wb.get_sheet_names()
        self.assertEqual(sorted(actual_names), sorted(names))


    def test_get_named_ranges(self):

        wb = Workbook()

        self.assertEqual(wb.get_named_ranges(), wb._named_ranges)

    def test_add_named_range(self):

        wb = Workbook()

        nsh = wb.create_sheet()

        nr = NamedRange(name = 'test_nr', worksheet = nsh, range = 'A1')

        wb.add_named_range(named_range = nr)

        named_ranges_list = wb.get_named_ranges()

        self.assertTrue(nr in named_ranges_list)

    def test_get_named_range(self):

        wb = Workbook()

        nsh = wb.create_sheet()

        nr = NamedRange(name = 'test_nr', worksheet = nsh, range = 'A1')

        wb.add_named_range(named_range = nr)

        fnr = wb.get_named_range(name = 'test_nr')

        self.assertEqual(nr, fnr)

    def test_remove_named_range(self):

        wb = Workbook()

        nsh = wb.create_sheet()

        nr = NamedRange(name = 'test_nr', worksheet = nsh, range = 'A1')

        wb.add_named_range(named_range = nr)

        wb.remove_named_range(named_range = nr)

        named_ranges_list = wb.get_named_ranges()

        self.assertFalse(nr in named_ranges_list)

