# file openpyxl/tests/test_iter.py

# Copyright (c) 2010-2011 openpyxl
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#
# @license: http://www.opensource.org/licenses/mit-license.php
# @author: see AUTHORS file

from nose.tools import eq_, raises, assert_raises
import os.path as osp
from openpyxl.tests.helper import DATADIR
from openpyxl.reader.iter_worksheet import get_range_boundaries
from openpyxl.reader.excel import load_workbook
from openpyxl.shared.compat import xrange
import datetime

class TestWorksheet(object):

    workbook_name = osp.join(DATADIR, 'genuine', 'empty.xlsx')

    def _open_wb(self):
        return load_workbook(filename = self.workbook_name, use_iterators = True)

class TestDims(TestWorksheet):
    expected = [ 'A1:G5', 'D1:K30', 'D2:D2', 'A1:C1' ]
    def test_get_dimensions(self):
        wb = self._open_wb()
        for i, sheetn in enumerate(wb.get_sheet_names()):
            ws = wb.get_sheet_by_name(name = sheetn)

            eq_(ws._dimensions, self.expected[i])

    def test_get_highest_column_iter(self):
        wb = self._open_wb()
        ws = wb.worksheets[0]
        eq_(ws.get_highest_column(), 7)

class TestText(TestWorksheet):
    sheet_name = 'Sheet1 - Text'

    expected = [['This is cell A1 in Sheet 1', None, None, None, None, None, None],
                [None, None, None, None, None, None, None],
                [None, None, None, None, None, None, None],
                [None, None, None, None, None, None, None],
                [None, None, None, None, None, None, 'This is cell G5'], ]

    def test_read_fast_integrated(self):

        wb = self._open_wb()
        ws = wb.get_sheet_by_name(name = self.sheet_name)

        for row, expected_row in zip(ws.iter_rows(), self.expected):

            row_values = [x.internal_value for x in row]

            eq_(row_values, expected_row)


    def test_get_boundaries_range(self):

        eq_(get_range_boundaries('C1:C4'), (3, 1, 3, 4))

    def test_get_boundaries_one(self):


        eq_(get_range_boundaries('C1'), (3, 1, 4, 1))

    def test_read_single_cell_range(self):

        wb = self._open_wb()
        ws = wb.get_sheet_by_name(name = self.sheet_name)

        eq_('This is cell A1 in Sheet 1', list(ws.iter_rows('A1'))[0][0].internal_value)

class TestIntegers(TestWorksheet):

    sheet_name = 'Sheet2 - Numbers'

    expected = [[x + 1] for x in xrange(30)]

    query_range = 'D1:E30'

    def test_read_fast_integrated(self):

        wb = self._open_wb()
        ws = wb.get_sheet_by_name(name = self.sheet_name)

        for row, expected_row in zip(ws.iter_rows(self.query_range), self.expected):

            row_values = [x.internal_value for x in row]

            eq_(row_values, expected_row)


class TestFloats(TestWorksheet):

    sheet_name = 'Sheet2 - Numbers'
    query_range = 'K1:L30'
    expected = [[(x + 1) / 100.0] for x in xrange(30)]

    def test_read_fast_integrated(self):

        wb = self._open_wb()
        ws = wb.get_sheet_by_name(name = self.sheet_name)

        for row, expected_row in zip(ws.iter_rows(self.query_range), self.expected):

            row_values = [x.internal_value for x in row]

            eq_(row_values, expected_row)


class TestDates(TestWorksheet):

    sheet_name = 'Sheet4 - Dates'

    def test_read_single_cell_date(self):

        wb = self._open_wb()
        ws = wb.get_sheet_by_name(name = self.sheet_name)

        eq_(datetime.datetime(1973, 5, 20), list(ws.iter_rows('A1'))[0][0].internal_value)
        eq_(datetime.datetime(1973, 5, 20, 9, 15, 2), list(ws.iter_rows('C1'))[0][0].internal_value)


class TestReadFormulae(object):

    xml_src = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><dimension ref="A1:B6"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A6" sqref="A6"/></sheetView></sheetViews><sheetFormatPr baseColWidth="10" defaultColWidth="9.140625" defaultRowHeight="15" x14ac:dyDescent="0.25"/><cols><col min="1" max="1" width="15.7109375" customWidth="1"/><col min="2" max="2" width="15.28515625" customWidth="1"/></cols><sheetData><row r="1" spans="1:2" x14ac:dyDescent="0.25">
<c r="A1" t="s"><v>0</v></c>
<c r="B1" t="str"><f>CONCATENATE(A1,A2)</f><v>Hello, world!</v></c></row><row r="2" spans="1:2" x14ac:dyDescent="0.25">
<c r="A2" t="s"><v>1</v></c></row><row r="4" spans="1:2" x14ac:dyDescent="0.25">
<c r="A4"><v>1</v></c></row><row r="5" spans="1:2" x14ac:dyDescent="0.25">
<c r="A5"><v>2</v></c></row><row r="6" spans="1:2" x14ac:dyDescent="0.25">
<c r="A6"><f>SUM(A4:A5)</f><v>3</v></c></row></sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>"""

    @classmethod
    def setup(self):
        from openpyxl.reader.iter_worksheet import iterparse
        from openpyxl.reader.worksheet import StringIO
        src = StringIO(self.xml_src)
        self.parser = iterparse(src)
        self.boundaries = (0, 0, 6, 2)

    def test_get_cells(self):
        from openpyxl.reader.iter_worksheet import get_cells
        cells = {}
        for cell in get_cells(self.parser, *self.boundaries):
            cells[cell.coordinate] = cell
        b1 = cells['B1']
        eq_(b1.data_type, 'f')
        eq_(b1.internal_value, '=CONCATENATE(A1,A2)')
        a6 = cells['A6']
        eq_(a6.data_type, 'f')
        eq_(a6.internal_value, '=SUM(A4:A5)')
