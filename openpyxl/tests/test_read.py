# file openpyxl/tests/test_read.py

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

# Python stdlib imports
import os.path
from datetime import datetime, date

# 3rd party imports
from nose.tools import eq_, raises

# package imports
from openpyxl.tests.helper import DATADIR
from openpyxl.worksheet import Worksheet
from openpyxl.workbook import Workbook
from openpyxl.style import NumberFormat, Style
from openpyxl.reader.worksheet import read_worksheet, read_dimension
from openpyxl.reader.excel import load_workbook
from openpyxl.shared.exc import InvalidFileException
from openpyxl.shared.date_time import CALENDAR_WINDOWS_1900, CALENDAR_MAC_1904


def test_read_standalone_worksheet():

    class DummyWb(object):

        encoding = 'utf-8'

        excel_base_date = CALENDAR_WINDOWS_1900

        def get_sheet_by_name(self, value):
            return None

    path = os.path.join(DATADIR, 'reader', 'sheet2.xml')
    ws = None
    handle = open(path)
    try:
        ws = read_worksheet(handle.read(), DummyWb(),
                'Sheet 2', {1: 'hello'}, {1: Style()})
    finally:
        handle.close()
    assert isinstance(ws, Worksheet)
    eq_(ws.cell('G5').value, 'hello')
    eq_(ws.cell('D30').value, 30)
    eq_(ws.cell('K9').value, 0.09)


def test_read_standard_workbook():
    path = os.path.join(DATADIR, 'genuine', 'empty.xlsx')
    wb = load_workbook(path)
    assert isinstance(wb, Workbook)

def test_read_standard_workbook_from_fileobj():
    path = os.path.join(DATADIR, 'genuine', 'empty.xlsx')
    fo = open(path, mode='rb')
    wb = load_workbook(fo)
    assert isinstance(wb, Workbook)

def test_read_worksheet():
    path = os.path.join(DATADIR, 'genuine', 'empty.xlsx')
    wb = load_workbook(path)
    sheet2 = wb.get_sheet_by_name('Sheet2 - Numbers')
    assert isinstance(sheet2, Worksheet)
    eq_('This is cell G5', sheet2.cell('G5').value)
    eq_(18, sheet2.cell('D18').value)


def test_read_nostring_workbook():
    genuine_wb = os.path.join(DATADIR, 'genuine', 'empty-no-string.xlsx')
    wb = load_workbook(genuine_wb)
    assert isinstance(wb, Workbook)

@raises(InvalidFileException)
def test_read_empty_file():

    null_file = os.path.join(DATADIR, 'reader', 'null_file.xlsx')
    wb = load_workbook(null_file)

@raises(InvalidFileException)
def test_read_empty_archive():

    null_file = os.path.join(DATADIR, 'reader', 'null_archive.xlsx')
    wb = load_workbook(null_file)

def test_read_dimension():

    path = os.path.join(DATADIR, 'reader', 'sheet2.xml')

    dimension = None
    handle = open(path)
    try:
        dimension = read_dimension(xml_source=handle.read())
    finally:
        handle.close()

    eq_(('D', 1, 'K', 30), dimension)

def test_calculate_dimension_iter():
    path = os.path.join(DATADIR, 'genuine', 'empty.xlsx')
    wb = load_workbook(filename=path, use_iterators=True)
    sheet2 = wb.get_sheet_by_name('Sheet2 - Numbers')
    dimensions = sheet2.calculate_dimension()
    eq_('%s%s:%s%s' % ('D', 1, 'K', 30), dimensions)

def test_get_highest_row_iter():
    path = os.path.join(DATADIR, 'genuine', 'empty.xlsx')
    wb = load_workbook(filename=path, use_iterators=True)
    sheet2 = wb.get_sheet_by_name('Sheet2 - Numbers')
    max_row = sheet2.get_highest_row()
    eq_(30, max_row)

def test_read_workbook_with_no_properties():
    genuine_wb = os.path.join(DATADIR, 'genuine', \
                'empty_with_no_properties.xlsx')
    wb = load_workbook(filename=genuine_wb)

class TestReadWorkbookWithStyles(object):

    @classmethod
    def setup_class(cls):
        cls.genuine_wb = os.path.join(DATADIR, 'genuine', \
                'empty-with-styles.xlsx')
        wb = load_workbook(cls.genuine_wb)
        cls.ws = wb.get_sheet_by_name('Sheet1')

    def test_read_general_style(self):
        eq_(self.ws.cell('A1').style.number_format.format_code,
                NumberFormat.FORMAT_GENERAL)

    def test_read_date_style(self):
        eq_(self.ws.cell('A2').style.number_format.format_code,
                NumberFormat.FORMAT_DATE_XLSX14)

    def test_read_number_style(self):
        eq_(self.ws.cell('A3').style.number_format.format_code,
                NumberFormat.FORMAT_NUMBER_00)

    def test_read_time_style(self):
        eq_(self.ws.cell('A4').style.number_format.format_code,
                NumberFormat.FORMAT_DATE_TIME3)

    def test_read_percentage_style(self):
        eq_(self.ws.cell('A5').style.number_format.format_code,
                NumberFormat.FORMAT_PERCENTAGE_00)


class TestReadBaseDateFormat(object):

    @classmethod
    def setup_class(cls):
        mac_wb_path = os.path.join(DATADIR, 'reader', 'date_1904.xlsx')
        cls.mac_wb = load_workbook(mac_wb_path)
        cls.mac_ws = cls.mac_wb.get_sheet_by_name('Sheet1')

        win_wb_path = os.path.join(DATADIR, 'reader', 'date_1900.xlsx')
        cls.win_wb = load_workbook(win_wb_path)
        cls.win_ws = cls.win_wb.get_sheet_by_name('Sheet1')

    def test_read_win_base_date(self):
        eq_(self.win_wb.properties.excel_base_date, CALENDAR_WINDOWS_1900)

    def test_read_mac_base_date(self):
        eq_(self.mac_wb.properties.excel_base_date, CALENDAR_MAC_1904)

    def test_read_date_style_mac(self):
        eq_(self.mac_ws.cell('A1').style.number_format.format_code,
                NumberFormat.FORMAT_DATE_XLSX14)

    def test_read_date_style_win(self):
        eq_(self.win_ws.cell('A1').style.number_format.format_code,
                NumberFormat.FORMAT_DATE_XLSX14)

    def test_read_date_value(self):
        datetuple = (2011, 10, 31)
        dt = datetime(datetuple[0], datetuple[1], datetuple[2])
        eq_(self.mac_ws.cell('A1').value, dt)
        eq_(self.win_ws.cell('A1').value, dt)
        eq_(self.mac_ws.cell('A1').value, self.win_ws.cell('A1').value)
