# file openpyxl/tests/test_read.py

# Copyright (c) 2010-2014 openpyxl
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
import pytest

# compatibility imports
from openpyxl.shared.compat import BytesIO, StringIO, unicode, file, tempfile

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
        _guess_types = True
        data_only = False

        def get_sheet_by_name(self, value):
            return None

        def get_sheet_names(self):
            return []

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

@pytest.mark.parametrize("filename", ["sheet2.xml", "sheet2_no_dimension.xml"])
def test_read_dimension(filename):
    path = os.path.join(DATADIR, 'reader', filename)
    dimension = None
    with open(path) as handle:
        dimension = read_dimension(handle.read())
    assert dimension == ('D', 1, 'AA', 30)

def test_calculate_dimension_iter():
    path = os.path.join(DATADIR, 'genuine', 'empty.xlsx')
    wb = load_workbook(filename=path, use_iterators=True)
    sheet2 = wb.get_sheet_by_name('Sheet2 - Numbers')
    dimensions = sheet2.calculate_dimension()
    eq_('%s%s:%s%s' % ('D', 1, 'AA', 30), dimensions)

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

def test_repair_central_directory():
    from openpyxl.reader.excel import repair_central_directory, CENTRAL_DIRECTORY_SIGNATURE

    data_a = "foobarbaz" + CENTRAL_DIRECTORY_SIGNATURE
    data_b = "bazbarfoo1234567890123456890"

    # The repair_central_directory looks for a magic set of bytes
    # (CENTRAL_DIRECTORY_SIGNATURE) and strips off everything 18 bytes past the sequence
    f = repair_central_directory(StringIO(data_a + data_b), True)
    eq_(f.read(), data_a + data_b[:18])

    f = repair_central_directory(StringIO(data_b), True)
    eq_(f.read(), data_b)


def test_read_no_theme():
    path = os.path.join(DATADIR, 'genuine', 'libreoffice_nrt.xlsx')
    wb = load_workbook(path)
    assert wb


class TestReadFormulae(object):

    xml_src = """<?xml version="1.0" ?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><dimension ref="A1:B6"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A6" sqref="A6"/></sheetView></sheetViews><sheetFormatPr baseColWidth="10" defaultColWidth="9.140625" defaultRowHeight="15" x14ac:dyDescent="0.25"/><cols><col min="1" max="1" width="15.7109375" customWidth="1"/><col min="2" max="2" width="15.28515625" customWidth="1"/></cols><sheetData><row r="1" spans="1:2" x14ac:dyDescent="0.25">
<c r="A1" t="s"><v>0</v></c>
<c r="B1" t="str"><f>CONCATENATE(A1,A2)</f><v>Hello, world!</v></c></row><row r="2" spans="1:2" x14ac:dyDescent="0.25">
<c r="A2" t="s"><v>1</v></c></row><row r="4" spans="1:2" x14ac:dyDescent="0.25">
<c r="A4"><v>1</v></c></row><row r="5" spans="1:2" x14ac:dyDescent="0.25">
<c r="A5"><v>2</v></c></row><row r="6" spans="1:2" x14ac:dyDescent="0.25">
<c r="A6"><f>SUM(A4:A5)</f><v>3</v></c></row></sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>"""

    @classmethod
    def setup(self):
        wb = Workbook()
        self.ws = wb.get_active_sheet()

    def test_fast_parse(self):
        from openpyxl.reader.worksheet import fast_parse
        fast_parse(self.ws, self.xml_src, {}, {}, None)
        b1 = self.ws.cell('B1')
        eq_(b1.data_type, 'f')
        eq_(b1.value, '=CONCATENATE(A1,A2)')
        a6 = self.ws.cell('A6')
        eq_(a6.data_type, 'f')
        eq_(a6.value, '=SUM(A4:A5)')


def test_read_complex_formulae():
    null_file = os.path.join(DATADIR, 'reader', 'formulae.xlsx')
    wb = load_workbook(null_file)
    ws = wb.get_active_sheet()

    # Test normal forumlae
    assert ws.cell('A1').data_type != 'f'
    assert ws.cell('A2').data_type != 'f'
    assert ws.cell('A3').data_type == 'f'
    assert 'A3' not in ws.formula_attributes
    assert ws.cell('A3').value == '=12345'
    assert ws.cell('A4').data_type == 'f'
    assert 'A4' not in ws.formula_attributes
    assert ws.cell('A4').value == '=A2+A3'
    assert ws.cell('A5').data_type == 'f'
    assert 'A5' not in ws.formula_attributes
    assert ws.cell('A5').value == '=SUM(A2:A4)'

    # Test shared forumlae
    assert ws.cell('B7').data_type == 'f'
    assert ws.formula_attributes['B7']['t'] == 'shared'
    assert ws.formula_attributes['B7']['si'] == '0'
    assert ws.formula_attributes['B7']['ref'] == 'B7:E7'
    assert ws.cell('B7').value == '=B4*2'
    assert ws.cell('C7').data_type == 'f'
    assert ws.formula_attributes['C7']['t'] == 'shared'
    assert ws.formula_attributes['C7']['si'] == '0'
    assert 'ref' not in ws.formula_attributes['C7']
    assert ws.cell('C7').value == '='
    assert ws.cell('D7').data_type == 'f'
    assert ws.formula_attributes['D7']['t'] == 'shared'
    assert ws.formula_attributes['D7']['si'] == '0'
    assert 'ref' not in ws.formula_attributes['D7']
    assert ws.cell('D7').value == '='
    assert ws.cell('E7').data_type == 'f'
    assert ws.formula_attributes['E7']['t'] == 'shared'
    assert ws.formula_attributes['E7']['si'] == '0'
    assert 'ref' not in ws.formula_attributes['E7']
    assert ws.cell('E7').value == '='

    # Test array forumlae
    assert ws.cell('C10').data_type == 'f'
    assert 'ref' not in ws.formula_attributes['C10']['ref']
    assert ws.formula_attributes['C10']['t'] == 'array'
    assert 'si' not in ws.formula_attributes['C10']
    assert ws.formula_attributes['C10']['ref'] == 'C10:C14'
    assert ws.cell('C10').value == '=SUM(A10:A14*B10:B14)'
    assert ws.cell('C11').data_type != 'f'


def test_data_only():
    null_file = os.path.join(DATADIR, 'reader', 'formulae.xlsx')
    wb = load_workbook(null_file, data_only=True)
    ws = wb.get_active_sheet()
    ws.parent.data_only = True
    # Test cells returning values only, not formulae
    assert ws.formula_attributes == {}
    assert ws.cell('A2').data_type == 'n' and ws.cell('A2').value == 12345
    assert ws.cell('A3').data_type == 'n' and ws.cell('A3').value == 12345
    assert ws.cell('A4').data_type == 'n' and ws.cell('A4').value == 24690
    assert ws.cell('A5').data_type == 'n' and ws.cell('A5').value == 49380


def test_read_contains_chartsheet():
    """
    Test reading workbook containing chartsheet.

    "contains_chartsheets.xlsx" has the following sheets:
    +---+------------+------------+
    | # | Name       | Type       |
    +===+============+============+
    | 1 | "data"     | worksheet  |
    +---+------------+------------+
    | 2 | "chart"    | chartsheet |
    +---+------------+------------+
    | 3 | "moredata" | worksheet  |
    +---+------------+------------+
    """
    # test data
    path = os.path.join(DATADIR, 'reader', 'contains_chartsheets.xlsx')
    wb = load_workbook(path)
    # workbook contains correct sheet names
    sheet_names = wb.get_sheet_names()
    eq_(sheet_names[0], 'data')
    eq_(sheet_names[1], 'moredata')


def test_guess_types():
    filename = os.path.join(DATADIR, 'genuine', 'guess_types.xlsx')
    for guess, dtype in ((True, float), (False, unicode)):
        wb = load_workbook(filename, guess_types=guess)
        ws = wb.get_active_sheet()
        assert isinstance(ws.cell('D2').value, dtype), 'wrong dtype (%s) when guess type is: %s (%s instead)' % (dtype, guess, type(ws.cell('A1').value))


def test_get_xml_iter():
    #1 file object
    #2 stream (file-like)
    #3 string
    #4 zipfile
    from openpyxl.reader.worksheet import _get_xml_iter
    from tempfile import TemporaryFile
    FUT = _get_xml_iter
    s = ""
    stream = FUT(s)
    assert isinstance(stream, BytesIO), type(stream)

    u = unicode(s)
    stream = FUT(u)
    assert isinstance(stream, BytesIO), type(stream)

    f = TemporaryFile(mode='rb+', prefix='openpyxl.', suffix='.unpack.temp')
    stream = FUT(f)
    assert isinstance(stream, tempfile), type(stream)
    f.close()

    from zipfile import ZipFile
    t = TemporaryFile()
    z = ZipFile(t, mode="w")
    z.writestr("test", "whatever")
    stream = FUT(z.open("test"))
    assert hasattr(stream, "read")
    z.close()
