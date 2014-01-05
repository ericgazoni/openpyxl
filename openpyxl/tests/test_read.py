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
from datetime import datetime
import zipfile

import pytest

# compatibility imports
from openpyxl.shared.compat import BytesIO, StringIO, unicode, tempfile

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
    assert ws.cell('G5').value == 'hello'
    assert ws.cell('D30').value == 30
    assert ws.cell('K9').value == 0.09


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
    assert 'This is cell G5' == sheet2['G5'].value
    assert 18 == sheet2['D18'].value
    assert sheet2['G9'].value is True
    assert sheet2['G10'].value is False

def test_read_nostring_workbook():
    genuine_wb = os.path.join(DATADIR, 'genuine', 'empty-no-string.xlsx')
    wb = load_workbook(genuine_wb)
    assert isinstance(wb, Workbook)

def test_read_empty_file():
    null_file = os.path.join(DATADIR, 'reader', 'null_file.xlsx')
    with pytest.raises(InvalidFileException):
        load_workbook(null_file)

def test_read_empty_archive():
    null_file = os.path.join(DATADIR, 'reader', 'null_archive.xlsx')
    with pytest.raises(InvalidFileException):
        load_workbook(null_file)

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
    assert '%s%s:%s%s' % ('D', 1, 'AA', 30) == dimensions

def test_get_highest_row_iter():
    path = os.path.join(DATADIR, 'genuine', 'empty.xlsx')
    wb = load_workbook(filename=path, use_iterators=True)
    sheet2 = wb.get_sheet_by_name('Sheet2 - Numbers')
    max_row = sheet2.get_highest_row()
    assert 30 == max_row

@pytest.mark.xfail
def test_read_workbook_with_no_properties():
    genuine_wb = os.path.join(DATADIR, 'genuine', \
                'empty_with_no_properties.xlsx')
    load_workbook(filename=genuine_wb)


class TestReadWorkbookWithStyles(object):

    @classmethod
    def setup_class(cls):
        cls.genuine_wb = os.path.join(DATADIR, 'genuine', \
                'empty-with-styles.xlsx')
        wb = load_workbook(cls.genuine_wb)
        cls.ws = wb.get_sheet_by_name('Sheet1')

    def test_read_general_style(self):
        assert self.ws.cell('A1').style.number_format.format_code == NumberFormat.FORMAT_GENERAL

    def test_read_date_style(self):
        assert self.ws.cell('A2').style.number_format.format_code == NumberFormat.FORMAT_DATE_XLSX14

    def test_read_number_style(self):
        assert self.ws.cell('A3').style.number_format.format_code == NumberFormat.FORMAT_NUMBER_00

    def test_read_time_style(self):
        assert self.ws.cell('A4').style.number_format.format_code == NumberFormat.FORMAT_DATE_TIME3

    def test_read_percentage_style(self):
        assert self.ws.cell('A5').style.number_format.format_code == NumberFormat.FORMAT_PERCENTAGE_00


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
        assert self.win_wb.properties.excel_base_date == CALENDAR_WINDOWS_1900

    def test_read_mac_base_date(self):
        assert self.mac_wb.properties.excel_base_date == CALENDAR_MAC_1904

    def test_read_date_style_mac(self):
        assert self.mac_ws.cell('A1').style.number_format.format_code ==                 NumberFormat.FORMAT_DATE_XLSX14

    def test_read_date_style_win(self):
        assert self.win_ws.cell('A1').style.number_format.format_code ==                 NumberFormat.FORMAT_DATE_XLSX14

    def test_read_date_value(self):
        datetuple = (2011, 10, 31)
        dt = datetime(datetuple[0], datetuple[1], datetuple[2])
        assert self.mac_ws.cell('A1').value == dt
        assert self.win_ws.cell('A1').value == dt
        assert self.mac_ws.cell('A1').value == self.win_ws.cell('A1').value

def test_repair_central_directory():
    from openpyxl.reader.excel import repair_central_directory, CENTRAL_DIRECTORY_SIGNATURE

    data_a = "foobarbaz" + CENTRAL_DIRECTORY_SIGNATURE
    data_b = "bazbarfoo1234567890123456890"

    # The repair_central_directory looks for a magic set of bytes
    # (CENTRAL_DIRECTORY_SIGNATURE) and strips off everything 18 bytes past the sequence
    f = repair_central_directory(StringIO(data_a + data_b), True)
    assert f.read() == data_a + data_b[:18]

    f = repair_central_directory(StringIO(data_b), True)
    assert f.read() == data_b


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
        assert b1.data_type == 'f'
        assert b1.value, '=CONCATENATE(A1 == A2)'
        a6 = self.ws.cell('A6')
        assert a6.data_type == 'f'
        assert a6.value == '=SUM(A4:A5)'


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


workbooks = [
    ("bug137.xlsx", [
        {'path': 'worksheets/sheet1.xml', 'title': 'Sheet1'}
        ]
     ),
    ("contains_chartsheets.xlsx", [
        {'path': 'worksheets/sheet1.xml', 'title': 'data'},
        {'path': 'worksheets/sheet2.xml', 'title': 'moredata'}
        ])
            ]
@pytest.mark.parametrize("excel_file, expected", workbooks)
def test_read_contains_chartsheet(excel_file, expected):
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
    path = os.path.join(DATADIR, 'reader', excel_file)
    wb = load_workbook(path)
    sheet_names = wb.get_sheet_names()
    assert sheet_names == [sheet['title'] for sheet in expected]


@pytest.mark.parametrize("excel_file, expected", workbooks)
def test_detect_worksheets(excel_file, expected):
    from openpyxl.reader.excel import detect_worksheets
    fname = os.path.join(DATADIR, "reader", excel_file)
    archive = zipfile.ZipFile(fname)
    assert list(detect_worksheets(archive)) == expected


def test_read_rels():
    from openpyxl.reader.workbook import read_rels
    from openpyxl.shared.ooxml import ARC_WORKBOOK_RELS
    fname = os.path.join(DATADIR, "reader", "bug137.xlsx")
    archive = zipfile.ZipFile(fname)
    assert read_rels(archive.read(ARC_WORKBOOK_RELS)) == {
        'rId1': {'path': 'chartsheets/sheet1.xml'},
        'rId2': {'path': 'worksheets/sheet1.xml'},
        'rId3': {'path': 'theme/theme1.xml'},
        'rId4': {'path': 'styles.xml'},
        'rId5': {'path': 'sharedStrings.xml'}
    }


def test_read_content_types():
    from openpyxl.reader.workbook import read_content_types
    from openpyxl.shared.ooxml import ARC_CONTENT_TYPES
    fname = os.path.join(DATADIR, "reader", "contains_chartsheets.xlsx")
    archive = zipfile.ZipFile(fname)
    assert list(read_content_types(archive.read(ARC_CONTENT_TYPES))) == [
    ('/xl/workbook.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'),
    ('/xl/worksheets/sheet1.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'),
    ('/xl/chartsheets/sheet1.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml'),
    ('/xl/worksheets/sheet2.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'),
    ('/xl/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'),
    ('/xl/styles.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'),
    ('/xl/sharedStrings.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'),
    ('/xl/drawings/drawing1.xml', 'application/vnd.openxmlformats-officedocument.drawing+xml'),
    ('/xl/charts/chart1.xml', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'),
    ('/xl/drawings/drawing2.xml', 'application/vnd.openxmlformats-officedocument.drawing+xml'),
    ('/xl/charts/chart2.xml', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'),
    ('/xl/calcChain.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml'),
    ('/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'),
    ('/docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml')
    ]


def test_read_sheets():
    from openpyxl.reader.workbook import read_sheets
    from openpyxl.shared.ooxml import ARC_WORKBOOK
    fname = os.path.join(DATADIR, "reader", "bug137.xlsx")
    archive = zipfile.ZipFile(fname)
    assert list(read_sheets(archive.read(ARC_WORKBOOK))) == [("Chart1", "rId1"),
                                                             ("Sheet1", "rId2")
                                                             ]


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
