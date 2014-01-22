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

# stdlib imports
import decimal
import os.path
import zipfile

# compatibility imports
from openpyxl.compat import BytesIO

import pytest

# package imports
from openpyxl.tests.helper import (
    DATADIR,
    compare_xml,
    )
from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook
from openpyxl.writer.excel import (
    save_workbook,
    save_virtual_workbook,
    ExcelWriter
    )
from openpyxl.writer.workbook import write_workbook, write_workbook_rels
from openpyxl.writer.worksheet import write_worksheet, write_worksheet_rels
from openpyxl.writer.strings import write_string_table
from openpyxl.writer.styles import StyleWriter

def test_write_empty_workbook(tmpdir):
    tmpdir.chdir()
    wb = Workbook()
    dest_filename = 'empty_book.xlsx'
    save_workbook(wb, dest_filename)
    assert os.path.isfile(dest_filename)


def test_write_virtual_workbook():
    old_wb = Workbook()
    saved_wb = save_virtual_workbook(old_wb)
    new_wb = load_workbook(BytesIO(saved_wb))
    assert new_wb


def test_write_workbook_rels():
    wb = Workbook()
    content = write_workbook_rels(wb)
    reference_file = os.path.join(DATADIR, 'writer', 'expected', 'workbook.xml.rels')
    with open(reference_file) as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff


def test_write_workbook():
    wb = Workbook()
    content = write_workbook(wb)
    reference_file = os.path.join(DATADIR, 'writer', 'expected', 'workbook.xml')
    with open(reference_file) as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff


def test_write_string_table():
    table = {'hello': 1, 'world': 2, 'nice': 3}
    content = write_string_table(table)
    reference_file = os.path.join(DATADIR, 'writer', 'expected', 'sharedStrings.xml')
    with open(reference_file) as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff


def test_write_worksheet():
    wb = Workbook()
    ws = wb.create_sheet()
    ws.cell('F42').value = 'hello'
    content = write_worksheet(ws, {'hello': 0}, {})
    reference_file = os.path.join(DATADIR, 'writer', 'expected', 'sheet1.xml')
    with open(reference_file) as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff


def test_write_hidden_worksheet():
    wb = Workbook()
    ws = wb.create_sheet()
    ws.sheet_state = ws.SHEETSTATE_HIDDEN
    ws.cell('F42').value = 'hello'
    content = write_worksheet(ws, {'hello': 0}, {})
    reference_file = os.path.join(DATADIR, 'writer', 'expected', 'sheet1.xml')
    with open(reference_file) as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff


def test_write_bool():
    wb = Workbook()
    ws = wb.create_sheet()
    ws.cell('F42').value = False
    ws.cell('F43').value = True
    content = write_worksheet(ws, {}, {})
    reference_file = os.path.join(DATADIR, 'writer', 'expected', 'sheet1_bool.xml')
    with open(reference_file) as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff


def test_write_formula():
    wb = Workbook()
    ws = wb.create_sheet()
    ws.cell('F1').value = 10
    ws.cell('F2').value = 32
    ws.cell('F3').value = '=F1+F2'
    ws.cell('A4').value = '=A1+A2+A3'
    ws.formula_attributes['A4'] = {'t': 'shared', 'ref': 'A4:C4', 'si': '0'}
    ws.cell('B4').value = '='
    ws.formula_attributes['B4'] = {'t': 'shared', 'si': '0'}
    ws.cell('C4').value = '='
    ws.formula_attributes['C4'] = {'t': 'shared', 'si': '0'}
    content = write_worksheet(ws, {}, {})
    reference_file = os.path.join(DATADIR, 'writer', 'expected', 'sheet1_formula.xml')
    with open(reference_file) as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff


def test_write_style():
    wb = Workbook()
    ws = wb.create_sheet()
    ws.cell('F1').value = '13%'
    ws._styles['F'] = ws._styles['F1']
    style_id_by_hash = StyleWriter(wb).get_style_by_hash()
    content = write_worksheet(ws, {}, style_id_by_hash)
    reference_file = os.path.join(DATADIR, 'writer', 'expected', 'sheet1_style.xml')
    with open(reference_file) as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff


def test_write_height():
    wb = Workbook()
    ws = wb.create_sheet()
    ws.cell('F1').value = 10
    ws.row_dimensions[ws.cell('F1').row].height = 30
    content = write_worksheet(ws, {}, {})
    reference_file = os.path.join(DATADIR, 'writer', 'expected', 'sheet1_height.xml')
    with open(reference_file) as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff


def test_write_hyperlink():
    wb = Workbook()
    ws = wb.create_sheet()
    ws.cell('A1').value = "test"
    ws.cell('A1').hyperlink = "http://test.com"
    content = write_worksheet(ws, {'test': 0}, {})
    reference_file = os.path.join(DATADIR, 'writer', 'expected', 'sheet1_hyperlink.xml')
    with open(reference_file) as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff


def test_write_hyperlink_rels():
    wb = Workbook()
    ws = wb.create_sheet()
    assert 0 == len(ws.relationships)
    ws.cell('A1').value = "test"
    ws.cell('A1').hyperlink = "http://test.com/"
    assert 1 == len(ws.relationships)
    ws.cell('A2').value = "test"
    ws.cell('A2').hyperlink = "http://test2.com/"
    assert 2 == len(ws.relationships)
    content = write_worksheet_rels(ws, 1, 1)
    reference_file = os.path.join(DATADIR, 'writer', 'expected', 'sheet1_hyperlink.xml.rels')
    with open(reference_file) as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff


@pytest.mark.xfail
@pytest.mark.pil_required
def test_write_hyperlink_image_rels(Workbook, Image):
    wb = Workbook()
    ws = wb.create_sheet()
    ws.cell('A1').value = "test"
    ws.cell('A1').hyperlink = "http://test.com/"
    img = os.path.join(DATADIR, "plain.png")
    i = Image(img)
    ws.add_image(i)
    raise ValueError("Resulting file is invalid")
    # TODO write integration test with duplicate relation ids then fix



def test_hyperlink_value():
    wb = Workbook()
    ws = wb.create_sheet()
    ws.cell('A1').hyperlink = "http://test.com"
    assert "http://test.com" == ws.cell('A1').value
    ws.cell('A1').value = "test"
    assert "test" == ws.cell('A1').value


def test_write_auto_filter():
    wb = Workbook()
    ws = wb.worksheets[0]
    ws.cell('F42').value = 'hello'
    ws.auto_filter.ref = 'A1:F1'
    content = write_worksheet(ws, {'hello': 0}, {})
    reference_file = os.path.join(DATADIR, 'writer', 'expected', 'sheet1_auto_filter.xml')
    with open(reference_file) as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None

    content = write_workbook(wb)
    reference_file = os.path.join(DATADIR, 'writer', 'expected', 'workbook_auto_filter.xml')
    with open(reference_file) as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff

def test_write_auto_filter_filter_column():
    wb = Workbook()
    ws = wb.worksheets[0]
    ws.cell('F42').value = 'hello'
    ws.auto_filter.ref = 'A1:F1'
    ws.auto_filter.add_filter_column(0, ["0"], blank=True)
    content = write_worksheet(ws, {'hello': 0}, {})
    reference_file = os.path.join(DATADIR, 'writer', 'expected', 'sheet1_auto_filter_filter_column.xml')
    with open(reference_file) as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None

    content = write_workbook(wb)
    reference_file = os.path.join(DATADIR, 'writer', 'expected', 'workbook_auto_filter.xml')
    with open(reference_file) as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff

def test_write_auto_filter_sort_condition():
    wb = Workbook()
    ws = wb.worksheets[0]
    ws.cell('A1').value = 'header'
    ws.cell('A2').value = 1
    ws.cell('A3').value = 0
    ws.auto_filter.ref = 'A2:A3'
    ws.auto_filter.add_sort_condition('A2:A3', descending=True)
    content = write_worksheet(ws, {'header': 0}, {})
    reference_file = os.path.join(DATADIR, 'writer', 'expected', 'sheet1_auto_filter_sort_condition.xml')
    with open(reference_file) as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None

    content = write_workbook(wb)
    reference_file = os.path.join(DATADIR, 'writer', 'expected', 'workbook_auto_filter.xml')
    with open(reference_file) as expected:
        diff = compare_xml(content, expected.read())

def test_freeze_panes_horiz():
    wb = Workbook()
    ws = wb.create_sheet()
    ws.cell('F42').value = 'hello'
    ws.freeze_panes = 'A4'
    content = write_worksheet(ws, {'hello': 0}, {})
    reference_file = os.path.join(DATADIR, 'writer', 'expected', 'sheet1_freeze_panes_horiz.xml')
    with open(reference_file) as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff

def test_freeze_panes_vert():
    wb = Workbook()
    ws = wb.create_sheet()
    ws.cell('F42').value = 'hello'
    ws.freeze_panes = 'D1'
    content = write_worksheet(ws, {'hello': 0}, {})
    reference_file = os.path.join(DATADIR, 'writer', 'expected', 'sheet1_freeze_panes_vert.xml')
    with open(reference_file) as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff

def test_freeze_panes_both():
    wb = Workbook()
    ws = wb.create_sheet()
    ws.cell('F42').value = 'hello'
    ws.freeze_panes = 'D4'
    content = write_worksheet(ws, {'hello': 0}, {})
    reference_file = os.path.join(DATADIR, 'writer', 'expected', 'sheet1_freeze_panes_both.xml')
    with open(reference_file) as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff

def test_long_number():
    wb = Workbook()
    ws = wb.create_sheet()
    ws.cell('A1').value = 9781231231230
    content = write_worksheet(ws, {}, {})
    reference_file = os.path.join(DATADIR, 'writer', 'expected', 'long_number.xml')
    with open(reference_file) as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff

def test_decimal():
    wb = Workbook()
    ws = wb.create_sheet()
    ws.cell('A1').value = decimal.Decimal('3.14')
    content = write_worksheet(ws, {}, {})
    reference_file = os.path.join(DATADIR, 'writer', 'expected', 'decimal.xml')
    with open(reference_file) as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff

def test_short_number():
    wb = Workbook()
    ws = wb.create_sheet()
    ws.cell('A1').value = 1234567890
    content = write_worksheet(ws, {}, {})
    reference_file = os.path.join(DATADIR, 'writer', 'expected', 'short_number.xml')
    with open(reference_file) as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff


@pytest.mark.pil_required
def test_write_images():
    wb = Workbook()
    ew = ExcelWriter(workbook=wb)
    from openpyxl.drawing import Image
    imagepath = os.path.join(DATADIR, "plain.png")
    img = Image(imagepath)

    buf = BytesIO()

    archive = zipfile.ZipFile(buf, 'w')
    ew._write_images([img], archive, 1)
    archive.close()

    buf.seek(0)
    archive = zipfile.ZipFile(buf, 'r')
    zipinfo = archive.infolist()
    assert len(zipinfo) == 1
    assert zipinfo[0].filename == 'xl/media/image1.png'
