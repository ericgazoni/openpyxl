# file openpyxl/tests/test_workbook.py

# 3rd party imports
from nose.tools import eq_

# package imports
from openpyxl.workbook import Workbook
from openpyxl.namedrange import NamedRange


def test_get_active_sheet():
    wb = Workbook()
    active_sheet = wb.get_active_sheet()
    eq_(active_sheet, wb.worksheets[0])


def test_create_sheet():
    wb = Workbook()
    new_sheet = wb.create_sheet(0)
    eq_(new_sheet, wb.worksheets[0])


def test_remove_sheet():
    wb = Workbook()
    new_sheet = wb.create_sheet(0)
    wb.remove_sheet(new_sheet)
    assert new_sheet not in wb.worksheets


def test_get_sheet_by_name():
    wb = Workbook()
    new_sheet = wb.create_sheet()
    title = 'my sheet'
    new_sheet.title = title
    found_sheet = wb.get_sheet_by_name(title)
    eq_(new_sheet, found_sheet)


def test_get_index():
    wb = Workbook()
    new_sheet = wb.create_sheet(0)
    sheet_index = wb.get_index(new_sheet)
    eq_(sheet_index, 0)


def test_get_sheet_names():
    wb = Workbook()
    names = ['Sheet', 'Sheet1', 'Sheet2', 'Sheet3', 'Sheet4', 'Sheet5']
    for count in range(5):
        wb.create_sheet(0)
    actual_names = wb.get_sheet_names()
    eq_(sorted(actual_names), sorted(names))


def test_get_named_ranges():
    wb = Workbook()
    eq_(wb.get_named_ranges(), wb._named_ranges)


def test_add_named_range():
    wb = Workbook()
    new_sheet = wb.create_sheet()
    named_range = NamedRange('test_nr', new_sheet, 'A1')
    wb.add_named_range(named_range)
    named_ranges_list = wb.get_named_ranges()
    assert named_range in named_ranges_list


def test_get_named_range():
    wb = Workbook()
    new_sheet = wb.create_sheet()
    named_range = NamedRange('test_nr', new_sheet, 'A1')
    wb.add_named_range(named_range)
    found_named_range = wb.get_named_range('test_nr')
    eq_(named_range, found_named_range)


def test_remove_named_range():
    wb = Workbook()
    new_sheet = wb.create_sheet()
    named_range = NamedRange('test_nr', new_sheet, 'A1')
    wb.add_named_range(named_range)
    wb.remove_named_range(named_range)
    named_ranges_list = wb.get_named_ranges()
    assert named_range not in named_ranges_list
