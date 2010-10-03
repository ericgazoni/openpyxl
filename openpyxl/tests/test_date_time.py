# file openpyxl/tests/test_date_time.py

# Python stdlib imports
from datetime import datetime

# 3rd party imports
from nose.tools import eq_, assert_raises

# package imports
from openpyxl.worksheet import Worksheet
from openpyxl.workbook import Workbook
from openpyxl.cell import Cell
from openpyxl.shared.date_time import SharedDate


def test_date_format_on_non_date():
    wb = Workbook()
    ws = Worksheet(wb)
    cell = Cell(ws, 'A', 1)

    def check_date_pair(count, date_string):
        cell.value = datetime.strptime(date_string, '%Y-%m-%d')
        eq_(count, cell._value)

    date_pairs = (
        (15, '1900-01-15'),
        (59, '1900-02-28'),
        (61, '1900-03-01'),
        (367, '1901-01-01'),
        (2958465, '9999-12-31'), )
    for count, date_string in date_pairs:
        yield check_date_pair, count, date_string


def test_1900_leap_year():
    shared_date = SharedDate()
    assert_raises(ValueError, shared_date.from_julian, 60)
    assert_raises(ValueError, shared_date.to_julian, 1900, 2, 29)


def test_bad_date():
    shared_date = SharedDate()

    def check_bad_date(year, month, day):
        assert_raises(ValueError, shared_date.to_julian, year, month, day)

    bad_dates = ((1776, 07, 04), (1899, 12, 31), )
    for year, month, day in bad_dates:
        yield check_bad_date, year, month, day


def test_bad_julian_date():
    shared_date = SharedDate()
    assert_raises(ValueError, shared_date.from_julian, -1)


def test_mac_date():
    shared_date = SharedDate()
    shared_date.excel_base_date = shared_date.CALENDAR_MAC_1904
    assert_raises(NotImplementedError, shared_date.to_julian, 2000, 1, 1)
