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
from datetime import time, datetime, timedelta

# 3rd party imports
from nose.tools import eq_, raises, assert_raises #pylint: disable=E0611
import pytest

# package imports
from openpyxl.worksheet import Worksheet
from openpyxl.workbook import Workbook
from openpyxl.shared.exc import ColumnStringIndexException, \
        CellCoordinatesException, DataTypeException
from openpyxl.shared.date_time import CALENDAR_WINDOWS_1900
from openpyxl.cell import column_index_from_string, \
        coordinate_from_string, get_column_letter, Cell, absolute_coordinate
from openpyxl.comments import Comment

import decimal

def build_dummy_worksheet():

    class Ws(object):
        class Wb(object):
            excel_base_date = CALENDAR_WINDOWS_1900
        encoding = 'utf-8'
        parent = Wb()
        title = "Dummy Worksheet"

    return Ws()


def test_coordinates():
    column, row = coordinate_from_string('ZF46')
    eq_("ZF", column)
    eq_(46, row)


@raises(CellCoordinatesException)
def test_invalid_coordinate():
    coordinate_from_string('AAA')

@raises(CellCoordinatesException)
def test_zero_row():
    coordinate_from_string('AQ0')

def test_absolute():
    eq_('$ZF$51', absolute_coordinate('ZF51'))

def test_absolute_multiple():

    eq_('$ZF$51:$ZF$53', absolute_coordinate('ZF51:ZF$53'))

@pytest.mark.parametrize("column, idx",
                         [
                         ('j', 10),
                         ('Jj', 270),
                         ('JJj', 7030),
                         ('A', 1),
                         ('Z', 26),
                         ('AA', 27),
                         ('AZ', 52),
                         ('BA', 53),
                         ('BZ',  78),
                         ('ZA',  677),
                         ('ZZ',  702),
                         ('AAA',  703),
                         ('AAZ',  728),
                         ('ABC',  731),
                         ('AZA', 1353),
                         ('ZZA', 18253),
                         ('ZZZ', 18278),
                         ]
                         )
def test_column_index(column, idx):
    assert column_index_from_string(column) == idx


@pytest.mark.parametrize("column",
                         ('JJJJ', '', '$', '1',)
                         )
def test_bad_column_index(column):
    with pytest.raises(ValueError):
        column_index_from_string(column)


@pytest.mark.parametrize("value", (0, 18729))
def test_column_letter_boundries(value):
    with pytest.raises(ValueError):
        get_column_letter(value)

@pytest.mark.parametrize("value, expected",
                         [
                        (18278, "ZZZ"),
                        (7030, "JJJ"),
                        (28, "AB"),
                        (27, "AA"),
                        (26, "Z")
                         ]
                         )
def test_column_letter(value, expected):
    assert get_column_letter(value) == expected


def test_initial_value():
    ws = build_dummy_worksheet()
    cell = Cell(ws, 'A', 1, value='17.5')
    eq_(cell.TYPE_NUMERIC, cell.data_type)


class TestCellValueTypes(object):

    @classmethod
    def setup_class(cls):

        ws = build_dummy_worksheet()
        cls.cell = Cell(ws, 'A', 1)

    def test_1st(self):
        eq_(self.cell.TYPE_NULL, self.cell.data_type)

    def test_null(self):
        self.cell.value = None
        eq_(self.cell.TYPE_NULL, self.cell.data_type)

    def test_numeric(self):

        def check_numeric(value):
            self.cell.value = value
            eq_(self.cell.TYPE_NUMERIC, self.cell.data_type)

        values = (42, '4.2', '-42.000', '0', 0, 0.0001, '0.9999', '99E-02', 1e1, '4', '-1E3', 4, decimal.Decimal('3.14'))
        for value in values:
            yield check_numeric, value

    def test_string(self):
        self.cell.value = 'hello'
        eq_(self.cell.TYPE_STRING, self.cell.data_type)

    def test_single_dot(self):
        self.cell.value = '.'
        eq_(self.cell.TYPE_STRING, self.cell.data_type)

    def test_formula(self):
        self.cell.value = '=42'
        eq_(self.cell.TYPE_FORMULA, self.cell.data_type)
        self.cell.value = '=if(A1<4;-1;1)'
        eq_(self.cell.TYPE_FORMULA, self.cell.data_type)

    def test_boolean(self):
        self.cell.value = True
        eq_(self.cell.TYPE_BOOL, self.cell.data_type)
        self.cell.value = False
        eq_(self.cell.TYPE_BOOL, self.cell.data_type)

    def test_leading_zero(self):
        self.cell.value = '0800'
        eq_(self.cell.TYPE_STRING, self.cell.data_type)

    def test_error_codes(self):

        def check_error(cell):
            eq_(cell.TYPE_ERROR, cell.data_type)

        for error_string in self.cell.ERROR_CODES.keys():
            self.cell.value = error_string
            yield check_error, self.cell


def test_data_type_check():
    ws = build_dummy_worksheet()
    cell = Cell(ws, 'A', 1)
    cell.bind_value(None)
    eq_(Cell.TYPE_NULL, cell._data_type)

    cell.bind_value('.0e000')
    eq_(Cell.TYPE_NUMERIC, cell._data_type)

    cell.bind_value('-0.e-0')
    eq_(Cell.TYPE_NUMERIC, cell._data_type)

    cell.bind_value('1E')
    eq_(Cell.TYPE_STRING, cell._data_type)

@raises(DataTypeException)
def test_set_bad_type():
    ws = build_dummy_worksheet()
    cell = Cell(ws, 'A', 1)
    cell.set_explicit_value(1, 'q')


def test_time():

    def check_time(raw_value, coerced_value):
        cell.value = raw_value
        eq_(cell.value, coerced_value)
        eq_(cell.TYPE_NUMERIC, cell.data_type)

    wb = Workbook()
    ws = Worksheet(wb)
    cell = Cell(ws, 'A', 1)
    values = (('03:40:16', time(3, 40, 16)), ('03:40', time(3, 40)),)
    for raw_value, coerced_value in values:
        yield check_time, raw_value, coerced_value


def test_timedelta():

    wb = Workbook()
    ws = Worksheet(wb)
    cell = Cell(ws, 'A', 1)
    cell.value = timedelta(days=1, hours=3)
    eq_(cell.value, 1.125)
    eq_(cell.TYPE_NUMERIC, cell.data_type)


def test_date_format_on_non_date():
    wb = Workbook()
    ws = Worksheet(wb)
    cell = Cell(ws, 'A', 1)
    cell.value = datetime.now()
    cell.value = 'testme'
    eq_('testme', cell.value)

def test_set_get_date():
    today = datetime(2010, 1, 18, 14, 15, 20, 1600)
    wb = Workbook()
    ws = Worksheet(wb)
    cell = Cell(ws, 'A', 1)
    cell.value = today
    eq_(today, cell.value)

def test_repr():
    wb = Workbook()
    ws = Worksheet(wb)
    cell = Cell(ws, 'A', 1)
    eq_(repr(cell), '<Cell Sheet1.A1>', 'Got bad repr: %s' % repr(cell))

def test_is_date():
    wb = Workbook()
    ws = Worksheet(wb)
    cell = Cell(ws, 'A', 1)
    cell.value = datetime.now()
    eq_(cell.is_date(), True)
    cell.value = 'testme'
    eq_('testme', cell.value)
    assert cell.is_date() is False

def test_is_not_date_color_format():

    wb = Workbook()
    ws = Worksheet(wb)
    cell = Cell(ws, 'A', 1)

    cell.value = -13.5
    cell.style.number_format.format_code = '0.00_);[Red]\(0.00\)'

    assert cell.is_date() is False

def test_comment_count():
    wb = Workbook()
    ws = Worksheet(wb)
    cell = ws.cell(coordinate="A1")
    assert ws._comment_count == 0
    cell.comment = Comment("text", "author")
    assert ws._comment_count == 1
    cell.comment = Comment("text", "author")
    assert ws._comment_count == 1
    cell.comment = None
    assert ws._comment_count == 0
    cell.comment = None
    assert ws._comment_count == 0

def test_comment_assignment():
    wb = Workbook()
    ws = Worksheet(wb)
    c = Comment("text", "author")
    ws.cell(coordinate="A1").comment = c
    with pytest.raises(AttributeError):
        ws.cell(coordinate="A2").commment = c
    ws.cell(coordinate="A2").comment = Comment("text2", "author2")
    with pytest.raises(AttributeError):
        ws.cell(coordinate="A1").comment = ws.cell(coordinate="A2").comment
    # this should orphan c, so that assigning it to A2 does not raise AttributeError
    ws.cell(coordinate="A1").comment = None
    ws.cell(coordinate="A2").comment = c