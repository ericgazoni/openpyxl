# Copyright (c) 2010-2014 openpyxl
#

# Python stdlib imports
from datetime import time, datetime, timedelta, date

# 3rd party imports
import pytest

# package imports
from openpyxl.compat import safe_string
from openpyxl.worksheet import Worksheet
from openpyxl.workbook import Workbook
from openpyxl.exceptions import (
    CellCoordinatesException,
    DataTypeException
    )
from openpyxl.date_time import CALENDAR_WINDOWS_1900
from openpyxl.cell import (
    column_index_from_string,
    coordinate_from_string,
    get_column_letter,
    Cell,
    absolute_coordinate
    )
from openpyxl.comments import Comment
from openpyxl.styles import NumberFormat

import decimal

@pytest.fixture
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
    assert "ZF" == column
    assert 46 == row


def test_invalid_coordinate():
    with pytest.raises(CellCoordinatesException):
        coordinate_from_string('AAA')

def test_zero_row():
    with pytest.raises(CellCoordinatesException):
        coordinate_from_string('AQ0')

def test_absolute():
    assert '$ZF$51' == absolute_coordinate('ZF51')

def test_absolute_multiple():

    assert '$ZF$51:$ZF$53' == absolute_coordinate('ZF51:ZF$53')

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
    assert cell.TYPE_NUMERIC == cell.data_type


class TestCellValueTypes(object):

    @classmethod
    def setup_class(cls):
        wb = Workbook()
        ws = Worksheet(wb)
        cls.cell = Cell(ws, 'A', 1)

    def test_1st(self):
        assert self.cell.TYPE_NULL == self.cell.data_type

    def test_null(self):
        self.cell.value = None
        assert self.cell.TYPE_NULL == self.cell.data_type

    @pytest.mark.parametrize("value, expected",
                             [
                                 (42, 42),
                                 ('4.2', 4.2),
                                 ('-42.000', -42),
                                 ( '0', 0),
                                 (0, 0),
                                 ( 0.0001, 0.0001),
                                 ('0.9999', 0.9999),
                                 ('99E-02', 0.99),
                                 ( 1e1, 10),
                                 ('4', 4),
                                 ('-1E3', -1000),
                                 (4, 4),
                                 (decimal.Decimal('3.14'), decimal.Decimal('3.14')),
                             ]
                            )
    def test_numeric(self, value, expected):
        self.cell.value = value
        assert self.cell.internal_value == expected
        assert self.cell.TYPE_NUMERIC == self.cell.data_type

    def test_string(self):
        self.cell.value = 'hello'
        assert self.cell.TYPE_STRING == self.cell.data_type

    def test_single_dot(self):
        self.cell.value = '.'
        assert self.cell.TYPE_STRING == self.cell.data_type

    def test_formula(self):
        self.cell.value = '=42'
        assert self.cell.TYPE_FORMULA == self.cell.data_type
        self.cell.value = '=if(A1<4;-1;1)'
        assert self.cell.TYPE_FORMULA == self.cell.data_type

    def test_boolean(self):
        self.cell.value = True
        assert self.cell.TYPE_BOOL == self.cell.data_type
        self.cell.value = False
        assert self.cell.TYPE_BOOL == self.cell.data_type

    def test_leading_zero(self):
        self.cell.value = '0800'
        assert self.cell.TYPE_STRING == self.cell.data_type

    @pytest.mark.parametrize("error_string", Cell.ERROR_CODES.keys())
    def test_error_codes(self, error_string):
        self.cell.value = error_string
        assert self.cell.TYPE_ERROR == self.cell.data_type

    def test_insert_float(self):
        self.cell.value = 3.14
        assert Cell.TYPE_NUMERIC == self.cell.data_type

    def test_insert_percentage(self):
        self.cell.value = '3.14%'
        assert Cell.TYPE_NUMERIC == self.cell.data_type
        assert safe_string(0.0314) == safe_string(self.cell.internal_value)

    def test_insert_datetime(self):
        self.cell.value = date.today()
        assert Cell.TYPE_NUMERIC == self.cell.data_type

    def test_insert_date(self):
        self.cell.value = datetime.now()
        assert Cell.TYPE_NUMERIC == self.cell.data_type

    def test_internal_date(self):
        dt = datetime(2010, 7, 13, 6, 37, 41)
        self.cell.value = dt
        assert 40372.27616898148 == self.cell.internal_value

    def test_datetime_interpretation(self):
        dt = datetime(2010, 7, 13, 6, 37, 41)
        self.cell.value = dt
        assert self.cell.value == dt
        assert self.cell.internal_value == 40372.27616898148

    def test_date_interpretation(self):
        dt = date(2010, 7, 13)
        self.cell.value = dt
        assert self.cell.value == datetime(2010, 7, 13, 0, 0)
        assert self.cell.internal_value == 40372

    def test_number_format_style(self):
        self.cell.value = '12.6%'
        assert NumberFormat.FORMAT_PERCENTAGE == self.cell.style.number_format.format_code


@pytest.mark.parametrize("value, datatype",
                         [
                            (None, Cell.TYPE_NULL),
                            ('.0e000', Cell.TYPE_NUMERIC),
                            ('-0.e-0', Cell.TYPE_NUMERIC),
                            ('1E', Cell.TYPE_STRING)
                         ])
def test_data_type_check(value, datatype):
    ws = build_dummy_worksheet()
    cell = Cell(ws, 'A', 1)
    cell.bind_value(value)
    assert cell.data_type == datatype


def test_set_bad_type():
    ws = build_dummy_worksheet()
    cell = Cell(ws, 'A', 1)
    with pytest.raises(DataTypeException):
        cell.set_explicit_value(1, 'q')


def test_illegal_chacters():
    from openpyxl.exceptions import IllegalCharacterError
    ws = build_dummy_worksheet()
    cell = Cell(ws, 'A', 1)

    # The bytes 0x00 through 0x1F inclusive must be manually escaped in values.

    illegal_chrs = range(9) + range(11, 13) + range(14, 32)
    for i in illegal_chrs:
        with pytest.raises(IllegalCharacterError):
            cell.value = chr(i)

        with pytest.raises(IllegalCharacterError):
            cell.value = "A {0} B".format(chr(i))

    cell.value = chr(33)
    cell.value = chr(9)  # Tab
    cell.value = chr(10)  # Newline
    cell.value = chr(13)  # Carriage return
    cell.value = " Leading and trailing spaces are legal "


values = (
    ('30:33.865633336', [('', '', '', '30', '33', '865633')]),
    ('03:40:16', [('03', '40', '16', '', '', '')]),
    ('03:40', [('03', '40', '',  '', '', '')]),
    ('55:72:12', []),
    )
@pytest.mark.parametrize("value, expected",
                             values)
def test_time_regex(value, expected):
    from openpyxl.cell.cell import TIME_REGEX
    m = TIME_REGEX.findall(value)
    assert m == expected


values = (
    ('03:40:16', time(3, 40, 16)),
    ('03:40', time(3, 40)),
    ('30:33.865633336', time(0, 30, 33, 865633))
)
@pytest.mark.parametrize("value, expected",
                         values)
def test_time(value, expected):
    wb = Workbook()
    ws = Worksheet(wb)
    cell = Cell(ws, 'A', 1)
    cell.value = value
    assert cell.value == expected
    assert cell.TYPE_NUMERIC == cell.data_type


def test_timedelta():

    wb = Workbook()
    ws = Worksheet(wb)
    cell = Cell(ws, 'A', 1)
    cell.value = timedelta(days=1, hours=3)
    assert cell.value == 1.125
    assert cell.TYPE_NUMERIC == cell.data_type


def test_date_format_on_non_date():
    wb = Workbook()
    ws = Worksheet(wb)
    cell = Cell(ws, 'A', 1)
    cell.value = datetime.now()
    cell.value = 'testme'
    assert 'testme' == cell.value

def test_set_get_date():
    today = datetime(2010, 1, 18, 14, 15, 20, 1600)
    wb = Workbook()
    ws = Worksheet(wb)
    cell = Cell(ws, 'A', 1)
    cell.value = today
    assert today == cell.value

def test_repr():
    wb = Workbook()
    ws = Worksheet(wb)
    cell = Cell(ws, 'A', 1)
    assert repr(cell), '<Cell Sheet1.A1>' == 'Got bad repr: %s' % repr(cell)

def test_is_date():
    wb = Workbook()
    ws = Worksheet(wb)
    cell = Cell(ws, 'A', 1)
    cell.value = datetime.now()
    assert cell.is_date() == True
    cell.value = 'testme'
    assert 'testme' == cell.value
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

def test_cell_offset():
    wb = Workbook()
    ws = Worksheet(wb)
    assert ws['B15'].offset(2, 1).coordinate == 'C17'
