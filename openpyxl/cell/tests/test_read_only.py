# Copyright (c) 2010-2014 openpyxl

import datetime
import pytest

from openpyxl.cell.read_only import ReadOnlyCell


def test_ctor():
    cell = ReadOnlyCell(None, None, 10, 'n')
    assert cell.value == 10


def test_empty_cell():
    from openpyxl.cell.read_only import EMPTY_CELL
    assert EMPTY_CELL.value is None
    assert EMPTY_CELL.data_type == 's'


def test_base_date():
    ReadOnlyCell.set_base_date(2415018.5)
    cell = ReadOnlyCell(None, None, 10, 'n')
    assert cell.base_date == 2415018.5


def test_style_table():
    ReadOnlyCell.set_style_table({})
    cell = ReadOnlyCell(None, None, 10, 'n')
    assert cell.style_table == {}


def test_string_table():
    ReadOnlyCell.set_string_table({1:'Hello world'})
    cell = ReadOnlyCell(None, None, 1, 's')
    assert cell.string_table == {1:'Hello world'}
    assert cell.value == 'Hello world'


def test_coordinate():
    cell = ReadOnlyCell(1, "A", 10, None)
    assert cell.coordinate == "A1"
    cell = ReadOnlyCell(None, None, 1, None)
    with pytest.raises(AttributeError):
        cell.coordinate


@pytest.mark.parametrize("value, expected",
                         [
                         ('1', True),
                         ('0', False),
                         ])
def test_bool(value, expected):
    cell = ReadOnlyCell(None, None, value, 'b')
    assert cell.value is expected


def test_inline_String():
    cell = ReadOnlyCell(None, None, "Hello World!", 'inlineStr')
    assert cell.value == "Hello World!"


def test_numeric():
    cell = ReadOnlyCell(None, None, "24555", 'n')
    assert cell.value == 24555
    cell = ReadOnlyCell(None, None, None, 'n')
    assert cell.value is None


@pytest.fixture(scope="class")
def DummyCell():
    class DummyNumberFormat:

        format_code = 'd-mmm-yy'


    class DummyStyle(object):

        number_format = DummyNumberFormat()

    style_table = {1:DummyStyle()}
    cell = ReadOnlyCell(None, None, "23596", 'n', '1')
    cell.set_style_table(style_table)
    return cell


class TestDateTime:

    def test_number_format(self, DummyCell):
        assert DummyCell.number_format == 'd-mmm-yy'

    def test_is_date(self, DummyCell):
        assert DummyCell.is_date is True

    def test_conversion(self, DummyCell):
        assert DummyCell.value == datetime.datetime(1964, 8, 7, 0, 0, 0)

    def test_interal_value(self, DummyCell):
        assert DummyCell.internal_value == 23596


def test_read_only():
    cell = ReadOnlyCell(None, None, 1, None)
    with pytest.raises(AttributeError):
        cell.value = 10
    with pytest.raises(AttributeError):
        cell.style_id = 1


def test_equality():
    c1 = ReadOnlyCell(None, None, 10, None)
    c2 = ReadOnlyCell(None, None, 10, None)
    assert c1 is not c2
    assert c1 == c2
    c3 = ReadOnlyCell(None, None, 5, None)
    assert c3 != c1
