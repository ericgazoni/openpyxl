from __future__ import absolute_import
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

""" Iterators-based worksheet reader
*Still very raw*
"""
# stdlib
import operator
from itertools import groupby
import re

# compatibility
from openpyxl.compat import unicode, lru_cache, xrange
from openpyxl.xml.functions import iterparse

# package
from openpyxl import LXML
from openpyxl.worksheet import Worksheet
from openpyxl.cell import (
    coordinate_from_string,
    column_index_from_string,
    get_column_letter,
    Cell,
)
from openpyxl.styles import is_date_format
from openpyxl.date_time import from_excel
from openpyxl.reader.worksheet import read_dimension
from openpyxl.xml.constants import (
    PACKAGE_WORKSHEETS,
    SHEET_MAIN_NS
)


class RawCell(object):

    __slots__ = ('row', 'column', '_value', 'data_type', '_style_id')


    def __init__(self, row, column, value, data_type, style_id=None):
        self.row = row
        self.column = column
        self.data_type = data_type
        self.style_id = style_id
        self.value = value

    def __eq__(self, other):
        for a in self.__slots__:
            if getattr(self, a) != getattr(other, a):
                return False
        return True

    @classmethod
    def set_string_table(cls, string_table):
        cls.string_table = string_table

    @classmethod
    def set_style_table(cls, style_table):
        cls.style_table = style_table

    @classmethod
    def set_base_date(cls, base_date):
        cls.base_date = base_date

    @property
    def coordinate(self):
        return "{1}{0}".format(self.row, self.column)

    @property
    def is_date(self):
        return is_date_format(self.number_format)

    @property
    def number_format(self):
        if self.style_id is None:
            return
        style = self.style_table[self._style_id]
        return style.number_format.format_code

    @property
    def style_id(self):
        return self._style_id

    @style_id.setter
    def style_id(self, value):
        if value is not None:
            value = int(value)
        self._style_id = value

    @property
    def internal_value(self):
        return self._value

    @property
    def value(self):
        if self._value is None:
            return
        if self.data_type == Cell.TYPE_BOOL:
            return self._value == '1'
        elif self.is_date:
            return from_excel(self._value, self.base_date)
        elif self.data_type in(Cell.TYPE_INLINE, Cell.TYPE_FORMULA_CACHE_STRING):
            return unicode(self._value)
        elif self.data_type == Cell.TYPE_STRING:
            return unicode(self.string_table[int(self._value)])
        return self._value

    @value.setter
    def value(self, value):
        if value is None:
            self.data_type = Cell.TYPE_NULL
        elif self.data_type == Cell.TYPE_NUMERIC:
            value = float(value)
        self._value = value


def get_range_boundaries(range_string, row_offset=0, column_offset=1):

    if ':' in range_string:
        min_range, max_range = range_string.split(':')
        min_col, min_row = coordinate_from_string(min_range)
        max_col, max_row = coordinate_from_string(max_range)

        min_col = column_index_from_string(min_col)
        max_col = column_index_from_string(max_col) + 1

    else:
        min_col, min_row = coordinate_from_string(range_string)
        min_col = column_index_from_string(min_col)
        max_col = min_col + column_offset
        max_row = min_row + row_offset

    return (min_col, min_row, max_col, max_row)


def empty_cell(row, column):
    return RawCell(row, column, None, Cell.TYPE_NULL, None)

#------------------------------------------------------------------------------

CELL_TAG = '{%s}c' % SHEET_MAIN_NS
VALUE_TAG = '{%s}v' % SHEET_MAIN_NS
FORMULA_TAG = '{%s}f' % SHEET_MAIN_NS


class IterableWorksheet(Worksheet):

    def __init__(self, parent_workbook, title, worksheet_path,
                 xml_source, string_table, style_table):
        Worksheet.__init__(self, parent_workbook, title)
        self.worksheet_path = worksheet_path
        RawCell.set_string_table(string_table)
        RawCell.set_style_table(style_table)
        RawCell.set_base_date(parent_workbook.excel_base_date)

        min_col, min_row, max_col, max_row = read_dimension(xml_source=self.xml_source)
        self.min_col = min_col
        self.min_row = min_row
        self.max_row = max_row
        self.max_col = max_col

    @property
    def xml_source(self):
        return self.parent._archive.open(self.worksheet_path)

    @xml_source.setter
    def xml_source(self, value):
        """Base class is always supplied XML source, IteratableWorksheet obtains it on demand."""
        pass

    def __getitem__(self, key):
        if isinstance(key, slice):
            key = "{0}:{1}".format(key)
        return self.iter_rows(key)

    def iter_rows(self, range_string='', row_offset=0, column_offset=1):
        """ Returns a squared range based on the `range_string` parameter,
        using generators.

        :param range_string: range of cells (e.g. 'A1:C4')
        :type range_string: string

        :param row_offset: additional rows (e.g. 4)
        :type row: int

        :param column_offset: additonal columns (e.g. 3)
        :type column: int

        :rtype: generator

        """
        if range_string:
            min_col, min_row, max_col, max_row = get_range_boundaries(range_string, row_offset, column_offset)
        else:
            min_col = column_index_from_string(self.min_col)
            max_col = column_index_from_string(self.max_col) + 1
            min_row = self.min_row
            max_row = self.max_row + 6

        return self.get_squared_range(min_col, min_row, max_col, max_row)

    def get_squared_range(self, min_col, min_row, max_col, max_row):
        """
        The source worksheet file may have columns or rows missing.
        Missing cells will be created.
        """
        expected_columns = [get_column_letter(ci) for ci in xrange(min_col, max_col)]
        row_counter = min_row

        # get cells row by row
        for row, cells in groupby(self.get_cells(min_row, min_col,
                                                 max_row, max_col),
                                  operator.attrgetter('row')):
            full_row = []
            if row_counter < row:
                # Rows requested before those in the worksheet
                for gap_row in xrange(row_counter, row):
                    yield tuple(empty_cell(row, column) for column in expected_columns)
                    row_counter = row

            retrieved_columns = dict([(c.column, c) for c in cells])
            for column in expected_columns:
                if column in retrieved_columns:
                    cell = retrieved_columns[column]
                    full_row.append(cell)
                else:
                    # create missing cell
                    full_row.append(empty_cell(row, column))
            row_counter = row + 1
            yield tuple(full_row)


    def get_cells(self, min_row, min_col, max_row, max_col):
        p = iterparse(self.xml_source, events=('end',), tag=CELL_TAG, remove_blank_text=True)
        for _event, element in p:
            if element.tag == CELL_TAG:
                coord = element.get('r')
                column_str, row = coordinate_from_string(coord)
                column = column_index_from_string(column_str)

                if (min_col <= column <= max_col
                    and min_row <= row <= max_row):
                    data_type = element.get('t', 'n')
                    style_id = element.get('s')
                    formula = element.findtext(FORMULA_TAG)
                    value = element.findtext(VALUE_TAG)
                    if formula is not None and not self.parent.data_only:
                        data_type = Cell.TYPE_FORMULA
                        value = "=%s" % formula

                    yield RawCell(row, column_str, value, data_type,
                                  style_id)
            if not LXML and element.tag in (VALUE_TAG, FORMULA_TAG):
                # sub-elements of cells should be skipped
                continue
            element.clear()


    def cell(self, *args, **kwargs):
        # TODO return an individual cell

        raise NotImplementedError("use 'iter_rows()' instead")

    def range(self, *args, **kwargs):
        # TODO return a range of cells, basically get_squared_range with same interface as Worksheet
        raise NotImplementedError("use 'iter_rows()' instead")

    def rows(self):
        return self.iter_rows()

    def calculate_dimension(self):
        return '%s%s:%s%s' % (self.min_col, self.min_row, self.max_col, self.max_row)

    def get_highest_column(self):
        return column_index_from_string(self.max_col)

    def get_highest_row(self):
        return self.max_row
