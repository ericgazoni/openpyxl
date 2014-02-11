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

# compatibility
from openpyxl.compat import xrange
from openpyxl.xml.functions import iterparse

# package
from openpyxl.worksheet import Worksheet
from openpyxl.cell import (
    coordinate_from_string,
    column_index_from_string,
    get_column_letter,
    Cell
)
from openpyxl.cell.read_only import ReadOnlyCell, EMPTY_CELL
from openpyxl.xml.functions import safe_iterator
from openpyxl.xml.constants import SHEET_MAIN_NS


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

from openpyxl.reader.worksheet import _get_xml_iter


def read_dimension(source):
    min_row = min_col =  max_row = max_col = None
    DIMENSION_TAG = '{%s}dimension' % SHEET_MAIN_NS
    DATA_TAG = '{%s}sheetData' % SHEET_MAIN_NS
    it = iterparse(source, tag=[DIMENSION_TAG, DATA_TAG])
    for _event, element in it:
        if element.tag == DIMENSION_TAG:
            dim = element.get("ref")
            if ':' in dim:
                start, stop = dim.split(':')
            else:
                start = stop = dim
            min_col, min_row = coordinate_from_string(start)
            max_col, max_row = coordinate_from_string(stop)
            return min_col, min_row, max_col, max_row
        elif element.tag == DATA_TAG:
            # Dimensions missing
            break


ROW_TAG = '{%s}row' % SHEET_MAIN_NS
CELL_TAG = '{%s}c' % SHEET_MAIN_NS
VALUE_TAG = '{%s}v' % SHEET_MAIN_NS
FORMULA_TAG = '{%s}f' % SHEET_MAIN_NS
DIMENSION_TAG = '{%s}dimension' % SHEET_MAIN_NS

class IterableWorksheet(Worksheet):

    min_col = min_row = max_col = max_row = 1


    def __init__(self, parent_workbook, title, worksheet_path,
                 xml_source, string_table, style_table):
        Worksheet.__init__(self, parent_workbook, title)
        self.worksheet_path = worksheet_path
        ReadOnlyCell.set_string_table(string_table)
        ReadOnlyCell.set_style_table(style_table)
        ReadOnlyCell.set_base_date(parent_workbook.excel_base_date)
        dimensions = read_dimension(self.xml_source)
        if dimensions is not None:
            self.min_col, self.min_row, self.max_col, self.max_row = dimensions

    @property
    def xml_source(self):
        return self.parent._archive.open(self.worksheet_path)

    @xml_source.setter
    def xml_source(self, value):
        """Base class is always supplied XML source, IteratableWorksheet obtains it on demand."""
        pass

    @property
    def dimensions(self):
        return '%s%s:%s%s' % (self.min_col, self.min_row, self.max_col, self.max_row)

    def __getitem__(self, key):
        if isinstance(key, slice):
            key = "{0}:{1}".format(key.start, key.stop)
        if ":" in key:
            return self.iter_rows(key)
        return self.cell(key)

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
                    yield tuple(EMPTY_CELL for column in expected_columns)
                    row_counter = row

            retrieved_columns = dict([(c.column, c) for c in cells])
            for column in expected_columns:
                if column in retrieved_columns:
                    cell = retrieved_columns[column]
                    full_row.append(cell)
                else:
                    # create missing cell
                    full_row.append(EMPTY_CELL)
            row_counter = row + 1
            yield tuple(full_row)


    def get_cells(self, min_row, min_col, max_row, max_col):
        p = iterparse(self.xml_source, tag=[ROW_TAG], remove_blank_text=True)
        for _event, element in p:
            if element.tag == ROW_TAG:
                row = int(element.get("r"))
                if row > max_row:
                    break
                if min_row <= row:
                    for cell in safe_iterator(element, CELL_TAG):
                        coord = cell.get('r')
                        column_str, row = coordinate_from_string(coord)
                        column = column_index_from_string(column_str)
                        if column > max_col:
                            break
                        if min_col <= column:
                            data_type = cell.get('t', 'n')
                            style_id = cell.get('s')
                            formula = cell.findtext(FORMULA_TAG)
                            value = cell.findtext(VALUE_TAG)
                            if formula is not None and not self.parent.data_only:
                                data_type = Cell.TYPE_FORMULA
                                value = "=%s" % formula
                            yield ReadOnlyCell(row, column_str, value, data_type,
                                          style_id)
            if element.tag in (CELL_TAG, VALUE_TAG, FORMULA_TAG):
                # sub-elements of rows should be skipped
                continue
            element.clear()


    def _get_cell(self, coordinate):
        """.iter_rows always returns a generator of rows each of which
        contains a generator of cells. This can be empty in which case
        return None"""
        result = list(self.iter_rows(coordinate))
        if result:
            return result[0][0]


    def range(self, *args, **kwargs):
        # TODO return a range of cells, basically get_squared_range with same interface as Worksheet
        raise NotImplementedError("use 'iter_rows()' instead")

    def rows(self):
        return self.iter_rows()

    def calculate_dimension(self):
        return self.dimensions

    def get_highest_column(self):
        return column_index_from_string(self.max_col)

    def get_highest_row(self):
        return self.max_row
