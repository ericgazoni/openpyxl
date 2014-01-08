# file openpyxl/reader/iter_worksheet.py

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
import tempfile
import zlib
import zipfile
import struct
from collections import namedtuple

# compatibility
from openpyxl.shared.compat import iterparse, xrange

# package
from openpyxl.worksheet import Worksheet
from openpyxl.cell import (
    coordinate_from_string,
    get_column_letter,
    Cell,
    column_index_from_string
)
from openpyxl.reader.style import read_style_table
from openpyxl.styles import is_date_format
from openpyxl.shared.date_time import SharedDate
from openpyxl.reader.worksheet import read_dimension
from openpyxl.shared.compat import unicode
from openpyxl.shared.ooxml import (
    PACKAGE_WORKSHEETS,
    ARC_STYLE,
    SHEET_MAIN_NS
)

TYPE_NULL = Cell.TYPE_NULL
MISSING_VALUE = None

RE_COORDINATE = re.compile('^([A-Z]+)([0-9]+)$')

SHARED_DATE = SharedDate()

_COL_CONVERSION_CACHE = dict((get_column_letter(i), i) for i in xrange(1, 18279))
def column_index_from_string(str_col, _col_conversion_cache=_COL_CONVERSION_CACHE):
    # we use a function argument to get indexed name lookup
    return _col_conversion_cache[str_col]
del _COL_CONVERSION_CACHE

RAW_ATTRIBUTES = ['row', 'column', 'coordinate', 'internal_value',
                  'data_type', 'style_id', 'number_format']


BaseRawCell = namedtuple('RawCell', RAW_ATTRIBUTES)


class RawCell(BaseRawCell):
    """Optimized version of the :class:`openpyxl.cell.Cell`, using named tuples.

    Useful attributes are:

    * row
    * column
    * coordinate
    * internal_value

    You can also access if needed:

    * data_type
    * number_format

    """

    @property
    def is_date(self):
        return is_date_format(self.number_format)

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


def get_missing_cells(row, columns):

    return dict([(column, RawCell(row, column, '%s%s' % (column, row),
                                  MISSING_VALUE, TYPE_NULL, None, None)) for column in columns])


#------------------------------------------------------------------------------

class IterableWorksheet(Worksheet):

    def __init__(self, parent_workbook, title, workbook_name,
            sheet_codename, xml_source, string_table):

        Worksheet.__init__(self, parent_workbook, title)
        self.archive = zipfile.ZipFile(workbook_name, 'r')
        self._workbook_name = workbook_name
        self._sheet_codename = sheet_codename
        self._string_table = string_table

        min_col, min_row, max_col, max_row = read_dimension(xml_source=self.xml_source)
        self.min_col = min_col
        self.min_row = min_row
        self.max_row = max_row
        self.max_col = max_col

        self._shared_date = SharedDate(base_date=parent_workbook.excel_base_date)

    @property
    def xml_source(self):
        worksheet_path = '%s/%s' % (PACKAGE_WORKSHEETS, self._sheet_codename)
        return self.archive.open(worksheet_path)

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
        expected_columns = [get_column_letter(ci) for ci in xrange(min_col, max_col)]
        current_row = min_row

        style_properties = read_style_table(self.archive.read(ARC_STYLE))
        style_table = style_properties.pop('table')

        for row, cells in groupby(self.get_cells(min_row, min_col,
                                                 max_row, max_col),
                                  operator.attrgetter('row')):
            full_row = []
            if current_row < row:

                for gap_row in xrange(current_row, row):
                    dummy_cells = get_missing_cells(gap_row, expected_columns)
                    yield tuple([dummy_cells[column] for column in expected_columns])
                    current_row = row

            temp_cells = list(cells)
            retrieved_columns = dict([(c.column, c) for c in temp_cells])
            missing_columns = list(set(expected_columns) - set(retrieved_columns.keys()))
            replacement_columns = get_missing_cells(row, missing_columns)

            for column in expected_columns:
                if column in retrieved_columns:
                    cell = retrieved_columns[column]
                    if cell.style_id is not None:
                        style = style_table[int(cell.style_id)]
                        cell = cell._replace(number_format=style.number_format.format_code) #pylint: disable-msg=W0212
                    if cell.internal_value is not None:
                        if cell.data_type in Cell.TYPE_STRING:
                            cell = cell._replace(internal_value=unicode(self._string_table[int(cell.internal_value)])) #pylint: disable-msg=W0212
                        elif cell.data_type == Cell.TYPE_BOOL:
                            cell = cell._replace(internal_value=cell.internal_value == '1')
                        elif cell.is_date:
                            cell = cell._replace(internal_value=self._shared_date.from_julian(float(cell.internal_value)))
                        elif cell.data_type == Cell.TYPE_NUMERIC:
                            cell = cell._replace(internal_value=float(cell.internal_value))
                        elif cell.data_type in(Cell.TYPE_INLINE, Cell.TYPE_FORMULA_CACHE_STRING):
                            cell = cell._replace(internal_value=unicode(cell.internal_value))
                    full_row.append(cell)
                else:
                    full_row.append(replacement_columns[column])
            current_row = row + 1
            yield tuple(full_row)


    def get_cells(self, min_row, min_col, max_row, max_col):
        p = iterparse(self.xml_source)

        for _event, element in p:

            if element.tag == '{%s}c' % SHEET_MAIN_NS:
                coord = element.get('r')
                column_str, row = RE_COORDINATE.match(coord).groups()

                row = int(row)
                column = column_index_from_string(column_str)

                if min_col <= column <= max_col and min_row <= row <= max_row:
                    data_type = element.get('t', 'n')
                    style_id = element.get('s')
                    formula = element.findtext('{%s}f' % SHEET_MAIN_NS)
                    value = element.findtext('{%s}v' % SHEET_MAIN_NS)
                    if formula is not None and not self.parent.data_only:
                        data_type = Cell.TYPE_FORMULA
                        value = "=" + formula
                    yield RawCell(row, column_str, coord, value, data_type, style_id, None)
            # sub-elements of cells should be skipped
            if (element.tag == '{%s}v' % SHEET_MAIN_NS
                or element.tag == '{%s}f' % SHEET_MAIN_NS):
                continue
            element.clear()


    def cell(self, *args, **kwargs):
        raise NotImplementedError("use 'iter_rows()' instead")

    def range(self, *args, **kwargs):
        raise NotImplementedError("use 'iter_rows()' instead")

    def calculate_dimension(self):
        return '%s%s:%s%s' % (self.min_col, self.min_row, self.max_col, self.max_row)

    def get_highest_column(self):
        return column_index_from_string(self.max_col)

    def get_highest_row(self):
        return self.max_row
