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

"""Manage individual cells in a spreadsheet.

The Cell class is required to know its value and type, display options,
and any other features of an Excel cell.  Utilities for referencing
cells using Excel's 'A1' column/row nomenclature are also provided.

"""

__docformat__ = "restructuredtext en"

# Python stdlib imports
import datetime
import re
import warnings

from openpyxl.compat import lru_cache, xrange
from openpyxl.units import (
    NUMERIC_TYPES,
    DEFAULT_ROW_HEIGHT,
    DEFAULT_COLUMN_WIDTH
)
from openpyxl.compat import unicode, basestring
from openpyxl.date_time import (
    to_excel,
    time_to_days,
    timedelta_to_days,
    from_excel
    )
from openpyxl.exceptions import (
    CellCoordinatesException,
    DataTypeException,
    IllegalCharacterError
)
from openpyxl.units import points_to_pixels
from openpyxl.styles import NumberFormat, is_date_format


# package imports

# constants
COORD_RE = re.compile('^[$]?([A-Z]+)[$]?(\d+)$')
ABSOLUTE_RE = re.compile('^[$]?([A-Z]+)[$]?(\d+)(:[$]?([A-Z]+)[$]?(\d+))?$')
ILLEGAL_CHARACTERS_RE = re.compile('|'.join(chr(x) for x in range(33)))



def coordinate_from_string(coord_string):
    """Convert a coordinate string like 'B12' to a tuple ('B', 12)"""
    match = COORD_RE.match(coord_string.upper())
    if not match:
        msg = 'Invalid cell coordinates (%s)' % coord_string
        raise CellCoordinatesException(msg)
    column, row = match.groups()
    row = int(row)
    if not row:
        msg = "There is no row 0 (%s)" % coord_string
        raise CellCoordinatesException(msg)
    return (column, row)


def absolute_coordinate(coord_string):
    """Convert a coordinate to an absolute coordinate string (B12 -> $B$12)"""
    m = ABSOLUTE_RE.match(coord_string)
    if m:
        parts = m.groups()
        if all(parts[-2:]):
            return '$%s$%s:$%s$%s' % (parts[0], parts[1], parts[3], parts[4])
        else:
            return '$%s$%s' % (parts[0], parts[1])
    else:
        return coord_string

@lru_cache(maxsize=1000)
def get_column_letter(col_idx):
    """Convert a column number into a column letter (3 -> 'C')

    Right shift the column col_idx by 26 to find column letters in reverse
    order.  These numbers are 1-based, and can be converted to ASCII
    ordinals by adding 64.

    """
    # these indicies corrospond to A -> ZZZ and include all allowed
    # columns
    if not 1 <= col_idx <= 18278:
        raise ValueError("Invalid column index {0}".format(col_idx))
    letters = []
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx, 26)
        # check for exact division and borrow if needed
        if remainder == 0:
            remainder = 26
            col_idx -= 1
        letters.append(chr(remainder+64))
    return ''.join(reversed(letters))


_COL_STRING_CACHE = dict((get_column_letter(i), i) for i in xrange(1, 18279))
def column_index_from_string(str_col, cache=_COL_STRING_CACHE):
    # we use a function argument to get indexed name lookup
    col = cache.get(str_col.upper())
    if col is None:
        raise ValueError("{0} is not a valid column name".format(str_col))
    return col
del _COL_STRING_CACHE


PERCENT_REGEX = re.compile(r'^\-?(?P<number>[0-9]*\.?[0-9]*\s?)\%$')
TIME_REGEX = re.compile(r"""
^(?: # HH:MM and HH:MM:SS
(?P<hour>[0-1]{0,1}[0-9]{2}):
(?P<minute>[0-5][0-9]):?
(?P<second>[0-5][0-9])?$)
|
^(?: # MM:SS.
([0-5][0-9]):
([0-5][0-9])?\.
(?P<microsecond>\d{1,6}))
""", re.VERBOSE)
NUMBER_REGEX = re.compile(r'^-?([\d]|[\d]+\.[\d]*|\.[\d]+|[1-9][\d]+\.?[\d]*)((E|e)-?[\d]+)?$')

class Cell(object):
    """Describes cell associated properties.

    Properties of interest include style, type, value, and address.

    """
    __slots__ = ('column',
                 'row',
                 '_value',
                 '_data_type',
                 'parent',
                 'xf_index',
                 '_hyperlink_rel',
                 'merged',
                 '_comment')

    ERROR_CODES = {'#NULL!': 0,
                   '#DIV/0!': 1,
                   '#VALUE!': 2,
                   '#REF!': 3,
                   '#NAME?': 4,
                   '#NUM!': 5,
                   '#N/A': 6}

    TYPE_STRING = 's'
    TYPE_FORMULA = 'f'
    TYPE_NUMERIC = 'n'
    TYPE_BOOL = 'b'
    TYPE_NULL = 's'
    TYPE_INLINE = 'inlineStr'
    TYPE_ERROR = 'e'
    TYPE_FORMULA_CACHE_STRING = 'str'

    VALID_TYPES = [TYPE_STRING, TYPE_FORMULA, TYPE_NUMERIC, TYPE_BOOL,
                   TYPE_NULL, TYPE_INLINE, TYPE_ERROR, TYPE_FORMULA_CACHE_STRING]

    def __init__(self, worksheet, column, row, value=None):
        self.column = column.upper()
        self.row = row
        # _value is the stored value, while value is the displayed value
        self._value = None
        self._hyperlink_rel = None
        self.data_type = self.TYPE_NULL
        if value:
            self.value = value
        self.parent = worksheet
        self.xf_index = 0
        self.merged = False
        self._comment = None

    @property
    def encoding(self):
        return self.parent.encoding

    @property
    def base_date(self):
        return self.parent.parent.excel_base_date

    def __repr__(self):
        return unicode("<Cell %s.%s>") % (self.parent.title, self.coordinate)

    def check_string(self, value):
        """Check string coding, length, and line break character"""
        # convert to unicode string
        if not isinstance(value, unicode):
            value = unicode(value, self.encoding)
        value = unicode(value)
        # string must never be longer than 32,767 characters
        # truncate if necessary
        value = value[:32767]
        if ILLEGAL_CHARACTERS_RE.match(value):
            raise IllegalCharacterError
        # we require that newline is represented as "\n" in core,
        # not as "\r\n" or "\r"
        value = value.replace('\r\n', '\n')
        return value

    def check_numeric(self, value):
        """Cast value to int or float if necessary"""
        if not isinstance(value, NUMERIC_TYPES):
            try:
                value = int(value)
            except ValueError:
                value = float(value)
        return value

    def check_error(self, value):
        """Tries to convert Error" else N/A"""
        try:
            return unicode(value)
        except:
            return unicode('#N/A')

    def set_explicit_value(self, value=None, data_type=TYPE_STRING):
        """Coerce values according to their explicit type"""
        type_coercion_map = {
            self.TYPE_INLINE: self.check_string,
            self.TYPE_STRING: self.check_string,
            self.TYPE_FORMULA: self.check_string,
            self.TYPE_NUMERIC: self.check_numeric,
            self.TYPE_BOOL: bool,
            self.TYPE_ERROR: self.check_error}
        try:
            self._value = type_coercion_map[data_type](value)
        except KeyError:
            if data_type not in self.VALID_TYPES:
                msg = 'Invalid data type: %s' % data_type
                raise DataTypeException(msg)
        self.data_type = data_type

    # preserve old method name
    set_value_explicit = set_explicit_value

    def data_type_for_value(self, value):
        """Given a value, infer the correct data type"""
        if value is None:
            data_type = self.TYPE_NULL
        elif value is True or value is False:
            data_type = self.TYPE_BOOL
        elif isinstance(value, NUMERIC_TYPES):
            data_type = self.TYPE_NUMERIC
        elif isinstance(value, (datetime.datetime, datetime.date, datetime.time, datetime.timedelta)):
            data_type = self.TYPE_NUMERIC
        elif not value:
            data_type = self.TYPE_STRING
        elif isinstance(value, basestring) and value[0] == '=':
            data_type = self.TYPE_FORMULA
        elif isinstance(value, unicode) and NUMBER_REGEX.match(value):
            data_type = self.TYPE_NUMERIC
        elif not isinstance(value, unicode) and NUMBER_REGEX.match(str(value)):
            data_type = self.TYPE_NUMERIC
        elif isinstance(value, basestring) and value.strip() in self.ERROR_CODES:
            data_type = self.TYPE_ERROR
        elif isinstance(value, list):
            data_type = self.TYPE_ERROR
        else:
            data_type = self.TYPE_STRING
        return data_type

    def bind_value(self, value):
        """Given a value, infer type and display options."""
        self.data_type = self.data_type_for_value(value)
        if value is None:
            self.set_explicit_value('', self.TYPE_NULL)
            return True
        elif self.data_type == self.TYPE_STRING:
            # percentage detection
            if self._bind_percentage(value):
                return
            # time detection
            if self._bind_time(value):
                return
        if self.data_type == self.TYPE_NUMERIC:
            if self._bind_datetime(value):
                return
        self.set_explicit_value(value, self.data_type)

    def _bind_percentage(self, value):
        if not isinstance(value, unicode):
            value = str(value)
        match = PERCENT_REGEX.match(value)
        if match:
            value = float(match.group('number')) / 100
            self.set_explicit_value(value, self.TYPE_NUMERIC)
            self.number_format = NumberFormat.FORMAT_PERCENTAGE
            return True

    def _bind_time(self, value):
        if not isinstance(value, unicode):
            value = str(value)
        match = TIME_REGEX.match(value)
        if match:
            if match.group("microsecond") is not None:
                value = value[:12]
                pattern = "%M:%S.%f"
                fmt = NumberFormat.FORMAT_DATE_TIME5
            elif match.group('second') is None:
                fmt = NumberFormat.FORMAT_DATE_TIME3
                pattern = "%H:%M"
            else:
                pattern = "%H:%M:%S"
                fmt = NumberFormat.FORMAT_DATE_TIME6
            value = datetime.datetime.strptime(value, pattern)
            value = time_to_days(value)
            self.set_explicit_value(value, self.TYPE_NUMERIC)
            self.number_format = fmt
            return True

    def _bind_datetime(self, value):
        if isinstance(value, datetime.date):
            value = to_excel(value, self.base_date)
            self.number_format = NumberFormat.FORMAT_DATE_YYYYMMDD2
        elif isinstance(value, datetime.time):
            value = time_to_days(value)
            self.number_format = NumberFormat.FORMAT_DATE_TIME6
        elif isinstance(value, datetime.timedelta):
            value = timedelta_to_days(value)
            self.number_format = NumberFormat.FORMAT_DATE_TIMEDELTA
        self.set_explicit_value(value, self.TYPE_NUMERIC)
        return True


    @property
    def value(self):
        """Get or set the value held in the cell.
            ':rtype: depends on the value (string, float, int or '
            ':class:`datetime.datetime`)'"""
        value = self._value
        if self.is_date():
            value = from_excel(value, self.base_date)
        return value

    @value.setter
    def value(self, value):
        """Set the value and infer type and display options."""
        self.bind_value(value)

    @property
    def internal_value(self):
        """Always returns the value for excel."""
        return self._value

    @property
    def hyperlink(self):
        """Return the hyperlink target or an empty string"""
        return self._hyperlink_rel is not None and \
                self._hyperlink_rel.target or ''

    @hyperlink.setter
    def hyperlink(self, val):
        """Set value and display for hyperlinks in a cell.
        Automatically setsthe `value` of the cell with link text,
        but you can modify it afterwards by setting the `value`
        property, and the hyperlink will remain.\n\n' ':rtype: string"""
        if self._hyperlink_rel is None:
            self._hyperlink_rel = self.parent.create_relationship("hyperlink")
        self._hyperlink_rel.target = val
        self._hyperlink_rel.target_mode = "External"
        if self._value is None:
            self.value = val

    @property
    def hyperlink_rel_id(self):
        """Return the id pointed to by the hyperlink, or None"""
        return self._hyperlink_rel is not None and \
                self._hyperlink_rel.id or None

    @property
    def number_format(self):
        style = self.parent.get_style(self.coordinate, read_only=True)
        return style.number_format.format_code

    @number_format.setter
    def number_format(self, format_code):
        """Set a new formatting code for numeric values"""
        self.style.number_format.format_code = format_code

    def _set_number_format(self, format_code):
        """Set a new formatting code for numeric values"""
        warnings.warn("cell._set_number_format(value) is deprecated use cell.number_format = value instead")
        self.number_format = format_code

    def is_date(self):
        """Whether the value is formatted as a date

        :rtype: bool
        """
        return (self.has_style
                and is_date_format(self.number_format)
                and self.data_type == self.TYPE_NUMERIC)

    @property
    def has_style(self):
        """Check if the parent worksheet has a style for this cell"""
        return self.coordinate in self.parent._styles  # pylint: disable=W0212

    @property
    def style(self):
        """Returns the :class:`openpyxl.style.Style` object for this cell"""
        return self.parent.get_style(self.coordinate)

    @property
    def data_type(self):
        """Return the data type represented by this cell"""
        return self._data_type

    @data_type.setter
    def data_type(self, value):
        self._data_type = value

    def get_coordinate(self):
        warnings.warn("cell.get_coordinate() is deprecated use cell.coordinate instead")
        return self.coordinate

    @property
    def coordinate(self):
        """Return the coordinate string for this cell (e.g. 'B12')

        :rtype: string
        """
        return '{1}{0}'.format(self.row, self.column)

    @property
    def address(self):
        warnings.warn("cell.address is deprecated, use cell.coordinate instead")
        return self.coordinate

    def offset(self, row=0, column=0):
        """Returns a cell location relative to this cell.

        :param row: number of rows to offset
        :type row: int

        :param column: number of columns to offset
        :type column: int

        :rtype: :class:`openpyxl.cell.Cell`
        """
        offset_column = get_column_letter(column_index_from_string(
            self.column) + column)
        offset_row = self.row + row
        return self.parent.cell('%s%s' % (offset_column, offset_row))

    @property
    def anchor(self):
        """ returns the expected position of a cell in pixels from the top-left
            of the sheet. For example, A1 anchor should be (0,0).

            :rtype: tuple(int, int)
        """
        left_columns = (column_index_from_string(self.column) - 1)
        column_dimensions = self.parent.column_dimensions
        left_anchor = 0
        default_width = points_to_pixels(DEFAULT_COLUMN_WIDTH)

        for col_idx in range(left_columns):
            letter = get_column_letter(col_idx + 1)
            if letter in column_dimensions:
                cdw = column_dimensions.get(letter).width
                if cdw > 0:
                    left_anchor += points_to_pixels(cdw)
                    continue
            left_anchor += default_width

        row_dimensions = self.parent.row_dimensions
        top_anchor = 0
        top_rows = (self.row - 1)
        default_height = points_to_pixels(DEFAULT_ROW_HEIGHT)
        for row_idx in range(1, top_rows + 1):
            if row_idx in row_dimensions:
                rdh = row_dimensions[row_idx].height
                if rdh > 0:
                    top_anchor += points_to_pixels(rdh)
                    continue
            top_anchor += default_height

        return (left_anchor, top_anchor)

    @property
    def comment(self):
        """ Returns the comment associated with this cell

            :rtype: :class:`openpyxl.comments.Comment`
        """
        return self._comment

    @comment.setter
    def comment(self, value):
        if value is not None and value._parent is not None and value is not self.comment:
            raise AttributeError(
                "Comment already assigned to %s in worksheet %s. Cannot assign a comment to more than one cell" %
                (value._parent.coordinate, value._parent.parent.title)
                )

        # Ensure the number of comments for the parent worksheet is up-to-date
        if value is None and self._comment is not None:
            self.parent._comment_count -= 1
        if value is not None and self._comment is None:
            self.parent._comment_count += 1

        # orphan the old comment
        if self._comment is not None:
            self._comment._parent = None

        self._comment = value
        if value is not None:
            self._comment._parent = self
