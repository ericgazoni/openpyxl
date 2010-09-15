'''
Copyright (c) 2010 openpyxl

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.

:License: http://www.opensource.org/licenses/mit-license.php
:Author: Eric Gazoni
'''

__docformat__ = "restructuredtext en"

from openpyxl.shared.date_time import SharedDate
from openpyxl.shared.exc import CellCoordinatesException, \
    ColumnStringIndexException, DataTypeException
from openpyxl.style import NumberFormat
import datetime
import re

def coordinate_from_string(coord_string):
    """Convert a coordinate string like 'B12' to a tuple ('B', 12)
    """

    matches = re.match(pattern = '[$]?([A-Z]+)[$]?(\d+)', string = coord_string.upper())

    if not matches:
        raise CellCoordinatesException('Invalid cell coordinates (%s)' % coord_string)
    else:
        column, row = matches.groups()
        return (column, int(row))

def absolute_coordinate(coord_string):
    """Convert coordinate string (e.g. 'B12') to absolute coordinate string (e.g. '$B$12')
    """

    return '$%s$%d' % coordinate_from_string(coord_string)

def column_index_from_string(column):
    """Convert a column letter (e.g. 'B') into a column number (e.g. 2)
    """

    column = column.upper()

    clen = len(column)

    if clen == 1:
        return ord(column[0]) - 64
    elif clen == 2:
        return ((1 + (ord(column[0]) - 65)) * 26) + (ord(column[1]) - 64)
    elif clen == 3:
        return ((1 + (ord(column[0]) - 65)) * 676) + ((1 + (ord(column[1]) - 65)) * 26) + (ord(column[2]) - 64);
    elif clen > 3:
        raise ColumnStringIndexException('Column string index can not be longer than 3 characters')
    else:
        raise ColumnStringIndexException('Column string index can not be empty')

def get_column_letter(col_idx):
    """Convert a column number (e.g. 3) into a column letter (e.g. 'C')
    """

    col_name = ""
    quotient = col_idx

    while col_idx > 26:
        quotient = col_idx / 26
        rest = col_idx % 26

        if rest > 0:
            col_name = chr(64 + rest) + col_name
        else:
            col_name = 'Z' + col_name # not beautiful, but it works fine ...
            quotient -= 1

        col_idx = quotient

    col_name = chr(64 + quotient) + col_name

    return col_name

class Cell(object):
    """Describes cell associated properties (style, type, value, address,...)
    """


    __slots__ = ('column',
                 'row',
                 '_value',
                 '_data_type',
                 'parent',
                 'xf_index',
                 '_hyperlink_rel')

    ERROR_CODES = {'#NULL!'  : 0,
                   '#DIV/0!' : 1,
                   '#VALUE!' : 2,
                   '#REF!'   : 3,
                   '#NAME?'  : 4,
                   '#NUM!'   : 5,
                   '#N/A'    : 6}

    TYPE_STRING = 's'
    TYPE_FORMULA = 'f'
    TYPE_NUMERIC = 'n'
    TYPE_BOOL = 'b'
    TYPE_NULL = 's'
    TYPE_INLINE = 'inlineStr'
    TYPE_ERROR = 'e'

    VALID_TYPES = [TYPE_STRING, TYPE_FORMULA, TYPE_NUMERIC, TYPE_BOOL,
                   TYPE_NULL, TYPE_INLINE, TYPE_ERROR]

    RE_PATTERNS = {'percentage' : re.compile('^\-?[0-9]*\.?[0-9]*\s?\%$'),
                   'time' : re.compile('^(\d|[0-1]\d|2[0-3]):[0-5]\d(:[0-5]\d)?$'),
                   'numeric' : re.compile('^\-?([0-9]+\\.?[0-9]*|[0-9]*\\.?[0-9]+)$'),
                   }

    def _set_hyperlink(self, val):
        if self._hyperlink_rel is None:
            self._hyperlink_rel = self.parent.create_relationship("hyperlink")
        self._hyperlink_rel.target = val
        self._hyperlink_rel.target_mode = "External"
        if self._value is None:
            self.value = val

    def _get_hyperlink(self):
        return self._hyperlink_rel is not None and self._hyperlink_rel.target or ""

    hyperlink = property(_get_hyperlink, _set_hyperlink)
    @property
    def hyperlink_rel_id(self):
        return self._hyperlink_rel is not None and self._hyperlink_rel.id or None


    def __repr__(self):

        return u"<Cell %s.%s>" % (self.parent.title, self.get_coordinate())

    def __init__(self, worksheet, column, row, value = None):

        self.column = column.upper()
        self.row = row

        self._value = None
        self._hyperlink_rel = None
        self._data_type = self.TYPE_NULL

        if value:
            self.value = value

        self.parent = worksheet

        self.xf_index = 0

    def _get_value(self):

        value = self._value

        if (self.has_style
            and self.style.number_format.is_date_format()
            and isinstance(value, (int, float))
            ):

            value = SharedDate().from_julian(value)

        return value

    def _set_value(self, value):

        self.bind_value(value)

    value = property(_get_value,
                     _set_value,
                     doc = """Get or set the value held in the cell
                     
                     :rtype: depends on the value (string, float, int or :class:`datetime.datetime`)
                     """)

    def bind_value(self, value):

        data_type = self._data_type = self.data_type_for_value(value)

        if data_type == self.TYPE_STRING:

            # percentage detection

            percentage_search = Cell.RE_PATTERNS['percentage'].match(value)

            if percentage_search and value.strip() != '%':

                value = float(value.replace('%', '')) / 100.0
                self.set_value_explicit(value = value,
                                        data_type = self.TYPE_NUMERIC)

                self._set_number_format(NumberFormat.FORMAT_PERCENTAGE)

                return True

            # time detection

            time_search = Cell.RE_PATTERNS['time'].match(value)

            if time_search:

                sep_count = value.count(':') #pylint: disable-msg=E1103

                if sep_count == 1:
                    h, m = map(int, value.split(':')) #pylint: disable-msg=E1103
                    s = 0
                elif sep_count == 2:
                    h, m, s = map(int, value.split(':')) #pylint: disable-msg=E1103

                days = (h / 24.0) + (m / 1440.0) + (s / 86400.0)

                self.set_value_explicit(value = days,
                                        data_type = self.TYPE_NUMERIC)

                self._set_number_format(NumberFormat.FORMAT_DATE_TIME3)

                return True

        if data_type == self.TYPE_NUMERIC:

            # date detection

            if isinstance(value, datetime.datetime):

                value = SharedDate().datetime_to_julian(date = value)

                self.set_value_explicit(value = value,
                                        data_type = self.TYPE_NUMERIC)

                self._set_number_format(NumberFormat.FORMAT_DATE_YYYYMMDD2)

                return True

        self.set_value_explicit(value, data_type)

    def _set_number_format(self, format_code):

        self.style.number_format.format_code = format_code

    @property
    def has_style(self):
        return self.get_coordinate() in self.parent._styles

    @property
    def style(self):
        """Returns the :class:`openpyxl.style.Style` object for this cell
        """
        return self.parent.get_style(self.get_coordinate())

    def data_type_for_value(self, value):

        if value is None:
            return self.TYPE_NULL
        elif value is True or value is False:
            return self.TYPE_BOOL
        elif isinstance(value, (int, float)):
            return self.TYPE_NUMERIC
        elif not value:
            return self.TYPE_STRING
        elif isinstance(value, datetime.datetime):
            return self.TYPE_NUMERIC
        elif isinstance(value, basestring) and value[0] == '=':
            return self.TYPE_FORMULA
        elif Cell.RE_PATTERNS['numeric'].match(value):
            return self.TYPE_NUMERIC
        elif value.strip() in self.ERROR_CODES:
            return self.TYPE_ERROR
        else:
            return self.TYPE_STRING


    def set_value_explicit(self, value = None, data_type = TYPE_STRING):

        if data_type == self.TYPE_INLINE:
            self._value = self.check_string(value)
        elif data_type == self.TYPE_FORMULA:
            self._value = unicode(value)
        elif data_type == self.TYPE_BOOL:
            self._value = bool(value)
        elif data_type == self.TYPE_STRING:
            self._value = self.check_string(value)
        elif data_type == self.TYPE_NUMERIC:

            if isinstance(value, (int, float)):
                self._value = value
            else:
                try:
                    self._value = int(value)
                except:
                    self._value = float(value)

        elif data_type not in self.VALID_TYPES:
            raise DataTypeException('Invalid data type: %s' % data_type)


        self._data_type = data_type

    def check_string(self, value):

        # convert to unicode string
        value = unicode(value)
        # string must never be longer than 32,767 characters, truncate if necessary
        value = value[:32767]
        # we require that newline is represented as "\n" in core, not as "\r\n" or "\r"
        value = value.replace('\r\n', '\n')

        return value

    @property
    def data_type(self):
        return self._data_type

    def get_coordinate(self):
        """Return the coordinate string for this cell (e.g. 'B12')
        
        :rtype: string
        """

        return '%s%s' % (self.column, self.row)

    def offset(self, row = 0, column = 0):
        """Returns the cell located at this cell's address, offsetted by `row` and `column`
        
        :param row: number of rows to offset
        :type row: int
        
        :param column: number of columns to offset
        :type column: int
        
        :rtype: :class:`openpyxl.cell.Cell`
        """

        offset_column = get_column_letter(column_index_from_string(column = self.column) + column)
        offset_row = self.row + row

        return self.parent.cell('%s%s' % (offset_column, offset_row))
