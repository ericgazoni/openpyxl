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

@license: http://www.opensource.org/licenses/mit-license.php
@author: Eric Gazoni
'''

import re

def coordinate_from_string(coord_string):

    matches = re.match(pattern = '[$]?([A-Z]+)[$]?(\d+)', string = coord_string.upper())

    if not matches:
        raise Exception('invalid cell coordinates')
    else:
        column, row = matches.groups()
        return (column, int(row))

def column_index_from_string(column):

    column = column.upper()

    clen = len(column)

    if clen == 1:
        return ord(column[0]) - 64
    elif clen == 2:
        return ((1 + (ord(column[0]) - 65)) * 26) + (ord(column[1]) - 64)
    elif clen == 3:
        return ((1 + (ord(column[0]) - 65)) * 676) + ((1 + (ord(column[1]) - 65)) * 26) + (ord(column[2]) - 64);
    elif clen > 3:
        raise Exception('Column string index can not be longer than 3 characters')
    else:
        raise Exception('Column string index can not be empty')

def get_column_letter(col_idx):

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

    TYPE_STRING = 's'
    TYPE_FORMULA = 'f'
    TYPE_NUMERIC = 'n'
    TYPE_BOOL = 'b'
    TYPE_NULL = 's'
    TYPE_INLINE = 'inlineStr'
    TYPE_ERROR = 'e'

    def __init__(self, worksheet, column, row, value = None):

        self.column = column.upper()
        self.row = row

        self._value = None
        self._data_type = self.TYPE_NULL

        if value:
            self.value = value

        self.parent = worksheet

    def _get_value(self):

        return self._value

    def _set_value(self, value):

        self._value = value
        if value is None:
            self._data_type = self.TYPE_NULL
        elif not value:
            self._data_type = self.TYPE_STRING
        elif isinstance(value, basestring) and value[0] == '=':
            self._data_type = self.TYPE_FORMULA
        elif isinstance(value, (int, float)):
            self._data_type = self.TYPE_NUMERIC
        elif re.match(pattern = '^\-?([0-9]+\\.?[0-9]*|[0-9]*\\.?[0-9]+)$', string = value):
            self._value = float(value)
            self._data_type = self.TYPE_NUMERIC
        else:
            self._data_type = self.TYPE_STRING

    value = property(_get_value, _set_value)

    @property
    def data_type(self):
        return self._data_type

    def get_coordinate(self):

        return '%s%s' % (self.column, self.row)

