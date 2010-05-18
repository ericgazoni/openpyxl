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

    matches = re.match(pattern = '[$]?([A-Z]+)[$]?(\d+)', string = coord_string)

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

    def __init__(self, worksheet, column, row, value = None, data_type = None):

        self.column = column.upper()
        self.row = row

        self.value = value

        self.parent = worksheet

        self.data_type = data_type

    def get_coordinate(self):

        return '%s%s' % (self.column, self.row)

