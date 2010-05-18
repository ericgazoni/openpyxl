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

from openpyxl.cell import Cell, coordinate_from_string
from openpyxl.shared.password_hasher import hash_password

class PageSetup(object): pass
class PageMargins(object): pass
class HeaderFooter(object): pass
class SheetView(object): pass
class RowDimension(object): pass
class ColumnDimension(object): pass

class SheetProtection(object):

    def __init__(self):

        self.sheet = False
        self.objects = False
        self.scenarios = False
        self.format_cells = False
        self.format_columns = False
        self.format_rows = False
        self.insert_columns = False
        self.insert_rows = False
        self.insert_hyperlinks = False
        self.delete_columns = False
        self.delete_rows = False
        self.select_locked_cells = False
        self.sort = False
        self.auto_filter = False
        self.pivot_tables = False
        self.select_unlocked_cells = False
        self._password = ''

    def set_password(self, value = '', already_hashed = False):

        if not already_hashed:
            value = hash_password(value)

        self._password = value

    def _set_raw_password(self, value):

        self.set_password(value, already_hashed = False)

    def _get_raw_password(self):

        return self._password

    password = property(_get_raw_password, _set_raw_password,
                        'get/set the password (if already hashed, use set_password() instead)')


class Worksheet(object):

    BREAK_NONE = 0
    BREAK_ROW = 1
    BREAK_COLUMN = 2

    SHEETSTATE_VISIBLE = 'visible'
    SHEETSTATE_HIDDEN = 'hidden'
    SHEETSTATE_VERYHIDDEN = 'veryHidden'

    def __init__(self, parent_workbook, title = None):

        self._parent = parent_workbook

        if not title:
            self.title = 'Sheet%d' % (1 + len(self._parent.worksheets))
        else:
            self.title = title

        self._cells = {}

        self.sheet_state = self.SHEETSTATE_VISIBLE

        self.page_setup = PageSetup()

        self.page_margins = PageMargins()

        self.header_footer = HeaderFooter()

        self.sheet_view = SheetView()

        self.protection = SheetProtection()

        self.show_gridlines = True
        self.print_gridlines = False

        self.show_summary_below = True
        self.show_summary_right = True


    def cell(self, coordinate):

        if not coordinate in self._cells:
            column, row = coordinate_from_string(coord_string = coordinate)
            new_cell = Cell(worksheet = self, column = column, row = row)
            self._cells[coordinate] = new_cell

        return self._cells[coordinate]


