# file openpyxl/writer/paged_worksheet.py

# Copyright (c) 2010 openpyxl
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
# @author: Eric Gazoni

""" Paged worksheet 
*Still very raw*
"""

from hashlib import md5
from uuid import uuid4
from collections import deque
from itertools import chain
from openpyxl.cell import column_index_from_string, coordinate_from_string, Cell
from openpyxl.worksheet import Worksheet, RowDimension, ColumnDimension
import cPickle as Pickle

PAGE_HEIGHT = 1000
PAGE_WIDTH = 10
CACHE_SIZE = 10

class PagedWorksheet(Worksheet):

    def __init__(self, parent_workbook):

        Worksheet.__init__(self, parent_workbook)
        self._sheet_codename = str(uuid4())
        self._all_page_files = set()
        self._page_cache_order = deque()
        self._cache = {}

    def __del__(self):

        self._commit_pages()

    def get_cell_collection(self):
        pages = (self._get_page(page_filename)
                    for page_filename in self._all_page_files)

        paged_cells = chain.from_iterable(x.values() for x in pages)
        cached_cells = chain.from_iterable(x.values() for x in self._cache.values())

        res = chain(paged_cells, cached_cells)

        return res

    def _commit_pages(self):

        while self._page_cache_order:
            page_filename = self._page_cache_order.popleft()
            self._persist_page(page_filename)

    def _get_page_from_coordinate(self, column, row):

        return '%s-%d-%d.page' % (self._sheet_codename ,
                                  (column_index_from_string(column) / PAGE_WIDTH),
                                  (row / PAGE_HEIGHT))

    def _get_page(self, page_filename):

        if not page_filename in self._cache:

            print "<<< loading ...", page_filename

            self._page_cache_order.append(page_filename)

            if len(self._page_cache_order) > CACHE_SIZE:
                prune = self._page_cache_order.popleft()
                self._persist_page(page_filename = prune)

            if page_filename in self._all_page_files:
                pagefile = open(page_filename, 'rb')
                res = Pickle.load(pagefile)
                self._cache[page_filename] = res
                pagefile.close()
            else:
                res = self._cache[page_filename] = {}

        return self._cache[page_filename]

    def _persist_page(self, page_filename):

        print ">>> writing ...", page_filename
        self._all_page_files.add(page_filename)
        pagefile = open(page_filename, 'wb')
        Pickle.dump(self._cache[page_filename], pagefile, -1)
        del self._cache[page_filename]
        pagefile.close()

    def _get_cell(self, coordinate):

        column, row = coordinate_from_string(coordinate)
        page_filename = self._get_page_from_coordinate(column, row)

        self._cells = self._get_page(page_filename)

        if not coordinate in self._cells:
            new_cell = Cell(self, column, row)
            self._cells[coordinate] = new_cell
            if column not in self.column_dimensions:
                self.column_dimensions[column] = ColumnDimension(column)
            if row not in self.row_dimensions:
                self.row_dimensions[row] = RowDimension(row)

        return self._cells[coordinate]

