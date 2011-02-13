# file openpyxl/writer/straight_worksheet.py

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

"""Write worksheets to xml representations in an optimized way"""

import datetime

from openpyxl.cell import column_index_from_string, get_column_letter
from openpyxl.shared.xmltools import XMLGenerator, get_document_content, \
        start_tag, end_tag, tag
from openpyxl.shared.date_time import SharedDate
from openpyxl.cell import Cell
from openpyxl.shared.ooxml import MAX_COLUMN, MAX_ROW
from tempfile import NamedTemporaryFile

STYLES = {'datetime' : {'type':Cell.TYPE_NUMERIC,
                        'style':'1'},
          'string':{'type':Cell.TYPE_INLINE,
                    'style':'0'},
          'numeric':{'type':Cell.TYPE_NUMERIC,
                     'style':'0'}
        }

BOUNDING_BOX_PLACEHOLDER = 'A1:%s%d' % (get_column_letter(MAX_COLUMN), MAX_ROW)

class DumpWorksheet(object):

    def __init__(self, parent_workbook):

        self._max_col = 0
        self._max_row = 0
        self._fileobj = NamedTemporaryFile(mode='w', prefix='openpyxl.')
        self.doc = XMLGenerator(self._fileobj, 'utf-8')

        self._shared_date = SharedDate()

    def write_header(self):

        doc = self.doc

        start_tag(doc, 'worksheet',
                {'xml:space': 'preserve',
                'xmlns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'})
        start_tag(doc, 'sheetPr')
        tag(doc, 'outlinePr',
                {'summaryBelow': '0', 
                'summaryRight': '0'})
        end_tag(doc, 'sheetPr')
        tag(doc, 'dimension', {'ref': '%s' % BOUNDING_BOX_PLACEHOLDER})
        start_tag(doc, 'sheetViews')
        start_tag(doc, 'sheetView', {'workbookViewId': '0'})
        tag(doc, 'selection', {'activeCell': 'A1',
                'sqref': 'A1'})
        end_tag(doc, 'sheetView')
        end_tag(doc, 'sheetViews')
        tag(doc, 'sheetFormatPr', {'defaultRowHeight': '15'})

    def close(self):

        self.doc.endDocument()
            
    def append(self, row):

        doc = self.doc

        self._max_row += 1
        span = len(row)
        self._max_col = max(self._max_col, span)

        row_idx = self._max_row

        attrs = {'r': '%d' % row_idx,
                 'spans': '1:%d' % span}

        start_tag(doc, 'row', attrs)

        for col_idx, cell in enumerate(row):

            coordinate = '%s%d' % (get_column_letter(col_idx+1), row_idx) 
            attributes = {'r': coordinate}


            if isinstance(cell, (int, float)):
                dtype = 'numeric'
            elif isinstance(cell, (datetime.datetime, datetime.date)):
                dtype = 'datetime'
                cell = self._shared_date.datetime_to_julian(cell)
            else:
                dtype = 'string'

            attributes['t'] = STYLES[dtype]['type']
            attributes['s'] = STYLES[dtype]['style']

            start_tag(doc, 'c', attributes)

            if cell is None:
                tag(doc, 'v', body='')
            else:
                tag(doc, 'v', body = '%s' % cell)
            
            end_tag(doc, 'c')


        end_tag(doc, 'row')


def save_dump(workbook, filename):

    pass
