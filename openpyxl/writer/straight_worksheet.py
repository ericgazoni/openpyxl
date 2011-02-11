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

from openpyxl.cell import column_index_from_string, get_column_letter
from openpyxl.shared.xmltools import XMLGenerator, get_document_content, \
        start_tag, end_tag, tag
from openpyxl.shared.date_time import SharedDate

STYLES = {'datetime' : {'type':'n',
                        'style':'1'},
          'string':{'type':'s',
                    'style':'0'},
          'numeric':{'type':'n',
                     'style':'0'}
        }


class StraightWorksheet(object):

    def __init__(self):

        self._max_col = 0
        self._max_row = 0
        self._fileobj = None

        self._shared_date = SharedDate()d

    def append(self, row):

        doc = self._fileobjd

        self._max_row += 1
        span = len(row)
        self._max_col = max(self._max_col, span)

        row_idx = self._max_row

        attrs = {'r': '%d' % row_idx,
                 'spans': '1:%d' % span}

        start_tag(doc, 'row', attrs)

        for col_idx, cell in enumerate(row):

            value = cell.value
            coordinate = '%s%d' % (get_column_letter(col_idx, row_idx)) 
            attributes = {'r': coordinate}


            if isinstance(cell, (int, float)):
                dtype = 'numeric'
            elif isinstance(cell, (datetime.datetime,
                                    datetime.date)):
                dtype = 'datetime'
                cell = self._shared_date.datetime_to_julian(cell)
            else:
                dtype = 'string'

            attribute['t'] = STYLES[dtype]['type']
            attribute['s'] = STYLES[dtype]['style']

            start_tag(doc, 'c', attributes)

            if cell is None:
                tag(doc, 'v', body='')
            elif cell.data_type == cell.TYPE_STRING:
                tag(doc, 'v', body = '%s' % string_table[value])
            elif cell.data_type == cell.TYPE_NUMERIC:
                tag(doc, 'v', body = '%s' % value)
            else:
                tag(doc, 'v', body = '%s' % value)
            
            end_tag(doc, 'c')


        end_tag(doc, 'row')

