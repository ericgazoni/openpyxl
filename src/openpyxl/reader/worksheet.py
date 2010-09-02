# coding=UTF-8
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

from openpyxl.shared.xmltools import fromstring, QName
from openpyxl.cell import Cell
from openpyxl.worksheet import Worksheet

def read_worksheet(xml_source, parent, preset_title, string_table, style_table):

    xmlns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

    ws = Worksheet(parent_workbook = parent, title = preset_title)

    root = fromstring(text = xml_source)

    sheet_data = root.find(QName(xmlns, 'sheetData').text)

    for row in sheet_data.getchildren():

        for cell in row.getchildren():

            coordinate = cell.get('r')
            data_type = cell.get('t', 'n')
            value = cell.findtext(QName(xmlns, 'v').text)

            if value is not None:

                if data_type == Cell.TYPE_STRING:
                    value = string_table[int(value)]

                ws.cell(coordinate).value = value

                style_id = cell.get('s')
                if style_id is not None:
                    ws._styles[coordinate] = style_table[int(style_id)]

    return ws
