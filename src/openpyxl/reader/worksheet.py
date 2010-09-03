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

from openpyxl.cell import Cell
from openpyxl.shared.xmltools import fromstring, QName
from openpyxl.worksheet import Worksheet
from xml.sax import parseString
from xml.sax.handler import ContentHandler

class WorksheetReader(ContentHandler):

    def __init__(self, ws, string_table, style_table):

        ContentHandler.__init__(self)
        self.ws = ws
        self.string_table = string_table
        self.style_table = style_table

    def startElement(self, name, attrs):

        if name == 'c':
            self.coordinate = attrs.get('r')
            self.data_type = attrs.get('t', 'n')
            self.style_id = attrs.get('s')

    def characters(self, value):

        if value is not None:

            if self.data_type == Cell.TYPE_STRING:
                value = self.string_table.get(int(value))

            self.ws.cell(self.coordinate).value = value

            if self.style_id is not None:
                self.ws._styles[self.coordinate] = self.style_table.get(int(self.style_id))

def read_worksheet(xml_source, parent, preset_title, string_table, style_table):

    ws = Worksheet(parent_workbook = parent, title = preset_title)

    h = WorksheetReader(ws, string_table, style_table)

    parseString(string = xml_source, handler = h)

    return ws
