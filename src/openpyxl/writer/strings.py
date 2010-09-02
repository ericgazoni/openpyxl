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

from openpyxl.shared.xmltools import ElementTree, Element, SubElement

from openpyxl.shared.xmltools import get_document_content

def create_string_table(workbook):

    strings_list = []

    for sheet in workbook.worksheets:

        for cell in sheet.get_cell_collection():

            if cell.data_type == cell.TYPE_STRING:

                strings_list.append(cell._value)

    return dict((key, i) for i, key in enumerate(set(strings_list)))

def write_string_table(string_table):

    root = Element('sst', {'xmlns' : 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                           'uniqueCount' : '%d' % len(string_table)})

    strings_to_write = sorted(string_table.iteritems(), key = lambda pair:pair[1])

    for key in [key for (key, rank) in strings_to_write]:

        si = SubElement(root, 'si')
        t = SubElement(si, 't')
        t.text = key

        if key.strip() != key:
            t.set('xml:space' , 'preserve')


    return get_document_content(xml_node = root)
