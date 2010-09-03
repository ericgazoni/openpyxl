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

from __future__ import with_statement
from openpyxl.cell import column_index_from_string
from openpyxl.shared.xmltools import ElementTree, Element, SubElement, \
    get_document_content, get_tempfile, start_tag, end_tag, tag, XMLGenerator

def row_sort(cell):

    return column_index_from_string(cell.column)

def write_worksheet(worksheet, string_table, style_table):

    filename = get_tempfile()

    with open(filename, 'w') as xml_file:

        doc = XMLGenerator(out = xml_file, encoding = 'utf-8')

        start_tag(doc = doc,
                  name = 'worksheet',
                  attr = {'xml:space':'preserve',
                          'xmlns':'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                          'xmlns:r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'})

        start_tag(doc, 'sheetPr')

        tag(doc, 'outlinePr', {'summaryBelow' : '%d' % (worksheet.show_summary_below),
                               'summaryRight' : '%d' % (worksheet.show_summary_right)})

        end_tag(doc, 'sheetPr')

        tag(doc, 'dimension', {'ref' : '%s' % worksheet.calculate_dimension()})

        start_tag(doc, 'sheetViews')
        start_tag(doc, 'sheetView', {'workbookViewId' : '0'})
        tag(doc, 'selection', {'activeCell' : worksheet.active_cell,
                               'sqref' : worksheet.selected_cell})
        end_tag(doc, 'sheetView')
        end_tag(doc, 'sheetViews')

        tag(doc, 'sheetFormatPr', {'defaultRowHeight' : '15'})

        write_worksheet_cols(doc, worksheet)

        write_worksheet_data(doc, worksheet, string_table, style_table)

        end_tag(doc, 'worksheet')

        doc.endDocument()

    return filename

def write_worksheet_cols(doc, worksheet):

    if worksheet.column_dimensions:

        start_tag(doc, 'cols')

        for column_string, columndimension in worksheet.column_dimensions.iteritems():

            cidx = column_index_from_string(column = column_string)

            col_def = {}
            col_def['collapsed'] = str(columndimension.style_index)
            col_def['min'] = str(cidx)
            col_def['max'] = str(cidx)

            if columndimension.width != worksheet.default_column_dimension.width:
                col_def['customWidth'] = 'true'

            if not columndimension.visible:
                col_def['hidden'] = 'true'

            if columndimension.outline_level > 0:
                col_def['outlineLevel'] = str(columndimension.outline_level)

            if columndimension.collapsed:
                col_def['collapsed'] = 'true'

            if columndimension.auto_size:
                col_def['bestFit'] = 'true'
            if columndimension.width > 0:
                col_def['width'] = str(columndimension.width)
            else:
                col_def['width'] = '9.10'

            tag(doc, 'col', col_def)

        end_tag(doc, 'cols')

def write_worksheet_data(doc, worksheet, string_table, style_table):

    style_id_by_hash = dict([(style.__crc__(), id) for style, id in style_table.iteritems()])

    start_tag(doc, 'sheetData')

    max_column = worksheet.get_highest_column()

    cells_by_row = {}
    for cell in worksheet.get_cell_collection():
        cells_by_row.setdefault(cell.row, []).append(cell)

    for row_idx in sorted(cells_by_row):
        row_dimension = worksheet.row_dimensions[row_idx]

        start_tag(doc, 'row', {'r' : '%d' % row_idx,
                               'spans' : '1:%d' % max_column})

        row_cells = cells_by_row[row_idx]

        sorted_cells = sorted(row_cells, key = row_sort)

        for cell in sorted_cells:

            value = cell._value

            coordinate = cell.get_coordinate()

            attributes = {'r' : coordinate}
            attributes['t'] = cell.data_type

            if coordinate in worksheet._styles:

                attributes['s'] = '%d' % style_id_by_hash[worksheet._styles[coordinate].__crc__()]

            start_tag(doc, 'c', attributes)

            if cell.data_type == cell.TYPE_STRING:
                tag(doc, 'v', body = '%s' % string_table[value])
            elif cell.data_type == cell.TYPE_FORMULA:
                tag(doc, 'f', body = '%s' % value[1:])
                tag(doc, 'v')
            elif cell.data_type == cell.TYPE_NUMERIC:
                tag(doc, 'v', body = '%s' % value)
            else:
                tag(doc, 'v', body = '%s' % value)

            end_tag(doc, 'c')

        end_tag(doc, 'row')

    end_tag(doc, 'sheetData')
