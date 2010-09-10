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

from openpyxl.shared.ooxml import ARC_SHARED_STRINGS, ARC_CORE, ARC_APP, ARC_WORKBOOK, PACKAGE_WORKSHEETS, ARC_STYLE
from openpyxl.shared.zip import ZipArchive

from openpyxl.workbook import Workbook

from openpyxl.reader.strings import read_string_table
from openpyxl.reader.style import read_style_table
from openpyxl.reader.workbook import read_sheets_titles, read_named_ranges, read_properties_core
from openpyxl.reader.worksheet import read_worksheet

def load_workbook(filename):
    """Open the given filename and return the workbook
    
    :param filename: the path to open
    :type filename: string
    
    :rtype: :class:`openpyxl.workbook.Workbook`
    
    """

    archive = ZipArchive(filename = filename, mode = 'r')

    # define the result workbook
    wb = Workbook()

    try:
        # add properties
        wb.properties = read_properties_core(xml_source = archive.get_from_name(arc_name = ARC_CORE))


        # add worksheets        
        wb.worksheets = [] # remove preset worksheet
        sheet_names = read_sheets_titles(xml_source = archive.get_from_name(arc_name = ARC_APP))

        if archive.is_in_archive(arc_name = ARC_SHARED_STRINGS):
            string_table = read_string_table(xml_source = archive.get_from_name(arc_name = ARC_SHARED_STRINGS))
        else:
            string_table = {}

        style_table = read_style_table(xml_source = archive.get_from_name(arc_name = ARC_STYLE))

        for i, sheet_name in enumerate(sheet_names):

            worksheet_path = '%s/%s' % (PACKAGE_WORKSHEETS, 'sheet%d.xml' % (i + 1))

            new_ws = read_worksheet(xml_source = archive.get_from_name(arc_name = worksheet_path),
                                    parent = wb,
                                    preset_title = sheet_name,
                                    string_table = string_table,
                                    style_table = style_table)

            wb.add_sheet(worksheet = new_ws, index = i)

        # add named ranges
        wb._named_ranges = read_named_ranges(xml_source = archive.get_from_name(arc_name = ARC_WORKBOOK),
                                             workbook = wb)

    finally:
        archive.close()

    return wb
