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

import os.path as osp
from openpyxl.tests.helper import BaseTestCase, DATADIR, TMPDIR
import datetime

from openpyxl.reader.workbook import read_properties_core, read_sheets_titles, get_number_of_parts
from openpyxl.writer.workbook import write_properties_core, write_properties_app

from openpyxl.shared.ooxml import ARC_APP, ARC_CORE

from openpyxl.shared.zip import ZipArchive
from openpyxl.workbook import DocumentProperties, Workbook

class TestReaderProps(BaseTestCase):

    def setUp(self):

        self.gen_filename = osp.join(DATADIR, 'genuine', 'empty.xlsx')

    def test_read_properties_core(self):

        archive = ZipArchive(filename = self.gen_filename)

        content = archive.get_from_name(arc_name = ARC_CORE)

        prop = read_properties_core(xml_source = content)

        self.assertEqual(prop.creator, '*.*')
        self.assertEqual(prop.last_modified_by, '*.*')

        self.assertEqual(prop.created, datetime.datetime(2010, 4, 9, 20, 43, 12))
        self.assertEqual(prop.modified, datetime.datetime(2010, 4, 11, 16, 20, 29))

    def test_read_sheets_titles(self):

        archive = ZipArchive(filename = self.gen_filename)

        content = archive.get_from_name(arc_name = ARC_APP)

        sheet_titles = read_sheets_titles(xml_source = content)

        self.assertEqual(sheet_titles, ['Sheet1 - Text', 'Sheet2 - Numbers', 'Sheet3 - Formulas'])

class TestReaderPropsMixed(BaseTestCase):

    def setUp(self):

        self.reference_filename = osp.join(DATADIR, 'reader', 'app-multi-titles.xml')

        with open(self.reference_filename) as ref_file:

            self.content = ref_file.read()

    def test_read_sheet_titles_mixed(self):

        sheet_titles = read_sheets_titles(xml_source = self.content)

        self.assertEqual(sheet_titles, ['ToC', 'ContractYear', 'ContractTier',
                                        'Demand', 'LinearizedFunction',
                                        'Market', 'Transmission'])

    def test_number_of_parts(self):

        parts_number = get_number_of_parts(xml_source = self.content)

        self.assertEqual(parts_number, ({'Worksheets':7,
                                        'Named Ranges':7}, ['Worksheets', 'Named Ranges']))


class TestWriteProps(BaseTestCase):

    def setUp(self):

        self.tmp_filename = osp.join(TMPDIR, 'test.xlsx')
        self.prop = DocumentProperties()

    def test_write_properties_core(self):

        self.prop.creator = 'TEST_USER'
        self.prop.last_modified_by = 'SOMEBODY'

        self.prop.created = datetime.datetime(2010, 4, 1, 20, 30, 00)
        self.prop.modified = datetime.datetime(2010, 4, 5, 14, 5, 30)

        content = write_properties_core(self.prop)

        self.assertEqualsFileContent(osp.join(DATADIR, 'writer', 'expected', 'core.xml'), content)

    def test_write_properties_app(self):

        wb = Workbook()

        wb.create_sheet()

        wb.create_sheet()

        content = write_properties_app(wb)

        self.assertEqualsFileContent(osp.join(DATADIR, 'writer', 'expected', 'app.xml'), content)
