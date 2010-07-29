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
from openpyxl.writer.excel import ExcelWriter
from openpyxl.reader.style import read_style_table
from openpyxl.workbook import Workbook
from openpyxl.style import NumberFormat
import datetime

from openpyxl.writer.styles import  create_style_table

from openpyxl.writer.styles import write_style_table

class TestCreateStyle(BaseTestCase):

    def setUp(self):

        self.workbook = Workbook()

        self.worksheet = self.workbook.create_sheet()

        self.worksheet.cell(coordinate = 'A1').value = '12.34%'

        now = datetime.datetime.now()
        self.worksheet.cell(coordinate = 'B4').value = now
        self.worksheet.cell(coordinate = 'B5').value = now


        self.worksheet.cell(coordinate = 'C14').value = u'This is a test'


        self.worksheet.cell(coordinate = 'D9').value = '31.31415'
        self.worksheet.cell(coordinate = 'D9').style.number_format.format_code = NumberFormat.FORMAT_NUMBER_00


    def test_create_style_table(self):

        table = create_style_table(workbook = self.workbook)

        self.assertEqual(3, len(table))

    def test_write_style_table(self):

        table = create_style_table(workbook = self.workbook)

        content = write_style_table(style_table = table)

        reference_file = osp.join(DATADIR, 'writer', 'expected', 'simple-styles.xml')
        self.assertEqualsFileContent(reference_file = reference_file,
                                     fixture = content)

class TestReadStyle(BaseTestCase):

    def test_read_style(self):

        reference_file = osp.join(DATADIR, 'reader', 'simple-styles.xml')

        with open(reference_file) as ref_file:

            content = ref_file.read()

            style_table = read_style_table(xml_source = content)

        self.assertEqual(4, len(style_table))

        self.assertEqual(NumberFormat._BUILTIN_FORMATS[9], style_table[1].number_format.format_code)

        self.assertEqual('yyyy-mm-dd', style_table[2].number_format.format_code)



