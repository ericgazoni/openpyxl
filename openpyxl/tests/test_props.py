# Copyright (c) 2010-2014 openpyxl
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
# @author: see AUTHORS file

# Python stdlib imports
from zipfile import ZipFile, ZIP_DEFLATED
from datetime import datetime
import os.path

# package imports
from openpyxl.tests.helper import (
    DATADIR,
    TMPDIR,
    make_tmpdir,
    clean_tmpdir,
    compare_xml
    )
from openpyxl.reader.workbook import read_properties_core
from openpyxl.writer.workbook import (
    write_properties_core,
    write_properties_app
)
from openpyxl.xml.constants import ARC_CORE
from openpyxl.date_time import CALENDAR_WINDOWS_1900
from openpyxl.workbook import DocumentProperties, Workbook


class TestReaderProps(object):

    @classmethod
    def setup_class(cls):
        cls.genuine_filename = os.path.join(DATADIR, 'genuine', 'empty.xlsx')
        cls.archive = ZipFile(cls.genuine_filename, 'r', ZIP_DEFLATED)

    @classmethod
    def teardown_class(cls):
        cls.archive.close()

    def test_read_properties_core(self):
        content = self.archive.read(ARC_CORE)
        prop = read_properties_core(content)
        assert prop.creator == '*.*'
        assert prop.last_modified_by == 'Charlie Clark'
        assert prop.created == datetime(2010, 4, 9, 20, 43, 12)
        assert prop.modified ==  datetime(2014, 1, 2, 14, 53, 6)

    def test_read_sheets_titles(self):
        from openpyxl.reader.workbook import read_sheets
        sheet_titles = [s[0] for s in read_sheets(self.archive)]
        assert sheet_titles == ['Sheet1 - Text', 'Sheet2 - Numbers', 'Sheet3 - Formulas', 'Sheet4 - Dates']


class TestLibreOfficeCompat(object):
    """
    Just tests that the correct date/time format is returned from LibreOffice saved version
    """

    @classmethod
    def setup_class(cls):
        cls.genuine_filename = os.path.join(DATADIR, 'genuine', 'empty_libre.xlsx')
        cls.archive = ZipFile(cls.genuine_filename, 'r', ZIP_DEFLATED)

    @classmethod
    def teardown_class(cls):
        cls.archive.close()

    def test_read_properties_core(self):
        content = self.archive.read(ARC_CORE)
        prop = read_properties_core(content)
        assert prop.excel_base_date == CALENDAR_WINDOWS_1900

    def test_read_sheets_titles(self):
        from openpyxl.reader.workbook import read_sheets
        sheet_titles = [s[0] for s in read_sheets(self.archive)]
        assert sheet_titles == ['Sheet1 - Text', 'Sheet2 - Numbers', 'Sheet3 - Formulas', 'Sheet4 - Dates']


class TestWriteProps(object):

    @classmethod
    def setup_class(cls):
        make_tmpdir()
        cls.tmp_filename = os.path.join(TMPDIR, 'test.xlsx')
        cls.prop = DocumentProperties()

    @classmethod
    def teardown_class(cls):
        clean_tmpdir()

    def test_write_properties_core(self):
        self.prop.creator = 'TEST_USER'
        self.prop.last_modified_by = 'SOMEBODY'
        self.prop.created = datetime(2010, 4, 1, 20, 30, 00)
        self.prop.modified = datetime(2010, 4, 5, 14, 5, 30)
        content = write_properties_core(self.prop)
        reference_file = os.path.join(DATADIR, 'writer', 'expected', 'core.xml')
        with open(reference_file) as expected:
            diff = compare_xml(content, expected.read())
            assert diff is None

    def test_write_properties_app(self):
        wb = Workbook()
        wb.create_sheet()
        wb.create_sheet()
        content = write_properties_app(wb)
        reference_file = os.path.join(DATADIR, 'writer', 'expected', 'app.xml')
        with open(reference_file) as expected:
            diff = compare_xml(content, expected.read())
            assert diff is None
