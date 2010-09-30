# file openpyxl/tests/test_props.py

# Python stdlib imports
from __future__ import with_statement
import os.path
import datetime

# package imports
from openpyxl.tests.helper import BaseTestCase, DATADIR, TMPDIR
from openpyxl.reader.workbook import read_properties_core, read_sheets_titles, get_number_of_parts
from openpyxl.writer.workbook import write_properties_core, write_properties_app
from openpyxl.shared.ooxml import ARC_APP, ARC_CORE
from openpyxl.shared.zip import ZipArchive
from openpyxl.workbook import DocumentProperties, Workbook


class TestReaderProps(BaseTestCase):

    def setUp(self):
        self.gen_filename = os.path.join(DATADIR, 'genuine', 'empty.xlsx')

    def test_read_properties_core(self):
        archive = ZipArchive(filename=self.gen_filename)
        content = archive.get_from_name(arc_name=ARC_CORE)
        prop = read_properties_core(xml_source=content)
        self.assertEqual(prop.creator, '*.*')
        self.assertEqual(prop.last_modified_by, '*.*')
        self.assertEqual(prop.created,
                datetime.datetime(2010, 4, 9, 20, 43, 12))
        self.assertEqual(prop.modified,
                datetime.datetime(2010, 4, 11, 16, 20, 29))

    def test_read_sheets_titles(self):
        archive = ZipArchive(filename = self.gen_filename)
        content = archive.get_from_name(arc_name = ARC_APP)
        sheet_titles = read_sheets_titles(xml_source = content)
        self.assertEqual(sheet_titles,
                ['Sheet1 - Text', 'Sheet2 - Numbers', 'Sheet3 - Formulas'])


class TestReaderPropsMixed(BaseTestCase):

    def setUp(self):
        self.reference_filename = os.path.join(DATADIR, 'reader', 'app-multi-titles.xml')
        with open(self.reference_filename) as ref_file:
            self.content = ref_file.read()

    def test_read_sheet_titles_mixed(self):
        sheet_titles = read_sheets_titles(xml_source = self.content)
        self.assertEqual(sheet_titles,
                ['ToC', 'ContractYear', 'ContractTier', 'Demand',
                'LinearizedFunction', 'Market', 'Transmission'])

    def test_number_of_parts(self):
        parts_number = get_number_of_parts(xml_source = self.content)
        self.assertEqual(parts_number,
                ({'Worksheets': 7, 'Named Ranges': 7},
                ['Worksheets', 'Named Ranges']))


class TestWriteProps(BaseTestCase):

    def setUp(self):
        self.tmp_filename = os.path.join(TMPDIR, 'test.xlsx')
        self.prop = DocumentProperties()

    def test_write_properties_core(self):
        self.prop.creator = 'TEST_USER'
        self.prop.last_modified_by = 'SOMEBODY'
        self.prop.created = datetime.datetime(2010, 4, 1, 20, 30, 00)
        self.prop.modified = datetime.datetime(2010, 4, 5, 14, 5, 30)
        content = write_properties_core(self.prop)
        self.assertEqualsFileContent(os.path.join(DATADIR, 'writer', 'expected', 'core.xml'), content)

    def test_write_properties_app(self):
        wb = Workbook()
        wb.create_sheet()
        wb.create_sheet()
        content = write_properties_app(wb)
        self.assertEqualsFileContent(os.path.join(DATADIR, 'writer',
                'expected', 'app.xml'), content)
