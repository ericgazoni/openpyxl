# file openpyxl/tests/test_write.py
from __future__ import with_statement
import os.path as osp
from openpyxl.tests.helper import BaseTestCase, TMPDIR, DATADIR
from openpyxl.workbook import Workbook
from openpyxl.writer.excel import ExcelWriter
from openpyxl.writer.workbook import write_workbook, write_workbook_rels
from openpyxl.writer.worksheet import write_worksheet, write_worksheet_rels
from openpyxl.writer.strings import write_string_table
from openpyxl.writer.styles import create_style_table

class TestWriter(BaseTestCase):

    def test_write_empty_workbook(self):

        wb = Workbook()

        ew = ExcelWriter(workbook = wb)

        dest_filename = osp.join(TMPDIR, 'empty_book.xlsx')

        ew.save(filename = dest_filename)

        self.assertTrue(osp.isfile(dest_filename))

class TestWriteWorkbook(BaseTestCase):

    def test_write_workbook_rels(self):

        wb = Workbook()

        content = write_workbook_rels(workbook = wb)

        self.assertEqualsFileContent(reference_file = osp.join(DATADIR, 'writer', 'expected', 'workbook.xml.rels'),
                                     fixture = content)

    def test_write_workbook(self):

        wb = Workbook()

        content = write_workbook(workbook = wb)

        self.assertEqualsFileContent(reference_file = osp.join(DATADIR, 'writer', 'expected', 'workbook.xml'),
                                     fixture = content)

class TestWriteStrings(BaseTestCase):

    def test_write_string_table(self):

        table = {'hello' : 1,
                 'world' : 2,
                 'nice' : 3}

        content = write_string_table(string_table = table)

        self.assertEqualsFileContent(reference_file = osp.join(DATADIR, 'writer', 'expected', 'sharedStrings.xml'),
                                     fixture = content)

class TestWriteWorksheet(BaseTestCase):

    def test_write_worksheet(self):

        wb = Workbook()

        ws = wb.create_sheet()

        ws.cell('F42').value = 'hello'

        content = write_worksheet(worksheet = ws, string_table = {'hello' : 0}, style_table = {})

        self.assertEqualsFileContent(reference_file = osp.join(DATADIR, 'writer', 'expected', 'sheet1.xml'),
                                     fixture = content)


    def test_write_worksheet_with_formula(self):

        wb = Workbook()

        ws = wb.create_sheet()

        ws.cell('F1').value = 10
        ws.cell('F2').value = 32
        ws.cell('F3').value = '=F1+F2'

        content = write_worksheet(worksheet = ws, string_table = { }, style_table = {})

        self.assertEqualsFileContent(reference_file = osp.join(DATADIR, 'writer', 'expected', 'sheet1_formula.xml'),
                                     fixture = content)

    def test_write_worksheet_with_style(self):

        wb = Workbook()

        ws = wb.create_sheet()

        ws.cell('F1').value = '13%'

        shared_style_table = create_style_table(workbook = wb)
        style_id_by_hash = dict([(style.__crc__(), id) for style, id in shared_style_table.iteritems()])

        content = write_worksheet(worksheet = ws, string_table = { }, style_table = style_id_by_hash)

        self.assertEqualsFileContent(reference_file = osp.join(DATADIR, 'writer', 'expected', 'sheet1_style.xml'),
                                     fixture = content)

    def test_write_worksheet_with_height(self):

        wb = Workbook()

        ws = wb.create_sheet()

        ws.cell('F1').value = 10
        ws.row_dimensions[ws.cell('F1').row].height = 30

        content = write_worksheet(worksheet = ws, string_table = { }, style_table = {})

        self.assertEqualsFileContent(reference_file = osp.join(DATADIR, 'writer', 'expected', 'sheet1_height.xml'),
                                     fixture = content)

    def test_write_worksheet_with_hyperlink(self):
        wb = Workbook()
        ws = wb.create_sheet()
        ws.cell('A1').value = "test"
        ws.cell('A1').hyperlink = "http://test.com"
        content = write_worksheet(worksheet = ws, string_table = {'test' : 0}, style_table = {})
        self.assertEqualsFileContent(reference_file = osp.join(DATADIR,
            'writer', 'expected', 'sheet1_hyperlink.xml'), fixture = content)

    def test_write_worksheet_with_hyperlink_relationships(self):
        wb = Workbook()
        ws = wb.create_sheet()
        self.assertEqual(0, len(ws.relationships))
        ws.cell('A1').value = "test"
        ws.cell('A1').hyperlink = "http://test.com/"
        self.assertEqual(1, len(ws.relationships))
        ws.cell('A2').value = "test"
        ws.cell('A2').hyperlink = "http://test2.com/"
        self.assertEqual(2, len(ws.relationships))

        content = write_worksheet_rels(worksheet = ws)
        self.assertEqualsFileContent(reference_file = osp.join(DATADIR,
            'writer', 'expected', 'sheet1_hyperlink.xml.rels'), fixture = content)

    def test_hyperlink_value(self):
        wb = Workbook()
        ws = wb.create_sheet()
        ws.cell('A1').hyperlink = "http://test.com"
        self.assertEqual("http://test.com", ws.cell('A1').value)
        content = write_worksheet(worksheet = ws, string_table = {"http://test.com" : 0}, style_table = {})
        ws.cell('A1').value = "test"
        self.assertEqual("test", ws.cell('A1').value)
