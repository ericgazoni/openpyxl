# file openpyxl/tests/test_meta.py

# Python stdlib imports
from __future__ import with_statement
import os.path

# package imports
from openpyxl.tests.helper import BaseTestCase, DATADIR
from openpyxl.writer.workbook import write_content_types, write_root_rels
from openpyxl.workbook import Workbook


class TestWriteMeta(BaseTestCase):

    def test_write_content_types(self):
        wb = Workbook()
        wb.create_sheet()
        wb.create_sheet()
        content = write_content_types(wb)
        reference_file = os.path.join(DATADIR, 'writer', 'expected',
                '[Content_Types].xml')
        self.assertEqualsFileContent(reference_file, fixture=content)

    def test_write_root_rels(self):
        wb = Workbook()
        content = write_root_rels(wb)
        reference_file = os.path.join(DATADIR, 'writer', 'expected', '.rels')
        self.assertEqualsFileContent(reference_file, fixture=content)
