# file openpyxl/tests/test_meta.py

# Python stdlib imports
from __future__ import with_statement
import os.path

# package imports
from openpyxl.tests.helper import DATADIR, assert_equals_file_content
from openpyxl.writer.workbook import write_content_types, write_root_rels
from openpyxl.workbook import Workbook


def test_write_content_types():
    wb = Workbook()
    wb.create_sheet()
    wb.create_sheet()
    content = write_content_types(wb)
    reference_file = os.path.join(DATADIR, 'writer', 'expected',
            '[Content_Types].xml')
    assert_equals_file_content(reference_file, content)


def test_write_root_rels():
    wb = Workbook()
    content = write_root_rels(wb)
    reference_file = os.path.join(DATADIR, 'writer', 'expected', '.rels')
    assert_equals_file_content(reference_file, content)
