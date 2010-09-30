# file openpyxl/tests/test_style.py

# Python stdlib imports
from __future__ import with_statement
import os.path
import datetime

# 3rd party imports
from nose.tools import eq_

# package imports
from openpyxl.tests.helper import DATADIR, assert_equals_file_content
from openpyxl.reader.style import read_style_table
from openpyxl.workbook import Workbook
from openpyxl.style import NumberFormat
from openpyxl.writer.styles import  create_style_table
from openpyxl.writer.styles import write_style_table


class TestCreateStyle(object):

    @classmethod
    def setup_class(cls):
        now = datetime.datetime.now()
        cls.workbook = Workbook()
        cls.worksheet = cls.workbook.create_sheet()
        cls.worksheet.cell(coordinate='A1').value = '12.34%'
        cls.worksheet.cell(coordinate='B4').value = now
        cls.worksheet.cell(coordinate='B5').value = now
        cls.worksheet.cell(coordinate='C14').value = u'This is a test'
        cls.worksheet.cell(coordinate='D9').value = '31.31415'
        cls.worksheet.cell(coordinate='D9').style.number_format.format_code = \
                NumberFormat.FORMAT_NUMBER_00

    def test_create_style_table(self):
        table = create_style_table(self.workbook)
        eq_(3, len(table))

    def test_write_style_table(self):
        table = create_style_table(self.workbook)
        content = write_style_table(table)
        reference_file = os.path.join(
                DATADIR, 'writer', 'expected', 'simple-styles.xml')
        assert_equals_file_content(reference_file, content)


def test_read_style():
    reference_file = os.path.join(DATADIR, 'reader', 'simple-styles.xml')
    with open(reference_file, 'r') as handle:
        content = handle.read()
    style_table = read_style_table(content)
    eq_(4, len(style_table))
    eq_(NumberFormat._BUILTIN_FORMATS[9],
            style_table[1].number_format.format_code)
    eq_('yyyy-mm-dd', style_table[2].number_format.format_code)


def test_read_cell_style():
    reference_file = os.path.join(
            DATADIR, 'reader', 'empty-workbook-styles.xml')
    with open(reference_file, 'r') as handle:
        content = handle.read()
    style_table = read_style_table(content)
    eq_(2, len(style_table))
