# file openpyxl/tests/test_number_format.py
from __future__ import with_statement
import os.path as osp
import datetime
from openpyxl.tests.helper import BaseTestCase, DATADIR

from openpyxl.workbook import Workbook
from openpyxl.worksheet import Worksheet
from openpyxl.cell import Cell
from openpyxl.style import NumberFormat


from openpyxl.shared.date_time import SharedDate

class TestNumberFormat(BaseTestCase):

    def setUp(self):

        self.workbook = Workbook()
        self.worksheet = Worksheet(parent_workbook = self.workbook,
                                   title = 'Test')

        self.sd = SharedDate()

    def test_convert_date_to_julian(self):

        self.assertEqual(40167, self.sd.to_julian(year = 2009, month = 12, day = 20))

    def test_convert_date_from_julian(self):

        self.assertEqual(datetime.datetime(2009, 12, 20) , self.sd.from_julian(40167))

    def test_convert_datetime_to_julian(self):

        self.assertEqual(40167, self.sd.datetime_to_julian(date = datetime.datetime(2009, 12, 20)))

    def test_insert_float(self):

        self.worksheet.cell(coordinate = 'A1').value = 3.14

        self.assertEqual(Cell.TYPE_NUMERIC, self.worksheet.cell(coordinate = 'A1')._data_type)

    def test_insert_percentage(self):

        self.worksheet.cell(coordinate = 'A1').value = '3.14%'

        self.assertEqual(Cell.TYPE_NUMERIC, self.worksheet.cell(coordinate = 'A1')._data_type)

        self.assertAlmostEqual(0.0314, self.worksheet.cell(coordinate = 'A1').value)

    def test_insert_date(self):

        self.worksheet.cell(coordinate = 'A1').value = datetime.datetime.now()

        self.assertEqual(Cell.TYPE_NUMERIC, self.worksheet.cell(coordinate = 'A1')._data_type)

    def test_internal_date(self):

        dt = datetime.datetime(2010, 7, 13, 6, 37, 41)

        self.worksheet.cell(coordinate = 'A3').value = dt

        self.assertEqual(40372.27616898148, self.worksheet.cell(coordinate = 'A3')._value)


    def test_date_interpretation(self):

        dt = datetime.datetime(2010, 7, 13, 6, 37, 41)

        self.worksheet.cell(coordinate = 'A3').value = dt

        self.assertEqual(dt, self.worksheet.cell(coordinate = 'A3').value)

    def test_number_format_style(self):

        ws = self.worksheet

        ws.cell('A1').value = '12.6%'

        self.assertEqual(NumberFormat.FORMAT_PERCENTAGE, ws.cell(coordinate = 'A1').style.number_format.format_code)

