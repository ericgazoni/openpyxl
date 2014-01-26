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
from datetime import datetime, date, timedelta

import pytest

# package imports
from openpyxl.compat import safe_string
from openpyxl.workbook import Workbook
from openpyxl.worksheet import Worksheet
from openpyxl.cell import Cell
from openpyxl.styles import NumberFormat
from openpyxl.date_time import SharedDate, CALENDAR_MAC_1904, CALENDAR_WINDOWS_1900


class TestNumberFormat(object):

    @classmethod
    def setup_class(cls):
        cls.workbook = Workbook()
        cls.worksheet = Worksheet(cls.workbook, 'Test')
        cls.sd = SharedDate()

    def test_convert_date_to_julian(self):
        assert 40167 == self.sd.to_julian(2009, 12, 20)

    @pytest.mark.parametrize("value, expected",
                             [
                                 (40167, datetime(2009, 12, 20)),
                                 (21980, datetime(1960,  3,  5)),
                             ])
    def test_convert_date_from_julian(self, value, expected):
        assert self.sd.from_julian(value) == expected

    def test_convert_datetime_to_julian(self):
        assert 40167 == self.sd.datetime_to_julian(datetime(2009, 12, 20))
        assert 40196.5939815 == self.sd.datetime_to_julian(datetime(2010, 1, 18, 14, 15, 20, 1600))

    def test_convert_timedelta_to_julian(self):
        assert 1.125 == self.sd.datetime_to_julian(timedelta(days=1, hours=3))

    def test_insert_float(self):
        self.worksheet.cell('A1').value = 3.14
        assert Cell.TYPE_NUMERIC == self.worksheet.cell('A1')._data_type

    def test_insert_percentage(self):
        self.worksheet.cell('A1').value = '3.14%'
        assert Cell.TYPE_NUMERIC == self.worksheet.cell('A1')._data_type
        assert safe_string(0.0314) == safe_string(self.worksheet.cell('A1').value)

    def test_insert_datetime(self):
        self.worksheet.cell('A1').value = date.today()
        assert Cell.TYPE_NUMERIC == self.worksheet.cell('A1')._data_type

    def test_insert_date(self):
        self.worksheet.cell('A1').value = datetime.now()
        assert Cell.TYPE_NUMERIC == self.worksheet.cell('A1')._data_type

    def test_internal_date(self):
        dt = datetime(2010, 7, 13, 6, 37, 41)
        self.worksheet.cell('A3').value = dt
        assert 40372.27616898148 == self.worksheet.cell('A3')._value

    def test_datetime_interpretation(self):
        dt = datetime(2010, 7, 13, 6, 37, 41)
        self.worksheet.cell('A3').value = dt
        assert dt == self.worksheet.cell('A3').value

    def test_date_interpretation(self):
        dt = date(2010, 7, 13)
        self.worksheet.cell('A3').value = dt
        assert datetime(2010, 7, 13, 0, 0) == self.worksheet.cell('A3').value

    def test_number_format_style(self):
        self.worksheet.cell('A1').value = '12.6%'
        assert NumberFormat.FORMAT_PERCENTAGE == self.worksheet.cell('A1').style.number_format.format_code

    @pytest.mark.parametrize("date_string, ordinal",
            [
            ('1900-01-15', 15),
            ('1900-02-28', 59),
            ('1900-03-01', 61),
            ('1901-01-01', 367),
            ('9999-12-31', 2958465),
            ]
            )
    def test_date_format_on_non_date(self, date_string, ordinal):
        cell = self.worksheet.cell('A1')
        cell.value = datetime.strptime(date_string, '%Y-%m-%d')
        assert cell.internal_value == ordinal

    def test_1900_leap_year(self):
        with pytest.raises(ValueError):
            self.sd.from_julian(60)
        with pytest.raises(ValueError):
            self.sd.to_julian(1900, 2, 29)

    bad_dates = (
        (1776,  7,  4),
        (1899, 12, 31),
    )
    @pytest.mark.parametrize("dt", bad_dates)
    def test_bad_date(self, dt):
        with pytest.raises(ValueError):
            self.sd.to_julian(*dt)

    def test_bad_julian_date(self):
        with pytest.raises(ValueError):
            self.sd.from_julian(-1)

    def test_mac_date(self):
        self.sd.excel_base_date = CALENDAR_MAC_1904

        datetuple = (2011, 10, 31)

        dt = date(datetuple[0],datetuple[1],datetuple[2])
        julian = self.sd.to_julian(datetuple[0],datetuple[1],datetuple[2])
        reverse = self.sd.from_julian(julian).date()
        assert dt == reverse
        self.sd.excel_base_date = CALENDAR_WINDOWS_1900
