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
from datetime import datetime, date, timedelta, time

import pytest

# package imports
from openpyxl.date_time import CALENDAR_MAC_1904, CALENDAR_WINDOWS_1900


def test_datetime_to_W3CDTF():
    from openpyxl.date_time import datetime_to_W3CDTF
    assert datetime_to_W3CDTF(datetime(2013, 7, 15, 6, 52, 33)) == "2013-07-15T06:52:33Z"


def test_W3CDTF_to_datetime():
    from openpyxl.date_time import W3CDTF_to_datetime
    value = "2011-06-30T13:35:26Z"
    assert W3CDTF_to_datetime(value) == datetime(2011, 6, 30, 13, 35, 26)
    value = "2013-03-04T12:19:01.00Z"
    assert W3CDTF_to_datetime(value) == datetime(2013, 3, 4, 12, 19, 1)


@pytest.fixture
def sd():
    from openpyxl.date_time import SharedDate
    return SharedDate()


def test_convert_date_to_julian(sd):
    assert 40167 == sd.to_julian(2009, 12, 20)

@pytest.mark.parametrize("value, expected",
                         [
                             (40167, datetime(2009, 12, 20)),
                             (21980, datetime(1960,  3,  5)),
                             (40196.5939815, datetime(2010, 1, 18, 14, 15, 20, 1600)),
                         ])
def test_convert_date_from_julian(sd, value, expected):
    assert sd.from_julian(value) == expected

def test_convert_datetime_to_julian(sd):
    assert 40167 == sd.datetime_to_julian(datetime(2009, 12, 20))
    assert 40196.5939815 == sd.datetime_to_julian(datetime(2010, 1, 18, 14, 15, 20, 1600))

def test_convert_timedelta_to_julian(sd):
    assert 1.125 == sd.datetime_to_julian(timedelta(days=1, hours=3))


def test_convert_time_to_julian(sd):
    t = time(3, 0, 0)
    assert sd.time_to_julian(t.hour, t.minute, t.second) == 0.125


def test_1900_leap_year(sd):
    with pytest.raises(ValueError):
        sd.to_julian(1900, 2, 29)

bad_dates = (
    (1776,  7,  4),
    (1899, 12, 31),
)
@pytest.mark.parametrize("dt", bad_dates)
def test_bad_date(sd, dt):
    with pytest.raises(ValueError):
        sd.to_julian(*dt)


def test_mac_to_date(sd):
    sd.excel_base_date = CALENDAR_MAC_1904
    assert  sd.to_julian(2011, 10, 3) == 39385


def test_mac_to_date(sd):
    sd.excel_base_date = CALENDAR_MAC_1904
    assert sd.from_julian(39385) == datetime(2011, 10, 31)


@pytest.mark.parametrize("value, expected",
                         [
                             (date(1899, 12, 31), 1),
                             (datetime(2010, 1, 18, 14, 15, 20, 1600), 40196.5939815),
                             (date(2009, 12, 20), 40167),
                             (datetime(1506, 10, 15), -143617.0)
                         ])
def test_to_excel(value, expected):
    from openpyxl.date_time import to_excel
    FUT = to_excel
    assert FUT(value) == expected


@pytest.mark.parametrize("value, expected",
                         [
                             (date(1904, 1, 1), 0),
                             (date(2011, 10, 31), 39385),
                             (datetime(2010, 1, 18, 14, 15, 20, 1600), 38734.5939815),
                             (date(2009, 12, 20), 38705),
                             (datetime(1506, 10, 15), -145079.0)
                         ])
def test_to_excel_mac(value, expected):
    from openpyxl.date_time import to_excel
    FUT = to_excel
    assert FUT(value, CALENDAR_MAC_1904) == expected


@pytest.mark.parametrize("value, expected",
                         [
                             (40167, datetime(2009, 12, 20)),
                             (21980, datetime(1960,  3,  5)),
                             (60, datetime(1900, 2, 28)),
                             (-25063, datetime(1831, 5, 18, 0, 0)),
                             (40372.27616898148, datetime(2010, 7, 13, 6, 37, 41)),
                             (40196.5939815, datetime(2010, 1, 18, 14, 15, 20, 1600)),
                         ])
def test_from_excel(value, expected):
    from openpyxl.date_time import from_excel
    FUT = from_excel
    assert FUT(value) == expected


@pytest.mark.parametrize("value, expected",
                         [
                             (39385, datetime(2011, 10, 31)),
                             (21980, datetime(1964,  3,  6)),
                             (0, datetime(1904, 1, 1)),
                             (-25063, datetime(1835, 5, 19))
                         ])
def test_from_excel_mac(value, expected):
    from openpyxl.date_time import from_excel
    FUT = from_excel
    assert FUT(value, CALENDAR_MAC_1904) == expected



def test_time_to_days():
    from openpyxl.date_time import time_to_days
    FUT = time_to_days
    t1 = time(13, 55, 12, 36)
    assert FUT(t1) == 0.5800000004166667
    t2 = time(3, 0, 0)
    assert FUT(t2) == 0.125


def test_timedelta_to_days():
    from openpyxl.date_time import timedelta_to_days
    FUT = timedelta_to_days
    td = timedelta(days=1, hours=3)
    assert FUT(td) == 1.125
