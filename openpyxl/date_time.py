from __future__ import absolute_import
from __future__ import division

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

"""Manage Excel date weirdness."""

# Python stdlib imports
from math import floor
import datetime
import re
import warnings

from jdcal import (
    gcal2jd,
    jd2gcal,
    MJD_0
)


# constants
MAC_EPOCH = datetime.datetime(1904, 1, 1)
WINDOWS_EPOCH = datetime.datetime(1899, 12, 30)
CALENDAR_WINDOWS_1900 = sum(gcal2jd(WINDOWS_EPOCH.year, WINDOWS_EPOCH.month, WINDOWS_EPOCH.day))
CALENDAR_MAC_1904 = sum(gcal2jd(MAC_EPOCH.year, MAC_EPOCH.month, MAC_EPOCH.day))
SECS_PER_DAY = 86400

EPOCH = datetime.datetime.utcfromtimestamp(0)
W3CDTF_FORMAT = '%Y-%m-%dT%H:%M:%SZ'
W3CDTF_REGEX = re.compile('(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(.(\d{2}))?Z?')


def datetime_to_W3CDTF(dt):
    """Convert from a datetime to a timestamp string."""
    return datetime.datetime.strftime(dt, W3CDTF_FORMAT)


def W3CDTF_to_datetime(formatted_string):
    """Convert from a timestamp string to a datetime object."""
    match = W3CDTF_REGEX.match(formatted_string)
    dt = [int(v) for v in match.groups()[:6]]
    return datetime.datetime(*dt)


class SharedDate(object):
    """Date formatting utilities for Excel with shared state.

    Excel has a two primary date tracking schemes:
      Windows - Day 1 == 1900-01-01
      Mac - Day 1 == 1904-01-01

    SharedDate stores which system we are using and converts dates between
    Python and Excel accordingly.

    """
    datetime_object_type = 'DateTime'

    def __init__(self,base_date=CALENDAR_WINDOWS_1900):
        warnings.warn("Use module functions directly fo conversion")
        if base_date not in (CALENDAR_MAC_1904, CALENDAR_WINDOWS_1900):
            raise ValueError("base_date:%s invalid" % base_date)
        else:
            self.excel_base_date = base_date

    def datetime_to_julian(self, date):
        """Convert from python datetime to excel julian date representation."""

        if isinstance(date, datetime.datetime):
            return to_excel(date, self.excel_base_date)
        elif isinstance(date, datetime.date):
            return to_excel(date, self.excel_base_date)
        elif isinstance(date, datetime.time):
            return time_to_days(date)
        elif isinstance(date, datetime.timedelta):
            return timedelta_to_days(date)

    def time_to_julian(self, hours, minutes, seconds):
        return ((hours * 3600) + (minutes * 60) + seconds) / SECS_PER_DAY

    def to_julian(self, year, month, day, hours=0, minutes=0, seconds=0):
        """Convert from Python date to Excel JD."""
        # explicitly disallow bad years
        # Excel 2000 treats JD=0 as 1/0/1900 (buggy, disallow)
        # Excel 2000 treats JD=2958466 as a bad date (Y10K bug!)
        if year < 1900 or year > 10000:
            msg = 'Year not supported by Excel: %s' % year
            raise ValueError(msg)
        if self.excel_base_date == CALENDAR_WINDOWS_1900:
            # Fudge factor for the erroneous fact that the year 1900 is
            # treated as a Leap Year in MS Excel.  This affects every date
            # following 28th February 1900
            if year == 1900 and month <= 2:
                excel_1900_leap_year = False
            else:
                excel_1900_leap_year = True
            excel_base_date = 2415020
        elif self.excel_base_date == CALENDAR_MAC_1904:
            excel_base_date = 2416481
            excel_1900_leap_year = False
        else:
            raise NotImplementedError('base date supported.')


        # Julian base date adjustment
        if month > 2:
            month = month - 3
        else:
            month = month + 9
            year -= 1

        # Calculate the Julian Date, then subtract the Excel base date
        # JD 2415020 = 31 - Dec - 1899 -> Excel Date of 0
        century, decade = divmod(year, 100)
        excel_date = floor(146097 * century / 4) + \
                floor((1461 * decade) / 4) + floor((153 * month + 2) / 5) + \
                day + 1721119 - excel_base_date
        if excel_1900_leap_year:
            excel_date += 1

        # check to ensure that we exclude 2/29/1900 as a possible value
        if self.excel_base_date == CALENDAR_WINDOWS_1900 \
                and excel_date == 60:
            msg = 'Error: Excel believes 1900 was a leap year'
            raise ValueError(msg)
        excel_time = self.time_to_julian(hours, minutes, seconds)
        return excel_date + excel_time

    def from_julian(self, value=0):
        return from_excel(value, self.excel_base_date)
        """Convert from the Excel JD back to a date"""


def to_excel(dt, offset=CALENDAR_WINDOWS_1900):
    jul = sum(gcal2jd(dt.year, dt.month, dt.day))
    if hasattr(dt, 'time'):
        # need all calcs at once to avoid rounding errors
        return jul - offset + time_to_days(dt)
    return jul - offset


def from_excel(value, offset=CALENDAR_WINDOWS_1900):
    parts = list(jd2gcal(MJD_0, value + offset - MJD_0))
    fractions = value - int(value)
    diff = datetime.timedelta(days=fractions)
    if 1 > value > 0 or 0 > value > -1:
        return days_to_time(diff)
    return datetime.datetime(*parts[:3]) + diff


def time_to_days(value):
    """Convert a time value to fractions of day"""
    return (
        (value.hour * 3600)
        + (value.minute * 60)
        + value.second
        + value.microsecond / 10**6
        ) / SECS_PER_DAY


def timedelta_to_days(value):
    """Convert a timedelta value to fractions of a day"""
    return value.total_seconds() / SECS_PER_DAY


def days_to_time(value):
    mins, seconds = divmod(value.seconds, 60)
    hours, mins = divmod(mins, 60)
    return datetime.time(hours, mins, seconds, value.microseconds)
