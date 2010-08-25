'''
Copyright (c) 2010 openpyxl

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.

@license: http://www.opensource.org/licenses/mit-license.php
@author: Eric Gazoni
'''
from __future__ import division
from math import floor
import calendar
import datetime
import time

W3CDTF_FORMAT = '%Y-%m-%dT%H:%M:%SZ'

def datetime_to_W3CDTF(dt):

    return datetime.datetime.strftime(dt, W3CDTF_FORMAT)

def W3CDTF_to_datetime(formatted_string):

    return datetime.datetime.strptime(formatted_string, W3CDTF_FORMAT)


class SharedDate(object):

    CALENDAR_WINDOWS_1900 = 1900
    CALENDAR_MAC_1904 = 1904

    datetime_object_type = 'DateTime'

    def __init__(self):

        self.excel_base_date = self.CALENDAR_WINDOWS_1900


    def datetime_to_julian(self, date):

        return self.to_julian(year = date.year,
                              month = date.month,
                              day = date.day,
                              hours = date.hour,
                              minutes = date.minute,
                              seconds = date.second)

    def to_julian(self, year, month, day, hours = 0, minutes = 0, seconds = 0):

        excel_1900_leap_year = None
        excel_base_date = 0

        if self.excel_base_date == self.CALENDAR_WINDOWS_1900:
            #
            #    Fudge factor for the erroneous fact that the year 1900 is treated as a Leap Year in MS Excel
            #    This affects every date following 28th February 1900
            #

            excel_1900_leap_year = True

            if year == 1900 and month <= 2:
                excel_1900_leap_year = False

            excel_base_date = 2415020
        else:

            excel_base_date = 2416481
            excel_1900_leap_year = False


        # Julian base date adjustment
        if month > 2:
            month = month - 3
        else:
            month = month + 9
            year -= 1

        # Calculate the Julian Date, then subtract the Excel base date (JD 2415020 = 31 - Dec - 1899 Giving Excel Date of 0)
        century, decade = int(str(year)[:2]), int(str(year)[2:])

        excel_date = floor(146097 * century / 4) + floor((1461 * decade) / 4) + floor((153 * month + 2) / 5) + day + 1721119 - excel_base_date

        if excel_1900_leap_year:
            excel_date += 1

        excel_time = ((hours * 3600) + (minutes * 60) + seconds) / 86400

        return excel_date + excel_time


    def from_julian(self, value = 0):

        excel_base_date = 0

        if self.excel_base_date == self.CALENDAR_WINDOWS_1900:

            excel_base_date = 25569

            if value < 60:
                excel_base_date -= 1

        else:
            excel_base_date = 24107


        if value >= 1:
            utc_days = value - excel_base_date

            seconds = round(utc_days * 24 * 60 * 60)

            return datetime.datetime.utcfromtimestamp(calendar.timegm(time.gmtime(seconds)))

        elif value >= 0:

            hours = floor(value * 24)
            mins = floor(value * 24 * 60) - floor(hours * 60)
            secs = floor(value * 24 * 60 * 60) - floor(hours * 60 * 60) - floor(mins * 60)

            return datetime.time(int(hours), int(mins), int(secs))

        else:

            raise ValueError('Negative dates (%s) are not supported on this platform' % value)
