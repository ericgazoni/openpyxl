# Copyright (c) 2010-2013 openpyxl
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

import re

from .hashable import HashableObject


class NumberFormat(HashableObject):
    """Numer formatting for use in styles."""
    FORMAT_GENERAL = 'General'
    FORMAT_TEXT = '@'
    FORMAT_NUMBER = '0'
    FORMAT_NUMBER_00 = '0.00'
    FORMAT_NUMBER_COMMA_SEPARATED1 = '#,##0.00'
    FORMAT_NUMBER_COMMA_SEPARATED2 = '#,##0.00_-'
    FORMAT_PERCENTAGE = '0%'
    FORMAT_PERCENTAGE_00 = '0.00%'
    FORMAT_DATE_YYYYMMDD2 = 'yyyy-mm-dd'
    FORMAT_DATE_YYYYMMDD = 'yy-mm-dd'
    FORMAT_DATE_DDMMYYYY = 'dd/mm/yy'
    FORMAT_DATE_DMYSLASH = 'd/m/y'
    FORMAT_DATE_DMYMINUS = 'd-m-y'
    FORMAT_DATE_DMMINUS = 'd-m'
    FORMAT_DATE_MYMINUS = 'm-y'
    FORMAT_DATE_XLSX14 = 'mm-dd-yy'
    FORMAT_DATE_XLSX15 = 'd-mmm-yy'
    FORMAT_DATE_XLSX16 = 'd-mmm'
    FORMAT_DATE_XLSX17 = 'mmm-yy'
    FORMAT_DATE_XLSX22 = 'm/d/yy h:mm'
    FORMAT_DATE_DATETIME = 'd/m/y h:mm'
    FORMAT_DATE_TIME1 = 'h:mm AM/PM'
    FORMAT_DATE_TIME2 = 'h:mm:ss AM/PM'
    FORMAT_DATE_TIME3 = 'h:mm'
    FORMAT_DATE_TIME4 = 'h:mm:ss'
    FORMAT_DATE_TIME5 = 'mm:ss'
    FORMAT_DATE_TIME6 = 'h:mm:ss'
    FORMAT_DATE_TIME7 = 'i:s.S'
    FORMAT_DATE_TIME8 = 'h:mm:ss@'
    FORMAT_DATE_TIMEDELTA = '[hh]:mm:ss'
    FORMAT_DATE_YYYYMMDDSLASH = 'yy/mm/dd@'
    FORMAT_CURRENCY_USD_SIMPLE = '"$"#,##0.00_-'
    FORMAT_CURRENCY_USD = '$#,##0_-'
    FORMAT_CURRENCY_EUR_SIMPLE = '[$EUR ]#,##0.00_-'
    _BUILTIN_FORMATS = {
        0: 'General',
        1: '0',
        2: '0.00',
        3: '#,##0',
        4: '#,##0.00',
        5: '"$"#,##0_);("$"#,##0)',
        6: '"$"#,##0_);[Red]("$"#,##0)',
        7: '"$"#,##0.00_);("$"#,##0.00)',
        8: '"$"#,##0.00_);[Red]("$"#,##0.00)',
        9: '0%',
        10: '0.00%',
        11: '0.00E+00',
        12: '# ?/?',
        13: '# ??/??',
        14: 'mm-dd-yy',
        15: 'd-mmm-yy',
        16: 'd-mmm',
        17: 'mmm-yy',
        18: 'h:mm AM/PM',
        19: 'h:mm:ss AM/PM',
        20: 'h:mm',
        21: 'h:mm:ss',
        22: 'm/d/yy h:mm',

        37: '#,##0_);(#,##0)',
        38: '#,##0_);[Red](#,##0)',
        39: '#,##0.00_);(#,##0.00)',
        40: '#,##0.00_);[Red](#,##0.00)',

        41: '_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)',
        42: '_("$"* #,##0_);_("$"* \(#,##0\);_("$"* "-"_);_(@_)',
        43: '_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)',

        44: '_("$"* #,##0.00_)_("$"* \(#,##0.00\)_("$"* "-"??_)_(@_)',
        45: 'mm:ss',
        46: '[h]:mm:ss',
        47: 'mmss.0',
        48: '##0.0E+0',
        49: '@', }
    _BUILTIN_FORMATS_REVERSE = dict(
            [(value, key) for key, value in _BUILTIN_FORMATS.items()])

    __fields__ = ('_format_code',
                  '_format_index')
    __slots__ = __fields__
    __leaf__ = True

    DATE_INDICATORS = 'dmyhs'
    BAD_DATE_RE = re.compile(r'(\[|").*[dmhys].*(\]|")')

    def __eq__(self, other):
        if isinstance(other, NumberFormat):
            return self.format_code == other.format_code
        return self.format_code == other

    def __init__(self):
        self._format_code = self.FORMAT_GENERAL
        self._format_index = 0

    @property
    def format_code(self):
        """Getter for the format_code property."""
        return self._format_code

    @format_code.setter
    def format_code(self, format_code = FORMAT_GENERAL):
        """Setter for the format_code property."""
        self._format_code = format_code
        self._format_index = self.builtin_format_id(format_code)

    def builtin_format_code(self, index):
        """Return one of the standard format codes by index."""
        return self._BUILTIN_FORMATS[index]

    def is_builtin(self):
        """Check if a format code is a standard format code."""
        return is_builtin(self.format_code)

    def builtin_format_id(self, fmt):
        """Return the id of a standard style."""
        return self._BUILTIN_FORMATS_REVERSE.get(fmt, None)

    def is_date_format(self):
        """Check if the number format is actually representing a date."""
        return is_date_format(self.format_code)


def is_date_format(fmt):
    if fmt is None:
        return False
    if any([x in fmt for x in NumberFormat.DATE_INDICATORS]):
        return not NumberFormat.BAD_DATE_RE.search(fmt)
    return False


def is_builtin(fmt):
    return fmt in NumberFormat._BUILTIN_FORMATS.values()
