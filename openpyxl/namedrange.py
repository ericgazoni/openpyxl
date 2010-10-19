# file openpyxl/namedrange.py

# Copyright (c) 2010 openpyxl
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
# @author: Eric Gazoni

"""Track named groups of cells in a worksheet"""

# Python stdlib imports
import re

# package imports
from openpyxl.shared.exc import NamedRangeException

# constants
NAMED_RANGE_RE = re.compile("'?([^']*)'?!\$([A-Za-z]+)\$([0-9]+)")


class NamedRange(object):
    """A named group of cells"""
    __slots__ = ('name', 'worksheet', 'range', 'local_only')

    def __init__(self, name, worksheet, range):
        self.name = name
        self.worksheet = worksheet
        self.range = range
        self.local_only = False

    def __str__(self):
        return u'%s!%s' % (self.worksheet.title, self.range)


def split_named_range(range_string):
    """Separate a named range into its component parts"""
    match = NAMED_RANGE_RE.match(range_string)
    if not match:
        raise NamedRangeException('Invalid named range string: "%s"')
    else:
        sheet_name, column, row = match.groups()
        return sheet_name, column, int(row)
