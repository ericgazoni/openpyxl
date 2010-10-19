# file openpyxl/tests/test_named_range.py

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

# Python stdlib imports
from __future__ import with_statement
import os.path

# 3rd-party imports
from nose.tools import eq_, assert_raises

# package imports
from openpyxl.tests.helper import DATADIR
from openpyxl.namedrange import split_named_range
from openpyxl.reader.workbook import read_named_ranges
from openpyxl.shared.exc import NamedRangeException


def test_split():
    eq_(('My Sheet', 'D', 8), split_named_range("'My Sheet'!$D$8"))


def test_split_no_quotes():
    eq_(('HYPOTHESES', 'B', 3), split_named_range('HYPOTHESES!$B$3:$L$3'))


def test_bad_range_name():
    assert_raises(NamedRangeException, split_named_range, 'HYPOTHESES$B$3')


def test_read_named_ranges():

    class DummyWs(object):
        title = 'My Sheeet'

    class DummyWB(object):

        def get_sheet_by_name(self, name):
            return DummyWs()

    with open(os.path.join(DATADIR, 'reader', 'workbook.xml')) as handle:
        content = handle.read()
    named_ranges = read_named_ranges(content, DummyWB())
    eq_(["My Sheeet!D8"], [str(range) for range in named_ranges])
