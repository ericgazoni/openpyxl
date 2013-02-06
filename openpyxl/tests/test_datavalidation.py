# file openpyxl/tests/test_cell.py

# Copyright (c) 2010-2011 openpyxl
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

# 3rd party imports
from nose.tools import eq_, raises, assert_raises #pylint: disable=E0611

# package imports
from openpyxl.datavalidation import collapse_cell_addresses

# There are already unit-tests in test_cell.py that test out the
# coordinate_from_string method.  This should be the only way the
# collapse_cell_addresses method can throw, so we don't bother using invalid
# cell coordinates in the test-data here.
COLLAPSE_TEST_DATA = [
    (["A1"], "A1"),
    (["A1", "B1"], "A1 B1"),
    (["A1", "A2", "A3", "A4", "B1", "B2", "B3", "B4"], "A1:A4 B1:B4"),
    (["A2", "A4", "A3", "A1", "A5"], "A1:A5"),
]


def test_collapse_cell_addresses():
    for data in COLLAPSE_TEST_DATA:
        eq_(collapse_cell_addresses(data[0]), data[1])
