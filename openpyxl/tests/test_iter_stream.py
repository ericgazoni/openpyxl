# file openpyxl/tests/test_iter_stream.py

# Copyright (c) 2011 openpyxl
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

from nose.tools import eq_, raises, assert_raises
import os.path as osp
from openpyxl.tests.helper import DATADIR
from openpyxl.reader.iter_worksheet import get_range_boundaries
from openpyxl.reader.excel import load_workbook
import openpyxl.tests.test_iter as test_iter
import datetime

class StreamTestWorksheet(object):
    workbook_name = osp.join(DATADIR, 'genuine', 'empty_no_dimensions.xlsx')

    def _open_wb(self):
        ff = open(self.workbook_name, 'rb')
        return load_workbook(filename = ff, use_iterators = True)

class TestDims(StreamTestWorksheet, test_iter.TestDims):
    pass

class TestText(StreamTestWorksheet, test_iter.TestText):
    def test_get_boundaries_range(self):
        pass

    def test_get_boundaries_one(self):
        pass

class TestIntegers(StreamTestWorksheet, test_iter.TestIntegers):
    workbook_name = osp.join(DATADIR, 'genuine', 'empty_no_dimensions.xlsx')

class TestFloats(StreamTestWorksheet, test_iter.TestFloats):
    workbook_name = osp.join(DATADIR, 'genuine', 'empty_no_dimensions.xlsx')

class TestDates(StreamTestWorksheet, test_iter.TestDates):
    workbook_name = osp.join(DATADIR, 'genuine', 'empty_no_dimensions.xlsx')

