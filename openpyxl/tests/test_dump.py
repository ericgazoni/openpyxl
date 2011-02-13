
# file openpyxl/tests/test_dump.py

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
from datetime import time, datetime

# 3rd party imports
from nose.tools import eq_, raises, assert_raises

from openpyxl.workbook import Workbook
from openpyxl.cell import get_column_letter

def test_dump_sheet():

    wb = Workbook(optimized_write = True)

    ws = wb.create_sheet()

    letters = [get_column_letter(x+1) for x in xrange(20)]

    for row in xrange(20):

        ws.append(['%s%d' % (letter, row+1) for letter in letters])

    wb.save('test.xlsx')
