# file openpyxl/tests/test_comments.py

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
import os

from nose.tools import eq_

from openpyxl import load_workbook
from openpyxl.tests.helper import DATADIR

def test_many_comments():
    path = os.path.join(DATADIR, 'reader', 'comments.xlsx')
    wb = load_workbook(path)
    ws = wb.worksheets[0]
    firstcomment = ws.cell(coordinate="A1").comment
    eq_(firstcomment.author, "Cuke")
    eq_(firstcomment.text,"Cuke:\nFirst Comment")
    secondcomment = ws.cell(coordinate="D1").comment
    eq_(secondcomment.author, "Cuke")
    eq_(secondcomment.text, "Cuke:\nSecond Comment")
    thirdcomment = ws.cell(coordinate="A2").comment
    eq_(thirdcomment.author, "Cuke")
    eq_(thirdcomment.text, "Cuke:\nThird Comment")
    fourthcomment = ws.cell(coordinate="C7").comment
    eq_(fourthcomment.author, "Cuke")
    eq_(fourthcomment.text, "Cuke:\nFourth Comment")
    differentcomment = ws.cell(coordinate="B9").comment
    eq_(differentcomment.author, "Not Cuke")
    eq_(differentcomment.text, "Not Cuke:\nBecause it has a different author")

def test_comment_absent():
    path = os.path.join(DATADIR, 'reader', 'comments.xlsx')
    wb = load_workbook(path)
    ws = wb.worksheets[1]
    nocomment = ws.cell(coordinate="A1").comment
    eq_(nocomment, None)

def test_singlerun_comment():
    path = os.path.join(DATADIR, 'reader', 'comments.xlsx')
    wb = load_workbook(path)
    ws = wb.worksheets[2]
    singlerun = ws.cell(coordinate="A1").comment
    eq_(singlerun.author, "Not Cuke")
    eq_(singlerun.text, "comment has one run")

