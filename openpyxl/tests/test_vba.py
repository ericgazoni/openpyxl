# file openpyxl/tests/test_theme.py

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

# Python stdlib imports
import os.path, zipfile, StringIO

# package imports
from openpyxl.tests.helper import DATADIR
from openpyxl.reader.excel import load_workbook
from openpyxl.writer.excel import save_virtual_workbook

def test_save_vba():
    path = os.path.join(DATADIR, 'reader', 'vba-test.xlsm')
    wb = load_workbook(path, keep_vba=True)
    buf = save_virtual_workbook(wb)
    files1 = set(zipfile.ZipFile(path, 'r').namelist())
    files2 = set(zipfile.ZipFile(StringIO.StringIO(buf), 'r').namelist())
    assert files1.issubset(files2), "Missing files: %s" % ', '.join(files1 - files2)
