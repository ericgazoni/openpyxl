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
from __future__ import with_statement
import os.path as osp
from openpyxl.tests.helper import BaseTestCase, DATADIR

from openpyxl.namedrange import split_named_range
from openpyxl.reader.workbook import read_named_ranges
from openpyxl.shared.zip import ZipArchive
from openpyxl.shared.ooxml import ARC_WORKBOOK

class TestNamedRanges(BaseTestCase):

    def test_split(self):

        self.assertEqual(('My Sheet', 'D', 8),
                         split_named_range(range_string = "'My Sheet'!$D$8"))

    def test_split_no_quotes(self):

        self.assertEqual(('HYPOTHESES', 'B', 3),
                         split_named_range(range_string = 'HYPOTHESES!$B$3:$L$3'))

class TestReadNamedRanges(BaseTestCase):

    def test_read_named_ranges(self):

        class DummyWs(object):

            title = 'My Sheeet'

        class DummyWB(object):

            def get_sheet_by_name(self, name):
                return DummyWs()

        with open(osp.join(DATADIR, 'reader', 'workbook.xml')) as f:
            content = f.read()

            named_ranges = read_named_ranges(xml_source = content,
                                             workbook = DummyWB())

            self.assertEqual(["My Sheeet!D8"], map(str, named_ranges))

