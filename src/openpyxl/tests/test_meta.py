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
from openpyxl.tests.helper import BaseTestCase, DATADIR, TMPDIR

from openpyxl.writer.workbook import write_content_types, write_root_rels

from openpyxl.workbook import Workbook

class TestWriteMeta(BaseTestCase):

    def test_write_content_types(self):

        wb = Workbook()

        wb.create_sheet()

        wb.create_sheet()

        content = write_content_types(wb)

        reference_file = osp.join(DATADIR, 'writer', 'expected', '[Content_Types].xml')

        self.assertEqualsFileContent(reference_file, fixture = content)

    def test_write_root_rels(self):

        wb = Workbook()

        content = write_root_rels(wb)

        reference_file = osp.join(DATADIR, 'writer', 'expected', '.rels')

        self.assertEqualsFileContent(reference_file, fixture = content)
