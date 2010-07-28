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

import os.path as osp

from zipfile import ZipFile, ZIP_DEFLATED

from openpyxl.tests.helper import BaseTestCase, TMPDIR
from openpyxl.shared.zip import ZipArchive

class TestZip(BaseTestCase):

    def test_write_zip(self):

        filename = osp.join(TMPDIR, 'test.zip')

        inner_filename = 'file.a'
        inner_content = "here is the content"


        z = ZipArchive(filename = filename, mode = 'w')

        z.add_from_string(inner_filename, inner_content)

        z.close()


        test_zip = ZipFile(file = filename,
                           mode = 'r',
                           compression = ZIP_DEFLATED,
                           allowZip64 = True)

        self.assertTrue(inner_filename in test_zip.namelist())

        self.assertEqual(test_zip.read(inner_filename), inner_content)

        test_zip.close()

    def test_read_zip(self):

        filename = osp.join(TMPDIR, 'test.zip')

        inner_filename = 'file.a'
        inner_content = "here is the content"

        # write the zip file
        z = ZipArchive(filename = filename, mode = 'w')
        z.add_from_string(inner_filename, inner_content)
        z.close()

        # read it again
        z = ZipArchive(filename = filename)
        read_content = z.get_from_name(inner_filename)
        z.close()

        self.assertEqual(read_content, inner_content)
