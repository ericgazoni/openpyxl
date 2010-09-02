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
import os
import os.path as osp
import shutil
import unittest
import difflib
from StringIO import StringIO
from openpyxl.shared.xmltools import fromstring, ElementTree
from pprint import pprint
from openpyxl.shared.xmltools import pretty_indent
from tempfile import gettempdir

DATADIR = osp.abspath(osp.join(osp.dirname(__file__), 'test_data'))
TMPDIR = osp.join(gettempdir(), 'openpyxl_test_temp')

def clean_tmpdir():
    if osp.isdir(TMPDIR):
        shutil.rmtree(TMPDIR, ignore_errors = True)
    os.makedirs(TMPDIR)

clean_tmpdir()

class BaseTestCase(unittest.TestCase):

    def assertEqualsFileContent(self, reference_file, fixture, filetype = 'xml'):

        fixture_content = fixture
        with open(reference_file) as expected_file:
            expected_content = expected_file.read()

        if filetype == 'xml':

            fixture_content = fromstring(fixture_content)
            pretty_indent(fixture_content)
            temp = StringIO()
            ElementTree(fixture_content).write(temp)
            fixture_content = temp.getvalue()

            expected_content = fromstring(expected_content)
            pretty_indent(expected_content)
            temp = StringIO()
            ElementTree(expected_content).write(temp)
            expected_content = temp.getvalue()

        fixture_lines = fixture_content.split('\n')
        expected_lines = expected_content.split('\n')

        differences = list(difflib.unified_diff(expected_lines, fixture_lines))

        if differences:
            temp = StringIO()
            pprint(differences, stream = temp)
            self.fail('Differences found : %s' % temp.getvalue())


    def tearDown(self):

        self.clean_tmpdir()

    def clean_tmpdir(self):

        clean_tmpdir()

