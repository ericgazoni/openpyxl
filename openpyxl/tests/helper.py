# file openpyxl/tests/helper.py

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
import os
import sys
import os.path
import shutil
import difflib
from pprint import pprint
from tempfile import gettempdir
from sys import version_info
from lxml.doctestcompare import LXMLOutputChecker, PARSE_XML

# package imports
from openpyxl.shared.compat import BytesIO, unicode, StringIO
from openpyxl.shared.xmltools import fromstring, ElementTree
from openpyxl.shared.xmltools import pretty_indent


# constants
DATADIR = os.path.abspath(os.path.join(os.path.dirname(__file__), 'test_data'))
TMPDIR = os.path.join(gettempdir(), 'openpyxl_test_temp')


def make_tmpdir():
    try:
        os.makedirs(TMPDIR)
    except OSError:
        pass


def clean_tmpdir():
    if os.path.isdir(TMPDIR):
        shutil.rmtree(TMPDIR, ignore_errors = True)


def get_xml(xml_node):

    io = BytesIO()
    if version_info[0] >= 3 and version_info[1] >= 2:
        ElementTree(xml_node).write(io, encoding='UTF-8', xml_declaration=False)
        ret = str(io.getvalue(), 'utf-8')
        ret = ret.replace('utf-8', 'UTF-8', 1)
    else:
        ElementTree(xml_node).write(io, encoding='UTF-8')
        ret = io.getvalue()
    io.close()
    return ret.replace('\n', '')


def compare_xml(generated, expected):
    """Use doctest checking from lxml for comparing XML trees. Returns diff if the two are not the same"""
    checker = LXMLOutputChecker()

    class DummyDocTest():
        pass

    ob = DummyDocTest()
    ob.want = generated

    check = checker.check_output(expected, generated, PARSE_XML)
    if check is False:
        diff = checker.output_difference(ob, expected, PARSE_XML)
        return diff


def safe_iterator(node):
    """Return an iterator that is compatible with Python 2.6"""
    if hasattr(node, "iter"):
        return node.iter()
    else:
        return node.getiterator()
