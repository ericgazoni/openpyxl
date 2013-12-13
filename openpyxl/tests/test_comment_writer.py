# file openpyxl/tests/test_comment_reader.py

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

import pytest

import os

from openpyxl.workbook import Workbook
from openpyxl.worksheet import Worksheet
from openpyxl.writer.comments import CommentWriter
from openpyxl.comments import Comment
from openpyxl.tests.helper import DATADIR, compare_xml

def _create_ws():
    wb = Workbook()
    ws = Worksheet(wb)
    comment1 = Comment(ws.cell(coordinate="B2"), "text", "author")
    comment2 = Comment(ws.cell(coordinate="C7"), "text2", "author2")
    comment3 = Comment(ws.cell(coordinate="D9"), "text3", "author3")
    ws.cell(coordinate="B2").comment = comment1
    ws.cell(coordinate="C7").comment = comment2
    ws.cell(coordinate="D9").comment = comment3
    return ws, comment1, comment2, comment3

def test_comment_writer_init():
    ws, comment1, comment2, comment3 = _create_ws()
    cw = CommentWriter(ws)
    assert set(cw.authors) == set(["author", "author2", "author3"])
    assert cw.author_to_id[cw.authors[0]] == "0"
    assert cw.author_to_id[cw.authors[1]] == "1"
    assert cw.author_to_id[cw.authors[2]] == "2"
    assert set(cw.comments) == set([comment1, comment2, comment3])

def test_write_comments():
    ws = _create_ws()[0]

    reference_file = os.path.join(DATADIR, 'writer', 'expected',
            'comments1.xml')
    cw = CommentWriter(ws)
    content = cw.write_comments()
    with open(reference_file) as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff


def test_write_comments_vml():
    ws = _create_ws()[0]
    cw = CommentWriter(ws)
    reference_file = os.path.join(DATADIR, 'writer', 'expected',
            'commentsDrawing1.vml')
    content = cw.write_comments_vml()
    with open(reference_file) as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff
