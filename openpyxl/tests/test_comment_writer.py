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
from openpyxl.shared.xmltools import fromstring, get_document_content
from openpyxl.shared.ooxml import SHEET_MAIN_NS
from openpyxl.writer.comments import vmlns, excelns, officens

def _create_ws():
    wb = Workbook()
    ws = Worksheet(wb)
    comment1 = Comment("text", "author")
    comment2 = Comment("text2", "author2")
    comment3 = Comment("text3", "author3")
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
        correct = fromstring(expected.read())
        check = fromstring(content)
        # check top-level elements have the same name
        for i, j in zip(correct.getchildren(), check.getchildren()):
            assert i.tag == j.tag

        correct_comments = correct.find('{%s}commentList' % SHEET_MAIN_NS).getchildren()
        check_comments = check.find('{%s}commentList' % SHEET_MAIN_NS).getchildren()
        correct_authors = correct.find('{%s}authors' % SHEET_MAIN_NS).getchildren()
        check_authors = check.find('{%s}authors' % SHEET_MAIN_NS).getchildren()

        # replace author ids with author names
        for i in correct_comments:
            i.attrib["authorId"] = correct_authors[int(i.attrib["authorId"])].text
        for i in check_comments:
            i.attrib["authorId"] = check_authors[int(i.attrib["authorId"])].text

        # sort the comment list
        correct_comments.sort(key=lambda tag: tag.attrib["ref"])
        check_comments.sort(key=lambda tag: tag.attrib["ref"])
        correct.find('{%s}commentList' % SHEET_MAIN_NS)[:] = correct_comments
        check.find('{%s}commentList' % SHEET_MAIN_NS)[:] = check_comments

        # sort the author list
        correct_authors.sort(key=lambda tag: tag.text)
        check_authors.sort(key=lambda tag:tag.text)
        correct.find('{%s}authors' % SHEET_MAIN_NS)[:] = correct_authors
        check.find('{%s}authors' % SHEET_MAIN_NS)[:] = check_authors
        
        diff = compare_xml(get_document_content(correct), get_document_content(check))
        assert diff is None, diff

def test_write_comments_vml():
    ws = _create_ws()[0]
    cw = CommentWriter(ws)
    reference_file = os.path.join(DATADIR, 'writer', 'expected',
            'commentsDrawing1.vml')
    content = cw.write_comments_vml()
    with open(reference_file) as expected:
        correct = fromstring(expected.read())
        check = fromstring(content)
        correct_ids = []
        correct_coords = []
        check_ids = []
        check_coords = []

        for i in correct.findall("{%s}shape" % vmlns):
            correct_ids.append(i.attrib["id"])
            row = i.find("{%s}ClientData" % excelns).find("{%s}Row" % excelns).text
            col = i.find("{%s}ClientData" % excelns).find("{%s}Column" % excelns).text
            correct_coords.append((row,col))
            # blank the data we are checking separately
            i.attrib["id"] = "0"
            i.find("{%s}ClientData" % excelns).find("{%s}Row" % excelns).text="0"
            i.find("{%s}ClientData" % excelns).find("{%s}Column" % excelns).text="0"

        for i in check.findall("{%s}shape" % vmlns):
            check_ids.append(i.attrib["id"])
            row = i.find("{%s}ClientData" % excelns).find("{%s}Row" % excelns).text
            col = i.find("{%s}ClientData" % excelns).find("{%s}Column" % excelns).text
            check_coords.append((row,col))
            # blank the data we are checking separately
            i.attrib["id"] = "0"
            i.find("{%s}ClientData" % excelns).find("{%s}Row" % excelns).text="0"
            i.find("{%s}ClientData" % excelns).find("{%s}Column" % excelns).text="0"

        assert set(correct_coords) == set(check_coords)
        assert set(correct_ids) == set(check_ids)
        diff = compare_xml(get_document_content(correct), get_document_content(check))
        assert diff is None, diff
