# Copyright (c) 2010-2014 openpyxl
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

from zipfile import ZipFile, ZIP_DEFLATED
import os.path
from collections import namedtuple

from openpyxl.workbook import Workbook
from openpyxl.worksheet import Worksheet
from openpyxl.comments import Comment
from openpyxl.reader import comments
from openpyxl.reader.excel import load_workbook
from openpyxl.shared.xmltools import fromstring
from openpyxl.tests.helper import DATADIR

def test_get_author_list():
    xml = """<?xml version="1.0" standalone="yes"?><comments
    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors>
    <author>Cuke</author><author>Not Cuke</author></authors><commentList>
    </commentList></comments>"""
    assert comments._get_author_list(fromstring(xml)) == ['Cuke', 'Not Cuke']

def test_read_comments():
    xml = """<?xml version="1.0" standalone="yes"?>
    <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors>
    <author>Cuke</author><author>Not Cuke</author></authors><commentList><comment ref="A1"
    authorId="0" shapeId="0"><text><r><rPr><b/><sz val="9"/><color indexed="81"/><rFont
    val="Tahoma"/><charset val="1"/></rPr><t>Cuke:\n</t></r><r><rPr><sz val="9"/><color
    indexed="81"/><rFont val="Tahoma"/><charset val="1"/></rPr>
    <t xml:space="preserve">First Comment</t></r></text></comment><comment ref="D1" authorId="0" shapeId="0">
    <text><r><rPr><b/><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><charset val="1"/>
    </rPr><t>Cuke:\n</t></r><r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/>
    <charset val="1"/></rPr><t xml:space="preserve">Second Comment</t></r></text></comment>
    <comment ref="A2" authorId="1" shapeId="0"><text><r><rPr><b/><sz val="9"/><color
    indexed="81"/><rFont val="Tahoma"/><charset val="1"/></rPr><t>Not Cuke:\n</t></r><r><rPr>
    <sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><charset val="1"/></rPr>
    <t xml:space="preserve">Third Comment</t></r></text></comment></commentList></comments>"""
    wb = Workbook()
    ws = Worksheet(wb)
    comments.read_comments(ws, xml)
    comments_expected = [['A1', 'Cuke', 'Cuke:\nFirst Comment'],
                         ['D1', 'Cuke', 'Cuke:\nSecond Comment'],
                         ['A2', 'Not Cuke', 'Not Cuke:\nThird Comment']
                        ]
    for cell, author, text in comments_expected:
        assert ws.cell(coordinate=cell).comment.author == author
        assert ws.cell(coordinate=cell).comment.text == text
        assert ws.cell(coordinate=cell).comment._parent == ws.cell(coordinate=cell)

def test_get_comments_file():
    path = os.path.join(DATADIR, 'reader', 'comments.xlsx')
    archive = ZipFile(path, 'r', ZIP_DEFLATED)
    valid_files = archive.namelist()
    assert comments.get_comments_file('sheet1.xml', archive, valid_files) == 'xl/comments1.xml'
    assert comments.get_comments_file('sheet3.xml', archive, valid_files) == 'xl/comments2.xml'
    assert comments.get_comments_file('sheet2.xml', archive, valid_files) is None

def test_comments_cell_association():
    path = os.path.join(DATADIR, 'reader', 'comments.xlsx')
    wb = load_workbook(path)
    assert wb.worksheets[0].cell(coordinate="A1").comment.author == "Cuke"
    assert wb.worksheets[0].cell(coordinate="A1").comment.text == "Cuke:\nFirst Comment"
    assert wb.worksheets[1].cell(coordinate="A1").comment is None
    assert wb.worksheets[0].cell(coordinate="D1").comment.text == "Cuke:\nSecond Comment"

def test_comments_with_iterators():
    wb = load_workbook(filename=os.path.join(DATADIR, 'reader', 'comments.xlsx'), use_iterators=True)
    ws = wb.get_sheet_by_name(name='Sheet1')
