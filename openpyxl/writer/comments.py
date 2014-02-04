from __future__ import absolute_import
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

from openpyxl.compat import iteritems
from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.xml.functions import Element, SubElement, get_document_content
from openpyxl.cell import column_index_from_string

vmlns="urn:schemas-microsoft-com:vml"
officens="urn:schemas-microsoft-com:office:office"
excelns="urn:schemas-microsoft-com:office:excel"

class CommentWriter(object):
    def __init__(self, sheet):
        self.sheet = sheet

        # get list of comments
        self.comments = []
        for coord, cell in iteritems(sheet._cells):
            if cell.comment is not None:
                self.comments.append(cell.comment)

        # get list of authors
        self.authors = []
        self.author_to_id = {}
        for comment in self.comments:
            if comment.author not in self.author_to_id:
                self.author_to_id[comment.author] = str(len(self.authors))
                self.authors.append(comment.author)

    def write_comments(self):
        # produce xml
        root = Element("{%s}comments" % SHEET_MAIN_NS)
        authorlist_tag = SubElement(root, "{%s}authors" % SHEET_MAIN_NS)
        for author in self.authors:
            leaf = SubElement(authorlist_tag, "{%s}author" % SHEET_MAIN_NS)
            leaf.text = author

        commentlist_tag = SubElement(root, "{%s}commentList" % SHEET_MAIN_NS)
        for comment in self.comments:
            attrs = {'ref': comment._parent.get_coordinate(),
                     'authorId': self.author_to_id[comment.author],
                     'shapeId': '0'}
            comment_tag = SubElement(commentlist_tag, "{%s}comment" % SHEET_MAIN_NS, attrs)

            text_tag = SubElement(comment_tag, "{%s}text" % SHEET_MAIN_NS)
            run_tag = SubElement(text_tag, "{%s}r" % SHEET_MAIN_NS)
            SubElement(run_tag, "{%s}rPr" % SHEET_MAIN_NS)
            t_tag = SubElement(run_tag, "{%s}t" % SHEET_MAIN_NS)
            t_tag.text = comment.text

        return get_document_content(root)

    def write_comments_vml(self):
        root = Element("xml")
        shape_layout = SubElement(root, "{%s}shapelayout" % officens, {"{%s}ext" % vmlns: "edit"})
        SubElement(shape_layout, "{%s}idmap" % officens, {"{%s}ext" % vmlns: "edit", "data": "1"})
        shape_type=SubElement(root, "{%s}shapetype" % vmlns, {"id": "_x0000_t202",
                                                              "coordsize": "21600,21600",
                                                              "{%s}spt" % officens: "202",
                                                              "path": "m,l,21600r21600,l21600,xe"})
        SubElement(shape_type, "{%s}stroke" % vmlns, {"joinstyle": "miter"})
        SubElement(shape_type, "{%s}path" % vmlns, {"gradientshapeok": "t",
                                                    "{%s}connecttype" % officens: "rect"})

        for i, comment in enumerate(self.comments):
            self._write_comment_shape(root, comment, i)

        return get_document_content(root)

    def _write_comment_shape(self, root, comment, idx):
        # get zero-indexed coordinates of the comment
        row = comment._parent.row - 1
        column = column_index_from_string(comment._parent.column) - 1

        attrs = {
            "id": "_x0000_s%s" % (idx+1026),
            "type": "#_x0000_t202",
            "style": "position:absolute; margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden",
            "fillcolor": "#ffffe1",
            "{%s}insetmode" % officens: "auto"
        }
        shape = SubElement(root, "{%s}shape" % vmlns, attrs)

        SubElement(shape, "{%s}fill" % vmlns, {"color2":"#ffffe1"})
        SubElement(shape, "{%s}shadow" % vmlns, {"color":"black", "obscured":"t"})
        SubElement(shape, "{%s}path" % vmlns, {"{%s}connecttype"%officens:"none"})
        textbox = SubElement(shape, "{%s}textbox" % vmlns, {"style":"mso-direction-alt:auto"})
        SubElement(textbox, "div", {"style": "text-align:left"})
        client_data = SubElement(shape, "{%s}ClientData" % excelns, {"ObjectType": "Note"})
        SubElement(client_data, "{%s}MoveWithCells" % excelns)
        SubElement(client_data, "{%s}SizeWithCells" % excelns)
        SubElement(client_data, "{%s}AutoFill" % excelns).text = "False"
        SubElement(client_data, "{%s}Row" % excelns).text = str(row)
        SubElement(client_data, "{%s}Column" % excelns).text = str(column)
