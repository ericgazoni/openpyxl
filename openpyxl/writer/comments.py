# coding=UTF-8
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


from openpyxl.shared.ooxml import COMMENTS_NS, REL_NS, PKG_REL_NS, SHEET_MAIN_NS
from openpyxl.shared.xmltools import Element, SubElement, get_document_content

def write_comments(sheet):
	# get list of comments
	comments = []
	for coord, cell in sheet._cells.iteritems():
		if cell.comment is not None:
			comments.append(cell.comment)

	# get list of authors
	authors = []
	author_to_id = {}
	for comment in comments:
		if comment.author not in author_to_id:
			author_to_id[comment.author] = str(len(authors))
			authors.append(comment.author)

	# produce xml
	root = Element("{%s}comments" % SHEET_MAIN_NS)
	authorlist_tag = SubElement(root, "{%s}authors" % SHEET_MAIN_NS)
	for author in authors:
		leaf = SubElement(authorlist_tag, "{%s}author" % SHEET_MAIN_NS)
		leaf.text = author

	commentlist_tag = SubElement(root, "{%s}commentList" % SHEET_MAIN_NS)
	for comment in comments:
		attrs = {'ref': comment.parent.get_coordinate(),
				 'authorId': author_to_id[comment.author]}
		comment_tag = SubElement(commentlist_tag, "{%s}comment" % SHEET_MAIN_NS, attrs)

		text_tag = SubElement(comment_tag, "{%s}text" % SHEET_MAIN_NS)
		SubElement(text_tag, "{%s}rPr" % SHEET_MAIN_NS)
		run_tag = SubElement(text_tag, "{%s}r" % SHEET_MAIN_NS)
		run_tag.text = comment.text.replace("\n", '')

	return get_document_content(root)

