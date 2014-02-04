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

"""Read the shared strings table."""

# package imports
from openpyxl.xml.functions import fromstring
from openpyxl.xml.constants import SHEET_MAIN_NS, XML_NS
from openpyxl.compat import unicode


def read_string_table(xml_source):
    """Read in all shared strings in the table"""
    table = {}
    root = fromstring(text=xml_source)
    string_index_nodes = root.findall('{%s}si' % SHEET_MAIN_NS)
    for index, string_index_node in enumerate(string_index_nodes):

        string = get_string(string_index_node)

        # fix XML escaping sequence for '_x'
        string = string.replace('x005F_', '')

        table[index] = string

    return table


def get_string(string_index_node):
    """Read the contents of a specific string index"""
    rich_nodes = string_index_node.findall('{%s}r' % SHEET_MAIN_NS)
    if rich_nodes:
        reconstructed_text = []
        for rich_node in rich_nodes:
            partial_text = get_text(rich_node)
            reconstructed_text.append(partial_text)
        return unicode(''.join(reconstructed_text))
    else:
        return get_text(string_index_node)


def get_text(rich_node):
    """Read rich text, discarding formatting if not disallowed"""
    text_node = rich_node.find('{%s}t' % SHEET_MAIN_NS)
    partial_text = text_node.text or unicode('')

    if text_node.get('{%s}space' % XML_NS) != 'preserve':
        partial_text = partial_text.strip()
    return unicode(partial_text)
