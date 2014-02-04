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

from openpyxl.xml.constants import REL_NS, PKG_REL_NS
from openpyxl.xml.functions import Element, SubElement, get_document_content

class Relationship(object):
    """Represents many kinds of relationships."""
    # TODO: Use this object for workbook relationships as well as
    # worksheet relationships

    TYPES = ("hyperlink", "drawing", "image")

    def __init__(self, rel_type, target=None, target_mode=None, id=None):
        if rel_type not in self.TYPES:
            raise ValueError("Invalid relationship type %s" % rel_type)
        self.type = "%s/%s" % (REL_NS, rel_type)
        self.target = target
        self.target_mode = target_mode
        self.id = id

    def __repr__(self):
        root = Element("{%s}Relationships" % PKG_REL_NS)
        body = SubElement(root, "{%s}Relationship" % PKG_REL_NS, self.__dict__)
        return get_document_content(root)

