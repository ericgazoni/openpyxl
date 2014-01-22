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


class RowDimension(object):
    """Information about the display properties of a row."""
    __slots__ = ('parent',
                 'row_index',
                 'height',
                 'visible',
                 'outline_level',
                 'collapsed',)

    def __init__(self,
                 worksheet,
                 index=0,
                 height=-1,
                 visible=True,
                 outline_level=0,
                 collapsed=False):
        self.parent = worksheet
        self.row_index = index
        self.height = height
        self.visible = visible
        self.outline_level = outline_level
        self.collapsed = collapsed

    @property
    def style(self):
        """Returns the :class:`openpyxl.styles.Style` object for this row"""
        return self.parent.get_style(self.row_index)


class ColumnDimension(object):
    """Information about the display properties of a column."""
    __slots__ = ('parent',
                 'column_index',
                 'width',
                 'auto_size',
                 'visible',
                 'outline_level',
                 'collapsed',)

    def __init__(self,
                 worksheet,
                 index='A',
                 width=-1,
                 auto_size=False,
                 visible=True,
                 outline_level=0,
                 collapsed=False):
        self.parent = worksheet
        self.column_index = index
        self.width = float(width)
        self.auto_size = auto_size
        self.visible = visible
        self.outline_level = int(outline_level)
        self.collapsed = collapsed

    @property
    def style(self):
        """Returns the :class:`openpyxl.styles.Style` object for this column"""
        return self.parent.get_style(self.column_index)


