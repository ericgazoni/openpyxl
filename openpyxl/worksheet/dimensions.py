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

from openpyxl.cell import get_column_letter


class Dimension(object):
    """Information about the display properties of a row or column."""
    __slots__ = ('index',
                 'visible',
                 'outline_level',
                 'collapsed',)

    def __init__(self,
                 index,
                 visible,
                 outline_level,
                 collapsed):
        self.index = index
        self.visible = visible
        self.outline_level = outline_level
        self.collapsed = collapsed


class RowDimension(Dimension):
    """Information about the display properties of a row."""

    __slots__ = Dimension.__slots__ + ('height',)

    def __init__(self,
                 index=0,
                 height=-1,
                 visible=True,
                 outline_level=0,
                 collapsed=False):
        super(RowDimension, self).__init__(index, visible, outline_level,
                                           collapsed)
        self.height = float(height)


class ColumnDimension(Dimension):
    """Information about the display properties of a column."""

    __slots__ = Dimension.__slots__ + ('width', 'auto_size')

    def __init__(self,
                 index='A',
                 width=-1,
                 auto_size=False,
                 visible=True,
                 outline_level=0,
                 collapsed=False):
        super(ColumnDimension, self).__init__(index, visible, outline_level,
                                              collapsed)
        self.width = float(width)
        self.auto_size = auto_size

    #@property
    #def col_label(self):
        #return get_column_letter(self.index)
