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

from .colors import Color
from .hashable import HashableObject


class Border(HashableObject):
    """Border options for use in styles."""
    BORDER_NONE = 'none'
    BORDER_DASHDOT = 'dashDot'
    BORDER_DASHDOTDOT = 'dashDotDot'
    BORDER_DASHED = 'dashed'
    BORDER_DOTTED = 'dotted'
    BORDER_DOUBLE = 'double'
    BORDER_HAIR = 'hair'
    BORDER_MEDIUM = 'medium'
    BORDER_MEDIUMDASHDOT = 'mediumDashDot'
    BORDER_MEDIUMDASHDOTDOT = 'mediumDashDotDot'
    BORDER_MEDIUMDASHED = 'mediumDashed'
    BORDER_SLANTDASHDOT = 'slantDashDot'
    BORDER_THICK = 'thick'
    BORDER_THIN = 'thin'

    __fields__ = ('border_style',
                  'color')
    __slots__ = __fields__

    def __init__(self):
        self.border_style = self.BORDER_NONE
        self.color = Color(Color.BLACK)


class Borders(HashableObject):
    """Border positioning for use in styles."""
    DIAGONAL_NONE = 0
    DIAGONAL_UP = 1
    DIAGONAL_DOWN = 2
    DIAGONAL_BOTH = 3

    __fields__ = ('left',
                  'right',
                  'top',
                  'bottom',
                  'diagonal',
                  'diagonal_direction',
                  'all_borders',
                  'outline',
                  'inside',
                  'vertical',
                  'horizontal')
    __slots__ = __fields__

    def __init__(self):
        self.left = Border()
        self.right = Border()
        self.top = Border()
        self.bottom = Border()
        self.diagonal = Border()
        self.diagonal_direction = self.DIAGONAL_NONE

        self.all_borders = Border()
        self.outline = Border()
        self.inside = Border()
        self.vertical = Border()
        self.horizontal = Border()
