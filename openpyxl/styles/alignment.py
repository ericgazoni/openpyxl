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

from .hashable import HashableObject


class Alignment(HashableObject):
    """Alignment options for use in styles."""
    HORIZONTAL_GENERAL = 'general'
    HORIZONTAL_LEFT = 'left'
    HORIZONTAL_RIGHT = 'right'
    HORIZONTAL_CENTER = 'center'
    HORIZONTAL_CENTER_CONTINUOUS = 'centerContinuous'
    HORIZONTAL_JUSTIFY = 'justify'
    VERTICAL_BOTTOM = 'bottom'
    VERTICAL_TOP = 'top'
    VERTICAL_CENTER = 'center'
    VERTICAL_JUSTIFY = 'justify'

    __fields__ = ('horizontal',
                  'vertical',
                  'text_rotation',
                  'wrap_text',
                  'shrink_to_fit',
                  'indent')
    __slots__ = __fields__
    __leaf__ = True

    def __init__(self):
        self.horizontal = self.HORIZONTAL_GENERAL
        self.vertical = self.VERTICAL_BOTTOM
        self.text_rotation = 0
        self.wrap_text = False
        self.shrink_to_fit = False
        self.indent = 0
