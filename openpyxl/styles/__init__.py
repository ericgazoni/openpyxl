# Copyright (c) 2010-2013 openpyxl
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

from copy import deepcopy

from .alignment import Alignment
from .borders import Borders, Border
from .colors import Color
from .fills import Fill
from .fonts import Font
from .hashable import HashableObject
from .numbers import NumberFormat, is_date_format, is_builtin
from .protection import Protection


class Style(HashableObject):
    """Style object containing all formatting details."""
    __fields__ = ('font',
                  'fill',
                  'borders',
                  'alignment',
                  'number_format',
                  'protection')
    __slots__ = __fields__

    def __init__(self, static=False):
        self.static = static
        self.font = Font()
        self.fill = Fill()
        self.borders = Borders()
        self.alignment = Alignment()
        self.number_format = NumberFormat()
        self.protection = Protection()

    def copy(self):
        new_style = Style()
        new_style.font = deepcopy(self.font)
        new_style.fill = deepcopy(self.fill)
        new_style.borders = deepcopy(self.borders)
        new_style.alignment = deepcopy(self.alignment)
        new_style.number_format = deepcopy(self.number_format)
        new_style.protection = deepcopy(self.protection)
        return new_style

DEFAULTS = Style()

