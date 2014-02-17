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

from .hashable import HashableObject

# Default Color Index as per http://dmcritchie.mvps.org/excel/colors.htm
COLOR_INDEX = ('FF000000', 'FFFFFFFF', 'FFFF0000', 'FF00FF00', 'FF0000FF',
               'FFFFFF00', 'FFFF00FF', 'FF00FFFF', 'FF800000', 'FF008000', 'FF000080',
               'FF808000', 'FF800080', 'FF008080', 'FFC0C0C0', 'FF808080', 'FF9999FF',
               'FF993366', 'FFFFFFCC', 'FFCCFFFF', 'FF660066', 'FFFF8080', 'FF0066CC',
               'FFCCCCFF', 'FF000080', 'FFFF00FF', 'FFFFFF00', 'FF00FFFF', 'FF800080',
               'FF800000', 'FF008080', 'FF0000FF', 'FF00CCFF', 'FFCCFFFF', 'FFCCFFCC',
               'FFFFFF99', 'FF99CCFF', 'FFFF99CC', 'FFCC99FF', 'FFFFCC99', 'FF3366FF',
               'FF33CCCC', 'FF99CC00', 'FFFFCC00', 'FFFF9900', 'FFFF6600', 'FF666699',
               'FF969696', 'FF003366', 'FF339966', 'FF003300', 'FF333300', 'FF993300',
               'FF993366', 'FF333399', 'FF333333')

class Color(HashableObject):
    """Named colors for use in styles."""
    BLACK = 'FF000000'
    WHITE = 'FFFFFFFF'
    RED = 'FFFF0000'
    DARKRED = 'FF800000'
    BLUE = 'FF0000FF'
    DARKBLUE = 'FF000080'
    GREEN = 'FF00FF00'
    DARKGREEN = 'FF008000'
    YELLOW = 'FFFFFF00'
    DARKYELLOW = 'FF808000'

    __fields__ = ('index',)
    __slots__ = __fields__
    __leaf__ = True

    def __init__(self, index):
        self.index = index
