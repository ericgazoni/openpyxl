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

import math

from openpyxl.styles import NumberFormat, is_date_format


def less_than_one(value):
    """Recalculate the maximum for a series if it is less than one
    by scaling by powers of 10 until is greater than 1
    """
    value = abs(value)
    if value < 1:
        exp = int(math.log10(value))
        return 10**((abs(exp)) + 1)


class Axis(object):

    POSITION_BOTTOM = 'b'
    POSITION_LEFT = 'l'
    ORIENTATION_MIN_MAX = "minMax"

    position = None
    tick_label_position = None
    crosses = None
    auto = None
    label_align = None
    label_offset = None
    cross_between = None
    orientation = ORIENTATION_MIN_MAX
    number_format = NumberFormat()
    delete_axis = False

    def __init__(self, auto_axis=True):
        self.auto_axis = auto_axis
        self.min = 0
        self.max = 0
        self.unit = None
        self.title = ''

    def _max_min(self):
        """
        Calculate minimum and maximum for the axis adding some padding.
        There are always a maximum of ten units for the length of the axis.
        """
        value = length = self._max - self._min

        sign = value/value
        zoom = less_than_one(value) or 1
        value = value * zoom
        ab = abs(value)
        value = math.ceil(ab * 1.1) * sign

        # calculate tick
        l = math.log10(abs(value))
        exp = int(l)
        mant = l - exp
        unit = math.ceil(math.ceil(10**mant) * 10**(exp-1))
        # recalculate max
        value = math.ceil(value / unit) * unit
        unit = unit / zoom

        if value / unit > 9:
            # no more that 10 ticks
            unit *= 2
        self.unit = unit
        scale = value / length
        mini = math.floor(self._min * scale) / zoom
        maxi = math.ceil(self._max * scale) / zoom
        return mini, maxi

    @property
    def min(self):
        if self.auto_axis:
            return self._max_min()[0]
        return self._min

    @min.setter
    def min(self, value):
        self._min = value

    @property
    def max(self):
        if self.auto_axis:
            return self._max_min()[1]
        return self._max

    @max.setter
    def max(self, value):
        self._max = value

    @property
    def unit(self):
        if self.auto_axis:
            self._max_min()
        return self._unit

    @unit.setter
    def unit(self, value):
        self._unit = value


class CategoryAxis(Axis):

    id = 60871424
    cross = 60873344
    position = Axis.POSITION_BOTTOM
    tick_label_position = 'nextTo'
    crosses = "autoZero"
    auto = True
    label_align = 'ctr'
    label_offset = 100
    cross_between = "midCat"
    type = "catAx"


class ValueAxis(Axis):

    id = 60873344
    cross = 60871424
    position = Axis.POSITION_LEFT
    major_gridlines = None
    tick_label_position = 'nextTo'
    crosses = 'autoZero'
    auto = False
    cross_between = 'between'
    type= "valAx"
