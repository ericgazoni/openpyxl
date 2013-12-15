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

import math
from numbers import Number

from openpyxl.style import NumberFormat, is_date_format, is_builtin
from openpyxl.drawing import Drawing, Shape
from openpyxl.shared.units import pixels_to_EMU, short_color
from openpyxl.cell import get_column_letter


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


class Reference(object):
    """ a simple wrapper around a serie of reference data """

    _data_type = None

    def __init__(self, sheet, pos1, pos2=None, data_type=None, number_format=None):

        self.sheet = sheet
        self.pos1 = pos1
        self.pos2 = pos2
        if data_type is not None:
            self.data_type = data_type
        self.number_format = number_format

    @property
    def data_type(self):
        return self._data_type

    @data_type.setter
    def data_type(self, value):
        if value not in ['n', 's']:
            raise ValueError("References must be either numeric or strings")
        self._data_type = value

    @property
    def number_format(self):
        return self._number_format

    @number_format.setter
    def number_format(self, value):
        if value is not None:
            if not is_builtin(value):
                raise ValueError("Invalid number format")
        self._number_format = value

    @property
    def values(self):
        """ read data in sheet - to be used at writing time """
        if hasattr(self, "_values"):
            return self._values
        if self.pos2 is None:
            cell = self.sheet.cell(row=self.pos1[0], column=self.pos1[1])
            self.data_type = cell.data_type
            self._values = [cell.internal_value]
        else:
            self._values = []

            for row in range(int(self.pos1[0]), int(self.pos2[0] + 1)):
                for col in range(int(self.pos1[1]), int(self.pos2[1] + 1)):
                    cell = self.sheet.cell(row=row, column=col)
                    self._values.append(cell.internal_value)
                    if cell.internal_value == '':
                        continue
                    if self.data_type is None and cell.data_type:
                        self.data_type = cell.data_type
        return self._values

    def __str__(self):
        """ format excel reference notation """

        if self.pos2:
            return "'%s'!$%s$%s:$%s$%s" % (self.sheet.title,
                get_column_letter(self.pos1[1] + 1), self.pos1[0] + 1,
                get_column_letter(self.pos2[1] + 1), self.pos2[0] + 1)
        else:
            return "'%s'!$%s$%s" % (self.sheet.title,
                get_column_letter(self.pos1[1] + 1), self.pos1[0] + 1)


class Serie(object):
    """ a serie of data and possibly associated labels """

    MARKER_NONE = 'none'

    def __init__(self, values, labels=None, legend=None, color=None,
                 xvalues=None):

        self.marker = Serie.MARKER_NONE
        self.values = values
        self.xvalues = xvalues
        self.labels = labels
        self.legend = legend
        if legend is not None:
            self.legend.data_type = 's'
        self.error_bar = None

    @property
    def color(self):
        return getattr(self, "_color", None)

    @color.setter
    def color(self, color):
        if color is None:
            raise ValueError("Colors must be strings of the format XXXXX")
        self._color = short_color(color)

    @property
    def values(self):
        """Return values from underlying reference"""
        return self._values

    @values.setter
    def values(self, reference):
        """Assign values from reference to serie"""
        if reference is not None:
            if not isinstance(reference, Reference):
                raise TypeError("Series values must be a Reference instance")
            self._values = reference.values
        else:
            self._values = None
        self.reference = reference

    @property
    def xvalues(self):
        """Return xvalues"""
        return self._xvalues

    @xvalues.setter
    def xvalues(self, reference):
        if reference is not None:
            if not isinstance(reference, Reference):
                raise TypeError("Series xvalues must be a Reference instance")
            self._xvalues = reference.values
        else:
            self._xvalues = None
        self.xreference = reference

    @property
    def labels(self):
        """Return values from reference set as label"""
        return self._labels

    @labels.setter
    def labels(self, reference):
        if reference is not None:
            if not isinstance(reference, Reference):
                raise TypeError("Series labels must be a Reference instance")
            reference.values
            self._labels = reference
        else:
            self._labels = None

    def max(self, attr='values'):
        """
        Return the maximum value for numeric series.
        NB None has a value of u'' which is ignored
        """
        values = getattr(self, attr)
        if self.error_bar:
            values = self._error_bar_values
        cleaned = [v for v in values if isinstance(v, Number)]
        if cleaned:
            return max(cleaned)

    def min(self, attr='values'):
        """
        Return the minimum value for numeric series
        NB None has a value of u'' which is ignored
        """
        values = getattr(self, attr)
        if self.error_bar:
            values = self._error_bar_values
        cleaned = [v for v in values if isinstance(v, Number)]
        if cleaned:
            return min(cleaned)

    @property
    def _error_bar_values(self):
        """Documentation required here"""
        err_cache = self.error_bar.values
        vals = [v + err_cache[i] \
            for i, v in enumerate(self.values)]
        return vals

    def get_min_max(self):
        """Legacy method. Replaced by properties"""
        return self.min(), self.max()

    def __len__(self):

        return len(self.values)


class Legend(object):

    def __init__(self):

        self.position = 'r'
        self.layout = None


class ErrorBar(object):

    PLUS = 1
    MINUS = 2
    PLUS_MINUS = 3

    def __init__(self, _type, values):

        self.type = _type
        self.values = values

    @property
    def values(self):
        """Return values from underlying reference"""
        return self._values

    @values.setter
    def values(self, reference):
        """Assign values from reference to serie"""
        if reference is not None:
            if not isinstance(reference, Reference):
                raise TypeError("Errorbar values must be a Reference instance")
            self._values = reference.values
        else:
            self._values = None


class Chart(object):
    """ raw chart class """

    GROUPING = 'standard'
    TYPE = None

    def mymax(self, values):
        return max([x for x in values if x is not None])

    def mymin(self, values):
        return min([x for x in values if x is not None])

    def __init__(self):

        self._series = []

        # public api
        self.legend = Legend()
        self.show_legend = True
        self.lang = 'en-GB'
        self.title = ''
        self.print_margins = dict(b=.75, l=.7, r=.7, t=.75, header=0.3, footer=.3)

        # the containing drawing
        self.drawing = Drawing()
        self.drawing.left = 10
        self.drawing.top = 400
        self.drawing.height = 400
        self.drawing.width = 800

        # the offset for the plot part in percentage of the drawing size
        self.width = .6
        self.height = .6
        self._margin_top = 1
        self._margin_top = self.margin_top
        self._margin_left = 0

        # the user defined shapes
        self._shapes = []

    def add_serie(self, serie):

        serie.id = len(self._series)
        self._series.append(serie)

    def add_shape(self, shape):

        shape._chart = self
        self._shapes.append(shape)

    def get_y_chars(self):
        """ estimate nb of chars for y axis """
        _max = max([s.max() for s in self._series])
        return len(str(int(_max)))

    @property
    def margin_top(self):
        """ get margin in percent """

        return min(self._margin_top, self._get_max_margin_top())

    @margin_top.setter
    def margin_top(self, value):
        """ set base top margin"""
        self._margin_top = value

    def _get_max_margin_top(self):

        mb = Shape.FONT_HEIGHT + Shape.MARGIN_BOTTOM
        plot_height = self.drawing.height * self.height
        return float(self.drawing.height - plot_height - mb) / self.drawing.height

    @property
    def margin_left(self):

        return max(self._get_min_margin_left(), self._margin_left)

    @margin_left.setter
    def margin_left(self, value):
        self._margin_left = value

    def _get_min_margin_left(self):

        ml = (self.get_y_chars() * Shape.FONT_WIDTH) + Shape.MARGIN_LEFT
        return float(ml) / self.drawing.width


class PieChart(Chart):

    TYPE = "pieChart"


class GraphChart(Chart):
    """Chart with axes"""

    x_axis = CategoryAxis
    y_axis = ValueAxis

    def __init__(self, auto_axis=True):
        super(GraphChart, self).__init__()
        self.auto_axis = auto_axis
        self.x_axis = getattr(self, "x_axis")(auto_axis)
        self.y_axis = getattr(self, "y_axis")(auto_axis)

    def compute_axes(self):
        """Calculate maximum value and units for axes"""
        mini, maxi = self._get_extremes()
        self.y_axis.min = mini
        self.y_axis.max = maxi
        self.y_axis._max_min()

        if not None in [s.xvalues for s in self._series]:
            mini, maxi = self._get_extremes('xvalues')
            self.x_axis.min = mini
            self.x_axis.max = maxi
            self.x_axis._max_min()

    def get_x_units(self):
        """ calculate one unit for x axis in EMU """
        return max([len(s.values) for s in self._series])

    def get_y_units(self):
        """ calculate one unit for y axis in EMU """

        dh = pixels_to_EMU(self.drawing.height)
        return (dh * self.height) / self.y_axis.max

    def _get_extremes(self, attr='values'):
        """Calculate the maximum and minimum values of all series for an axis
        'values' for columns
        'xvalues for rows
        """
        # calculate the maximum and minimum for all series
        series_max = [0]
        series_min = [0]
        for s in self._series:
            if s is not None:
                series_max.append(s.max(attr))
                series_min.append(s.min(attr))
        return min(series_min), max(series_max)


class BarChart(GraphChart):

    TYPE = "barChart"
    GROUPING = "clustered"


class LineChart(GraphChart):

    TYPE = "lineChart"


class ScatterChart(GraphChart):

    TYPE = "scatterChart"

    def __init__(self):
        super(ScatterChart, self).__init__()
        self.x_axis.type = "valAx"
        self.x_axis.cross_between = "midCat"
        self.y_axis.cross_between = "midCat"
