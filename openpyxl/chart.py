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

from openpyxl.style import NumberFormat
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

    def __init__(self):

        self.orientation = self.ORIENTATION_MIN_MAX
        self.number_format = NumberFormat()
        for attr in ('position', 'tick_label_position', 'crosses',
            'auto', 'label_align', 'label_offset', 'cross_between'):
            setattr(self, attr, None)
        self.min = 0
        self.max = 0
        self.unit = None
        self.title = ''

    def set_values(self, mini, maxi):
        self.min = mini
        self.max = maxi

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
        mini, maxi = self._max_min()
        return mini

    @min.setter
    def min(self, value):
        self._min = value

    @property
    def max(self):
        mini, maxi = self._max_min()
        return maxi

    @max.setter
    def max(self, value):
        self._max = value

    @property
    def unit(self):
        self._max_min
        return self._unit

    @unit.setter
    def unit(self, value):
        self._unit = value

    @classmethod
    def default_category(cls):
        """ default values for category axes """

        ax = Axis()
        ax.id = 60871424
        ax.cross = 60873344
        ax.position = Axis.POSITION_BOTTOM
        ax.tick_label_position = 'nextTo'
        ax.crosses = "autoZero"
        ax.auto = True
        ax.label_align = 'ctr'
        ax.label_offset = 100
        return ax

    @classmethod
    def default_value(cls):
        """ default values for value axes """

        ax = Axis()
        ax.id = 60873344
        ax.cross = 60871424
        ax.position = Axis.POSITION_LEFT
        ax.major_gridlines = None
        ax.tick_label_position = 'nextTo'
        ax.crosses = 'autoZero'
        ax.auto = False
        ax.cross_between = 'between'
        return ax


class Reference(object):
    """ a simple wrapper around a serie of reference data """


    def __init__(self, sheet, pos1, pos2=None, data_type='n'):

        self.sheet = sheet
        self.pos1 = pos1
        self.pos2 = pos2
        self.data_type = data_type

    def get_type(self):
        """Legacy method"""
        return self.data_type

    @property
    def data_type(self):
        return self._data_type

    @data_type.setter
    def data_type(self, value):
        if value not in ['n', 's']:
            raise ValueError("References must be either numeric or strings")
        self._data_type = value

    @property
    def values(self):
        """ read data in sheet - to be used at writing time """
        if hasattr(self, "_values"):
            return self._values
        if self.pos2 is None:
            cell = self.sheet.cell(row=self.pos1[0], column=self.pos1[1])
            self.data_type = cell.data_type
            self._values = [cell.excel_value]
        else:
            self._values = []

            for row in range(int(self.pos1[0]), int(self.pos2[0] + 1)):
                for col in range(int(self.pos1[1]), int(self.pos2[1] + 1)):
                    cell = self.sheet.cell(row=row, column=col)
                    self._values.append(cell.excel_value)

            if self.data_type is None:
                self.data_type = 'n'

        return self._values

    def _get_ref(self):
        """ legace method """
        return str(self)

    def __str__(self):
        """ format excel reference notation """

        if self.pos2:
            return "'%s'!$%s$%s:$%s$%s" % (self.sheet.title,
                get_column_letter(self.pos1[1] + 1), self.pos1[0] + 1,
                get_column_letter(self.pos2[1] + 1), self.pos2[0] + 1)
        else:
            return "'%s'!$%s$%s" % (self.sheet.title,
                get_column_letter(self.pos1[1] + 1), self.pos1[0] + 1)

    def _get_cache(self):
        """ legacy method """
        return self.values


class Serie(object):
    """ a serie of data and possibly associated labels """

    MARKER_NONE = 'none'

    def __init__(self, values, labels=None, legend=None, color=None,
                 xvalues=None, data_type='n'):

        self.marker = Serie.MARKER_NONE
        self.values = values
        self.xvalues = xvalues
        self.labels = labels
        self.legend = legend
        self.error_bar = None
        self.data_type = data_type

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

    def mymax(self, values):
        return max([x for x in values])

    def max(self):
        """
        Return the maximum value for numeric series.
        NB None has a value of u'' which is ignored
        """
        if self.data_type == 'n':
            cleaned = [v for v in self.values if v]
            if cleaned:
                return max(cleaned)

    def min(self):
        """
        Return the minimum value for numeric series
        NB None has a value of u'' which is ignored
        """
        if self.data_type == 'n':
            cleaned = [v for v in self.values if v]
            if cleaned:
                return min(cleaned)

    def get_min_max(self):

        if self.error_bar:
            err_cache = self.error_bar.values
            vals = [v + err_cache[i] \
                for i, v in enumerate(self.values)]
        else:
            vals = self.values
        return min(vals), self.mymax(vals)

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
        self.reference = reference


class Chart(object):
    """ raw chart class """

    GROUPING_CLUSTERED = 'clustered'
    GROUPING_STANDARD = 'standard'

    BAR_CHART = 1
    LINE_CHART = 2
    SCATTER_CHART = 3

    def mymax(self, values):
        return max([x for x in values if x is not None])

    def mymin(self, values):
        return min([x for x in values if x is not None])

    def __init__(self, _type, grouping):

        self._series = []

        # public api
        self.type = _type
        self.grouping = grouping
        self.x_axis = Axis.default_category()
        self.y_axis = Axis.default_value()
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

    def get_x_units(self):
        """ calculate one unit for x axis in EMU """

        return self.mymax([len(s.values) for s in self._series])

    def get_y_units(self):
        """ calculate one unit for y axis in EMU """

        dh = pixels_to_EMU(self.drawing.height)
        return (dh * self.height) / self.y_axis.max

    def get_y_chars(self):
        """ estimate nb of chars for y axis """

        _max = max([self.mymax(s.values) for s in self._series])
        return len(str(int(_max)))

    def compute_axes(self):
        """Calculate maximum value and units for axes"""
        mini, maxi = self._get_extremes()
        self.y_axis.set_values(mini, maxi)

        if not None in [s.xvalues for s in self._series]:
            mini, maxi = self._get_extremes('xvalues')
            self.x_axis.set_values(mini, maxi)

    def _get_extremes(self, attr='values'):
        """Calculate the maximum and minimum values of all series for an axis
        'values' for columns
        'xvalues for rows
        """
        # calculate the maximum for all series
        series_max = [0]
        series_min = [0]
        for s in self._series:
            series = getattr(s, attr)
            if series is not None:
                maxi = self.mymax(series)
                series_max.append(maxi)
                mini = self.mymin(series)
                series_min.append(mini)
        return min(series_min), max(series_max)

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

    @property
    def y_labels(self):
        """Labels for the y-axis"""
        return []

    @property
    def x_labels(self):
        """Labels for the x-axis"""
        return []


class BarChart(Chart):
    def __init__(self):
        super(BarChart, self).__init__(Chart.BAR_CHART, Chart.GROUPING_CLUSTERED)


class LineChart(Chart):
    def __init__(self):
        super(LineChart, self).__init__(Chart.LINE_CHART, Chart.GROUPING_STANDARD)


class ScatterChart(Chart):
    def __init__(self):
        super(ScatterChart, self).__init__(Chart.SCATTER_CHART, Chart.GROUPING_STANDARD)
