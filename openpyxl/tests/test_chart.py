# file openpyxl/tests/test_chart.py

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

from datetime import date
import pytest


@pytest.mark.parametrize("value, result",
                         [
                          (1, None),
                          (0.9, 10),
                          (0.09, 100),
                          (-0.09, 100)
                         ]
                         )
def test_less_than_one(value, result):
    from openpyxl.chart import less_than_one
    assert less_than_one(value) == result

def test_axis_ctor(Axis):
    axis = Axis()
    assert axis.title == ""
    assert axis.auto_axis is True
    with pytest.raises(ZeroDivisionError):
        axis.max == 0
    with pytest.raises(ZeroDivisionError):
        axis.min == 0
    with pytest.raises(ZeroDivisionError):
        axis.unit == 0


@pytest.mark.parametrize("set_max, set_min, min, max, unit",
                         [
                         (10, 0, 0, 12, 2),
                         (5, 0, 0, 6, 1),
                         (50000, 0, 0, 60000, 12000),
                         (1, 0, 0, 2, 1),
                         (0.9, 0, 0, 1, 0.2),
                         (0.09, 0, 0, 0.1, 0.02),
                         (0, -0.09, -0.1, 0, 0.02),
                         (8, -2, -3, 10, 2)
                         ]
                         )
def test_scaling(Axis, set_max, set_min, min, max, unit):
    axis = Axis()
    axis.max = set_max
    axis.min = set_min
    assert axis.min == min
    assert axis.max == max
    assert axis.unit == unit


@pytest.fixture
def sheet(ten_row_sheet):
    ten_row_sheet.title = "reference"
    return ten_row_sheet


@pytest.fixture
def cell(sheet, Reference):
    return Reference(sheet, (0, 0))


@pytest.fixture
def cell_range(sheet, Reference):
    return Reference(sheet, (0, 0), (9, 0))


@pytest.fixture()
def empty_range(sheet, Reference):
    for i in range(10):
        sheet.cell(row=i, column=1).value = None
    return Reference(sheet, (0, 1), (9, 1))


@pytest.fixture()
def missing_values(sheet, Reference):
    vals = [None, None, 1, 2, 3, 4, 5, 6, 7, 8]
    for idx, val in enumerate(vals):
        sheet.cell(row=idx, column=2).value = val
    return Reference(sheet, (0, 2), (9, 2))


@pytest.fixture
def column_of_letters(sheet, Reference):
    for idx, l in enumerate("ABCDEFGHIJ"):
        sheet.cell(row=idx, column=1).value = l
    return Reference(sheet, (0, 1), (9, 1))

class TestReference(object):

    def test_single_cell_ctor(self, cell):
        assert cell.pos1 == (0, 0)
        assert cell.pos2 == None

    def test_range_ctor(self, cell_range):
        assert cell_range.pos1 == (0, 0)
        assert cell_range.pos2 == (9, 0)

    def test_single_cell_ref(self, cell):
        assert cell.values == [0]
        assert str(cell) == "'reference'!$A$1"

    def test_cell_range_ref(self, cell_range):
        assert cell_range.values == [0, 1, 2, 3, 4, 5, 6, 7, 8 , 9]
        assert str(cell_range) == "'reference'!$A$1:$A$10"

    def test_data_type(self, cell):
        with pytest.raises(ValueError):
            cell.data_type = 'f'
            cell.data_type = None

    def test_type_inference(self, cell, cell_range, column_of_letters,
                            missing_values):
        assert cell.values == [0]
        assert cell.data_type == 'n'

        assert cell_range.values == [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
        assert cell_range.data_type == 'n'

        assert column_of_letters.values == list("ABCDEFGHIJ")
        assert column_of_letters.data_type == "s"

        assert missing_values.values == ['', '', 1, 2, 3, 4, 5, 6, 7, 8]
        missing_values.values
        assert missing_values.data_type == 'n'

    def test_number_format(self, cell):
        with pytest.raises(ValueError):
            cell.number_format = 'YYYY'
        cell.number_format = 'd-mmm'
        assert cell.number_format == 'd-mmm'


class TestErrorBar(object):

    def test_ctor(self, ErrorBar):
        with pytest.raises(TypeError):
            ErrorBar(None, range(10))


class TestSerie(object):

    def test_ctor(self, Serie, cell):
        series = Serie(cell)
        assert series.values == [0]
        assert series.color == None
        assert series.error_bar == None
        assert series.xvalues == None
        assert series.labels == None
        assert series.title == None

    def test_invalid_values(self, Serie, cell):
        series = Serie(cell)
        with pytest.raises(TypeError):
            series.values = 0

    def test_invalid_xvalues(self, Serie, cell):
        series = Serie(cell)
        with pytest.raises(TypeError):
            series.xvalues = 0

    def test_color(self, Serie, cell):
        series = Serie(cell)
        assert series.color == None
        series.color = "blue"
        assert series.color, "blue"
        with pytest.raises(ValueError):
            series.color = None

    def test_min(self, Serie, cell, cell_range, empty_range):
        series = Serie(cell)
        assert series.min() == 0
        series = Serie(cell_range)
        assert series.min() == 0
        series = Serie(empty_range)
        assert series.min() == None

    def test_max(self, Serie, cell, cell_range, empty_range):
        series = Serie(cell)
        assert series.max() == 0
        series = Serie(cell_range)
        assert series.max() == 9
        series = Serie(empty_range)
        assert series.max() == None

    def test_min_max(self, Serie, cell, cell_range, empty_range):
        series = Serie(cell)
        assert series.get_min_max() == (0, 0)
        series = Serie(cell_range)
        assert series.get_min_max() == (0, 9)
        series = Serie(empty_range)
        assert series.get_min_max() == (None, None)

    def test_len(self, Serie, cell):
        series = Serie(cell)
        assert len(series) == 1

    def test_error_bar(self, Serie, ErrorBar, cell):
        series = Serie(cell)
        series.error_bar = ErrorBar(None, cell)
        assert series.get_min_max() == (0, 0)


@pytest.fixture()
def series(cell_range, Serie):
    return Serie(values=cell_range)


class TestChart(object):

    def test_ctor(self, Chart):
        from openpyxl.chart import Legend
        from openpyxl.drawing import Drawing
        c = Chart()
        assert c.TYPE == None
        assert c.GROUPING == "standard"
        assert isinstance(c.legend, Legend)
        assert c.show_legend
        assert c.lang == 'en-GB'
        assert c.title == ''
        assert c.print_margins == {'b':0.75, 'l':0.7, 'r':0.7, 't':0.75,
                                   'header':0.3, 'footer':0.3}
        assert isinstance(c.drawing, Drawing)
        assert c.width == 0.6
        assert c.height == 0.6
        assert c.margin_top == 0.31
        assert c._shapes == []
        with pytest.raises(ValueError):
            assert c.margin_left == 0

    def test_mymax(self, Chart):
        c = Chart()
        assert c.mymax(range(10)) == 9
        from string import ascii_letters as letters
        assert c.mymax(list(letters)) == "z"
        assert c.mymax(range(-10, 1)) == 0
        assert c.mymax([""]*10) == ""

    def test_mymin(self, Chart):
        c = Chart()
        assert c.mymin(range(10)) == 0
        from string import ascii_letters as letters
        assert c.mymin(list(letters)) == "A"
        assert c.mymin(range(-10, 1)) == -10
        assert c.mymin([""]*10) == ""

    def test_margin_top(self, Chart):
        c = Chart()
        assert c.margin_top == 0.31

    def test_margin_left(self, series, Chart):
        c = Chart()
        c._series.append(series)
        assert c.margin_left == 0.03375

    def test_set_margin_top(self, Chart):
        c = Chart()
        c.margin_top = 1
        assert c.margin_top == 0.31

    def test_set_margin_left(self, series, Chart):
        c = Chart()
        c._series.append(series)
        c.margin_left = 0
        assert c.margin_left  == 0.03375


class TestGraphChart(object):

    def test_ctor(self, GraphChart, Axis):
        c = GraphChart()
        assert isinstance(c.x_axis, Axis)
        assert isinstance(c.y_axis, Axis)

    def test_get_x_unit(self, GraphChart, series):
        c = GraphChart()
        c._series.append(series)
        assert c.get_x_units() == 10

    def test_get_y_unit(self, GraphChart, series):
        c = GraphChart()
        c._series.append(series)
        c.y_axis.max = 10
        assert c.get_y_units() == 190500

    def test_get_y_char(self, GraphChart, series):
        c = GraphChart()
        c._series.append(series)
        assert c.get_y_chars() == 1

    def test_compute_series_extremes(self, GraphChart, series):
        c = GraphChart()
        c._series.append(series)
        mini, maxi = c._get_extremes()
        assert mini == 0
        assert maxi == 9

    def test_compute_series_max_dates(self, ws, Reference, Serie, GraphChart):
        for i in range(1, 10):
            ws.append([date(2013, i, 1)])
        c = GraphChart()
        ref = Reference(ws, (0, 0), (9, 0))
        series = Serie(ref)
        c._series.append(series)
        mini, maxi = c._get_extremes()
        assert mini == 0
        assert maxi == 41518.0

    def test_override_axis(self, GraphChart, series):
        c = GraphChart()
        c.add_serie(series)
        c.compute_axes()
        assert c.y_axis.min == 0
        assert c.y_axis.max == 10
        c.y_axis.min = -1
        c.y_axis.max = 5
        assert c.y_axis.min == -2
        assert c.y_axis.max == 6


class TestLineChart(object):

    def test_ctor(self, LineChart):
        c = LineChart()
        assert c.TYPE == "lineChart"
        assert c.x_axis.type == "catAx"
        assert c.y_axis.type == "valAx"


class TestPieChart(object):

    def test_ctor(self, PieChart):
        c = PieChart()
        assert c.TYPE, "pieChart"


class TestBarChart(object):

    def test_ctor(self, BarChart):
        c = BarChart()
        assert c.TYPE == "barChart"
        assert c.x_axis.type == "catAx"
        assert c.y_axis.type == "valAx"


class TestScatterChart(object):

    def test_ctor(self, ScatterChart):
        c = ScatterChart()
        assert c.TYPE == "scatterChart"
        assert c.x_axis.type == "valAx"
        assert c.y_axis.type == "valAx"
