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
import os

import pytest

from openpyxl.tests.helper import (get_xml,
                                   DATADIR,
                                   compare_xml,
                                   safe_iterator,
                                   )

from openpyxl.shared.xmltools import Element, get_document_content
from openpyxl.shared.ooxml import CHART_NS
from openpyxl.writer.charts import (ChartWriter,
                                    PieChartWriter,
                                    LineChartWriter,
                                    BarChartWriter,
                                    ScatterChartWriter
                                    )
from openpyxl.workbook import Workbook
from openpyxl.chart import (Chart,
                            BarChart,
                            ScatterChart,
                            Serie,
                            Reference,
                            PieChart,
                            LineChart
                            )
from openpyxl.style import Color
from openpyxl.drawing import Image

from .schema import chart_schema, fromstring

# Class Fixtures

@pytest.fixture
def Chart():
    from openpyxl.chart import Chart
    return Chart


@pytest.fixture
def GraphChart():
    from openpyxl.chart import GraphChart
    return GraphChart


@pytest.fixture
def Axis():
    from openpyxl.chart import Axis
    return Axis


@pytest.fixture
def PieChart():
    from openpyxl.chart import PieChart
    return PieChart


@pytest.fixture
def LineChart():
    from openpyxl.chart import LineChart
    return LineChart


@pytest.fixture
def BarChart():
    from openpyxl.chart import BarChart
    return BarChart


@pytest.fixture
def ScatterChart():
    from openpyxl.chart import ScatterChart
    return ScatterChart


@pytest.fixture
def Reference():
    from openpyxl.chart import Reference
    return Reference


@pytest.fixture
def Serie():
    from openpyxl.chart import Serie
    return Serie


@pytest.fixture
def ErrorBar():
    from openpyxl.chart import ErrorBar
    return ErrorBar

@pytest.fixture
def less_than_one():
    from openpyxl.chart import less_than_one
    return less_than_one


def test_less_than_one(less_than_one):
    mul = less_than_one(1)
    assert mul == None

    mul = less_than_one(0.9)
    assert mul == 10.0

    mul = less_than_one(0.09)
    assert mul == 100.0

    mul = less_than_one(-0.09)
    assert mul == 100.0


def test_safe_string():
    from openpyxl.writer.charts import safe_string
    v = safe_string('s')
    assert v == 's'

    v = safe_string(2.0/3)
    assert v == '0.666666666666667'

    v = safe_string(1)
    assert v == '1'

    v = safe_string(None)
    assert v == 'None'


class TestAxis(object):


    def test_scaling(self, Axis):
        axis = Axis()
        axis.max = 10
        assert axis.min == 0.0
        assert axis.max == 12.0
        assert axis.unit == 2.0

        axis.max = 5
        assert axis.min == 0.0
        assert axis.max == 6.0
        assert axis.unit == 1.0

        axis.max = 50000
        assert axis.min == 0.0
        assert axis.max == 60000.0
        assert axis.unit == 12000.0

        axis.max = 1
        assert axis.min == 0.0
        assert axis.max == 2.0
        assert axis.unit == 1.0

        axis.max = 0.9
        assert axis.min == 0.0
        assert axis.max == 1.0
        assert axis.unit == 0.2

        axis.max = 0.09
        assert axis.min== 0.0
        assert axis.max== 0.1
        assert axis.unit == 0.02

        axis.min = -0.09
        axis.max = 0
        assert axis.min == -0.1
        assert axis.max == 0.0
        assert axis.unit == 0.02

        axis.min = -2
        axis.max = 8
        assert axis.min == -3.0
        assert axis.max == 10.0
        assert axis.unit == 2.0


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


class TestReference(object):

    def test_single_cell_ctor(self, cell):
        assert cell.pos1 == (0, 0)
        assert cell.pos2 == None

    def test_range_ctor(self, cell_range):
        assert cell_range.pos1 == (0, 0)
        assert cell_range.pos2 == (9, 0)

    def test_caching_cell(self, cell):
        assert cell._get_cache() == [0]

    def test_caching_range(self, cell_range):
        assert cell_range._get_cache() == [0, 1, 2, 3, 4, 5, 6, 7, 8 , 9]

    def test_ref_cell(self, cell):
        assert str(cell) == "'reference'!$A$1"

    def test_ref_range(self, cell_range):
        assert str(cell_range) == "'reference'!$A$1:$A$10"

    def test_data_type(self, cell, cell_range):
        with pytest.raises(ValueError):
            cell.data_type = 'f'
        assert cell.data_type == 'n'
        assert cell_range.data_type == 'n'

    def test_number_format(self, cell):
        with pytest.raises(ValueError):
            cell.number_format = 'YYYY'
        cell.number_format = 'd-mmm'


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
        assert series.legend == None

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

    #def setup(self):
        #wb = Workbook()
        #ws = wb.get_active_sheet()
        #for i in range(10):
            #ws.cell(row=i, column=0).value = 1
        #values = Reference(ws, (0, 0), (0, 9))
        #self.range = Serie(values=values)

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


@pytest.fixture
def bar_chart(ten_row_sheet, BarChart, Serie, Reference):
    ws = ten_row_sheet
    chart = BarChart()
    chart.title = "TITLE"
    series = Serie(Reference(ws, (0, 0), (10, 0)))
    series.color = Color.GREEN
    chart.add_serie(series)
    return chart


@pytest.fixture
def root_xml():
    return Element("test")


class TestChartWriter(object):

    def test_write_title(self, bar_chart, root_xml):
        cw = BarChartWriter(bar_chart)
        cw._write_title(root_xml)
        expected = """<?xml version='1.0' ?><test xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:title><c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr /></a:pPr><a:r><a:rPr lang="en-GB" /><a:t>TITLE</a:t></a:r></a:p></c:rich></c:tx><c:layout /></c:title></test>"""
        xml = get_xml(root_xml)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_xaxis(self, bar_chart, root_xml):
        cw = BarChartWriter(bar_chart)
        cw._write_axis(root_xml, bar_chart.x_axis, '{%s}catAx' % CHART_NS)
        expected = """<?xml version='1.0' ?><test xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:catAx><c:axId val="60871424" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:axPos val="b" /><c:tickLblPos val="nextTo" /><c:crossAx val="60873344" /><c:crosses val="autoZero" /><c:auto val="1" /><c:lblAlgn val="ctr" /><c:lblOffset val="100" /></c:catAx></test>"""
        xml = get_xml(root_xml)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_yaxis(self, bar_chart, root_xml):
        cw = BarChartWriter(bar_chart)
        cw._write_axis(root_xml, bar_chart.y_axis, '{%s}valAx' % CHART_NS)
        expected = """<?xml version='1.0' ?><test xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:valAx><c:axId val="60873344" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="l" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:tickLblPos val="nextTo" /><c:crossAx val="60871424" /><c:crosses val="autoZero" /><c:crossBetween val="between" /><c:majorUnit val="2.0" /></c:valAx></test>"""
        xml = get_xml(root_xml)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_series(self, bar_chart, root_xml):
        cw = BarChartWriter(bar_chart)
        cw._write_series(root_xml)
        expected = """<?xml version='1.0' ?><test xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:ser><c:idx val="0" /><c:order val="0" /><c:spPr><a:solidFill><a:srgbClr val="00FF00" /></a:solidFill><a:ln><a:solidFill><a:srgbClr val="00FF00" /></a:solidFill></a:ln></c:spPr><c:val><c:numRef><c:f>\'data\'!$A$1:$A$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></test>"""
        xml = get_xml(root_xml)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_legend(self, bar_chart, root_xml):
        cw = BarChartWriter(bar_chart)
        cw._write_legend(root_xml)
        expected = """<?xml version='1.0' ?><test xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:legend><c:legendPos val="r" /><c:layout /></c:legend></test>"""
        xml = get_xml(root_xml)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_no_write_legend(self, bar_chart, root_xml):
        cw = BarChartWriter(bar_chart)
        bar_chart.show_legend = False
        cw._write_legend(root_xml)
        children = [e for e in root_xml]
        assert len(children) == 0

    def test_write_print_settings(self, root_xml):
        tagnames = ['test',
                    '{%s}printSettings' % CHART_NS,
                    '{%s}headerFooter' % CHART_NS,
                    '{%s}pageMargins' % CHART_NS,
                    '{%s}pageSetup' % CHART_NS]
        for e in root_xml:
            assert_true(e.tag in tagnames)
            if e.tag == "{%s}pageMargins" % CHART_NS:
                assert e.keys() == list(self.chart.print_margins.keys())
                for k, v in e.items():
                    assert float(v) == self.chart.print_margins[k]
            else:
                assert e.text == None
                assert e.attrib == {}

    def test_write_chart(self, bar_chart):
        cw = BarChartWriter(bar_chart)
        root = Element('{%s}chartSpace' % CHART_NS)
        cw._write_chart(root)
        tree = fromstring(get_xml(root))
        assert chart_schema.validate(tree)

        expected = """<?xml version='1.0' ?><c:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:title><c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr /></a:pPr><a:r><a:rPr lang="en-GB" /><a:t>TITLE</a:t></a:r></a:p></c:rich></c:tx><c:layout /></c:title><c:plotArea><c:layout><c:manualLayout><c:layoutTarget val="inner" /><c:xMode val="edge" /><c:yMode val="edge" /><c:x val="0.03375" /><c:y val="0.31" /><c:w val="0.6" /><c:h val="0.6" /></c:manualLayout></c:layout><c:barChart><c:barDir val="col" /><c:grouping val="clustered" /><c:ser><c:idx val="0" /><c:order val="0" /><c:spPr><a:solidFill><a:srgbClr val="00FF00" /></a:solidFill><a:ln><a:solidFill><a:srgbClr val="00FF00" /></a:solidFill></a:ln></c:spPr><c:val><c:numRef><c:f>'data'!$A$1:$A$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val="60871424" /><c:axId val="60873344" /></c:barChart><c:catAx><c:axId val="60871424" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:axPos val="b" /><c:tickLblPos val="nextTo" /><c:crossAx val="60873344" /><c:crosses val="autoZero" /><c:auto val="1" /><c:lblAlgn val="ctr" /><c:lblOffset val="100" /></c:catAx><c:valAx><c:axId val="60873344" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="l" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:tickLblPos val="nextTo" /><c:crossAx val="60871424" /><c:crosses val="autoZero" /><c:crossBetween val="between" /><c:majorUnit val="2.0" /></c:valAx></c:plotArea><c:legend><c:legendPos val="r" /><c:layout /></c:legend><c:plotVisOnly val="1" /></c:chart></c:chartSpace>"""

        xml = get_xml(root)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_rels(self, bar_chart):
        cw = BarChartWriter(bar_chart)
        xml = cw.write_rels(1)
        expected = """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes" Target="../drawings/drawing1.xml"/></Relationships>"""
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_no_ascii(self, ten_row_sheet, Serie, BarChart, Reference):
        ws = ten_row_sheet
        ws.append(["D\xc3\xbcsseldorf"]*10)
        serie = Serie(values=Reference(ws, (0,0), (0,9)),
                      legend=Reference(ws, (1,0), (1,9))
                      )
        c = BarChart()
        c.add_serie(serie)
        cw = ChartWriter(c)

    def test_label_no_number_format(self, ten_column_sheet, Reference, Serie, BarChart, root_xml):
        ws = ten_column_sheet
        for i in range(10):
            ws.append([i, i])
        labels = Reference(ws, (0,0), (0,9))
        values = Reference(ws, (0,0), (0,9))
        serie = Serie(values=values, labels=labels)
        c = BarChart()
        c.add_serie(serie)
        cw = BarChartWriter(c)
        cw._write_serial(root_xml, c._series[0].labels)
        expected = """<?xml version='1.0' ?><test xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:numRef><c:f>'data'!$A$1:$J$1</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="10" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt></c:numCache></c:numRef></test>"""
        xml = get_xml(root_xml)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_label_number_format(self, ten_column_sheet, Reference, Serie, BarChart):
        ws = ten_column_sheet
        labels = Reference(ws, (0,0), (0,9))
        labels.number_format = 'd-mmm'
        values = Reference(ws, (0,0), (0,9))
        serie = Serie(values=values, labels=labels)
        c = BarChart()
        c.add_serie(serie)
        cw = BarChartWriter(c)
        root = Element('test')
        cw._write_serial(root, c._series[0].labels)

        expected = """<?xml version='1.0' ?><test xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:numRef><c:f>'data'!$A$1:$J$1</c:f><c:numCache><c:formatCode>d-mmm</c:formatCode><c:ptCount val="10" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt></c:numCache></c:numRef></test>"""

        xml = get_xml(root)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

@pytest.fixture()
def scatter_chart(ws, ScatterChart, Reference, Serie):
    ws.title = 'Scatter'
    for i in range(10):
        ws.cell(row=i, column=0).value = i
        ws.cell(row=i, column=1).value = i
    chart = ScatterChart()
    chart.add_serie(Serie(Reference(ws, (0, 0), (10, 0)),
                                      xvalues=Reference(ws, (0, 1), (10, 1))))
    return chart

class TestScatterChartWriter(object):

    #def setup(self, ws, ScatterChart, Reference, Serie):
        #ws.title = 'Scatter'
        #for i in range(10):
            #ws.cell(row=i, column=0).value = i
            #ws.cell(row=i, column=1).value = i
        #self.scatterchart = ScatterChart()
        #self.scatterchart.add_serie(Serie(Reference(ws, (0, 0), (10, 0)),
                                          #xvalues=Reference(ws, (0, 1), (10, 1))))
        #self.cw = ScatterChartWriter(self.scatterchart)
        #root_xml = Element('test')

    def test_write_xaxis(self, scatter_chart, root_xml):
        cw = ScatterChartWriter(scatter_chart)
        scatter_chart.x_axis.title = 'test x axis title'
        cw._write_axis(root_xml, scatter_chart.x_axis, '{%s}valAx' % CHART_NS)

        expected = """<test xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:valAx><c:axId val="60871424" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="b" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:title><c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr /></a:pPr><a:r><a:rPr lang="en-GB" /><a:t>test x axis title</a:t></a:r></a:p></c:rich></c:tx><c:layout /></c:title><c:tickLblPos val="nextTo" /><c:crossAx val="60873344" /><c:crosses val="autoZero" /><c:auto val="1" /><c:lblAlgn val="ctr" /><c:lblOffset val="100" /><c:crossBetween val="midCat" /><c:majorUnit val="2.0" /></c:valAx></test>"""
        xml = get_xml(root_xml)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_yaxis(self, scatter_chart, root_xml):
        cw = ScatterChartWriter(scatter_chart)
        scatter_chart.y_axis.title = 'test y axis title'
        cw._write_axis(root_xml, scatter_chart.y_axis, '{%s}valAx' % CHART_NS)

        expected = """<test xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:valAx><c:axId val="60873344" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="l" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:title><c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr /></a:pPr><a:r><a:rPr lang="en-GB" /><a:t>test y axis title</a:t></a:r></a:p></c:rich></c:tx><c:layout /></c:title><c:tickLblPos val="nextTo" /><c:crossAx val="60871424" /><c:crosses val="autoZero" /><c:crossBetween val="midCat" /><c:majorUnit val="2.0" /></c:valAx></test>"""
        xml = get_xml(root_xml)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_series(self, scatter_chart, root_xml):
        cw = ScatterChartWriter(scatter_chart)
        cw._write_series(root_xml)

        expected = """<test xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:ser><c:idx val="0" /><c:order val="0" /><c:xVal><c:numRef><c:f>\'Scatter\'!$B$1:$B$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:xVal><c:yVal><c:numRef><c:f>\'Scatter\'!$A$1:$A$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:yVal></c:ser></test>"""
        xml = get_xml(root_xml)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_legend(self, scatter_chart, root_xml):
        cw = ScatterChartWriter(scatter_chart)
        cw._write_legend(root_xml)
        expected = """<test xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:legend><c:legendPos val="r" /><c:layout /></c:legend></test>"""
        xml = get_xml(root_xml)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_print_settings(self, scatter_chart, root_xml):
        cw = ScatterChartWriter(scatter_chart)
        cw._write_print_settings(root_xml)

        expected = """<test xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:printSettings><c:headerFooter /><c:pageMargins b="0.75" footer="0.3" header="0.3" l="0.7" r="0.7" t="0.75" /><c:pageSetup /></c:printSettings></test>"""
        xml = get_xml(root_xml)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_chart(self, scatter_chart, root_xml):
        cw = ScatterChartWriter(scatter_chart)
        cw._write_chart(root_xml)
        xml = get_xml(root_xml)
        expected = """<test xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:layout><c:manualLayout><c:layoutTarget val="inner" /><c:xMode val="edge" /><c:yMode val="edge" /><c:x val="0.03375" /><c:y val="0.31" /><c:w val="0.6" /><c:h val="0.6" /></c:manualLayout></c:layout><c:scatterChart><c:scatterStyle val="lineMarker" /><c:ser><c:idx val="0" /><c:order val="0" /><c:xVal><c:numRef><c:f>\'Scatter\'!$B$1:$B$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:xVal><c:yVal><c:numRef><c:f>\'Scatter\'!$A$1:$A$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:yVal></c:ser><c:axId val="60871424" /><c:axId val="60873344" /></c:scatterChart><c:valAx><c:axId val="60871424" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="b" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:tickLblPos val="nextTo" /><c:crossAx val="60873344" /><c:crosses val="autoZero" /><c:auto val="1" /><c:lblAlgn val="ctr" /><c:lblOffset val="100" /><c:crossBetween val="midCat" /><c:majorUnit val="2.0" /></c:valAx><c:valAx><c:axId val="60873344" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="l" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:tickLblPos val="nextTo" /><c:crossAx val="60871424" /><c:crosses val="autoZero" /><c:crossBetween val="midCat" /><c:majorUnit val="2.0" /></c:valAx></c:plotArea><c:legend><c:legendPos val="r" /><c:layout /></c:legend><c:plotVisOnly val="1" /></c:chart></test>"""
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_serialised(self, scatter_chart):
        cw = ScatterChartWriter(scatter_chart)
        xml = cw.write()
        fname = os.path.join(DATADIR, "writer", "expected", "ScatterChart.xml")
        with open(fname) as expected:
            diff = compare_xml(xml, expected.read())
            assert diff is None, diff


@pytest.fixture
def pie_chart(ws, Reference, Serie, PieChart):
    ws.title = 'Pie'
    for i in range(1, 5):
        ws.append([i])
    chart = PieChart()
    values = Reference(ws, (0, 0), (9, 0))
    series = Serie(values, labels=values)
    chart.add_serie(series)
    return chart


class TestPieChartWriter(object):

    def test_write_chart(self, pie_chart, root_xml):
        """check if some characteristic tags of PieChart are there"""
        cw = PieChartWriter(pie_chart)
        cw._write_chart(root_xml)

        tagnames = ['{%s}pieChart' % CHART_NS,
                    '{%s}varyColors' % CHART_NS
                    ]
        root = safe_iterator(root_xml)
        chart_tags = [e.tag for e in root]
        for tag in tagnames:
            assert tag in chart_tags

        assert 'c:catAx' not in chart_tags

    def test_serialised(self, pie_chart):
        """Check the serialised file against sample"""
        cw = PieChartWriter(pie_chart)
        xml = cw.write()
        expected_file = os.path.join(DATADIR, "writer", "expected", "piechart.xml")
        with open(expected_file) as expected:
            diff = compare_xml(xml, expected.read())
            assert diff is None, diff


@pytest.fixture
def line_chart(ws, Reference, Serie, LineChart):
    ws.title = 'Line'
    for i in range(1, 5):
        ws.append([i])
    chart = LineChart()
    chart.add_serie(Serie(Reference(ws, (0, 0), (4, 0))))
    return chart


class TestLineChartWriter(object):

    def test_write_chart(self, line_chart, root_xml):
        """check if some characteristic tags of LineChart are there"""
        cw = LineChartWriter(line_chart)
        cw._write_chart(root_xml)
        tagnames = ['{%s}lineChart' % CHART_NS,
                    '{%s}valAx' % CHART_NS,
                    '{%s}catAx' % CHART_NS]

        root = safe_iterator(root_xml)
        chart_tags = [e.tag for e in root]
        for tag in tagnames:
            assert tag in chart_tags

    def test_serialised(self, line_chart):
        """Check the serialised file against sample"""
        cw = LineChartWriter(line_chart)
        xml = cw.write()
        expected_file = os.path.join(DATADIR, "writer", "expected", "LineChart.xml")
        with open(expected_file) as expected:
            diff = compare_xml(xml, expected.read())
            assert diff is None, diff


@pytest.fixture
def bar_chart_2(ws, BarChart, Reference, Serie):
    ws.title = 'Numbers'
    for i in range(10):
        ws.append([i])
    chart = BarChart()
    chart.add_serie(Serie(Reference(ws, (0, 0), (9, 0))))
    return chart


class TestBarChartWriter(object):

    def test_write_chart(self, bar_chart_2, root_xml):
        """check if some characteristic tags of LineChart are there"""
        cw = BarChartWriter(bar_chart_2)
        cw._write_chart(root_xml)
        tagnames = ['{%s}barChart' % CHART_NS,
                    '{%s}valAx' % CHART_NS,
                    '{%s}catAx' % CHART_NS]
        root = safe_iterator(root_xml)
        chart_tags = [e.tag for e in root]
        for tag in tagnames:
            assert tag in chart_tags

    def test_serialised(self, bar_chart_2):
        """Check the serialised file against sample"""
        cw = BarChartWriter(bar_chart_2)
        xml = cw.write()
        expected_file = os.path.join(DATADIR, "writer", "expected", "BarChart.xml")
        with open(expected_file) as expected:
            diff = compare_xml(xml, expected.read())
            assert diff is None, diff


class TestAnchoring(object):
    def _get_dummy_class(self):
        class DummyImg(object):
            def __init__(self):
                self.size = (200, 200)

        class DummyImage(Image):
            def _import_image(self, img):
                return DummyImg()

        return DummyImage

    def test_cell_anchor(self, ws):
        assert ws.cell('A1').anchor == (0, 0)
        assert ws.cell('D32').anchor == (210, 620)

    def test_image_anchor(self, ws):
        DummyImage = self._get_dummy_class()
        cell = ws.cell('D32')
        img = DummyImage(None)
        img.anchor(cell)
        assert (img.drawing.top, img.drawing.left) == (620, 210)

    def test_image_end(self, ws):
        DummyImage = self._get_dummy_class()
        cell = ws.cell('A1')
        img = DummyImage(None)
        img.drawing.width, img.drawing.height = (50, 50)
        end = img.anchor(cell)
        assert end[1] == ('A', 3)
