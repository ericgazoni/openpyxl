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

from nose.tools import eq_, assert_raises, assert_true, assert_false

from openpyxl.tests.helper import (get_xml,
                                   assert_equals_string,
                                   TMPDIR,
                                   DATADIR,
                                   assert_equals_file_content,
                                   make_tmpdir
                                   )

from openpyxl.shared.xmltools import Element, get_document_content
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
from re import sub
from openpyxl.drawing import Image


def test_less_than_one():
    from openpyxl.chart import less_than_one
    mul = less_than_one(1)
    eq_(mul, None)

    mul = less_than_one(0.9)
    eq_(mul, 10.0)

    mul = less_than_one(0.09)
    eq_(mul, 100.0)

    mul = less_than_one(-0.09)
    eq_(mul, 100.0)


def test_safe_string():
    from openpyxl.writer.charts import safe_string
    v = safe_string('s')
    eq_(v, 's')

    v = safe_string(2.0/3)
    eq_(v, '0.666666666666667')

    v = safe_string(1)
    eq_(v, '1')

    v = safe_string(None)
    eq_(v, 'None')


class TestAxis(object):

    def setUp(self):
        from openpyxl.chart import Axis
        self.axis = Axis()

    def test_scaling(self):
        self.axis.max = 10
        eq_(self.axis.min, 0.0)
        eq_(self.axis.max, 12.0)
        eq_(self.axis.unit, 2.0)

        self.axis.max = 5
        eq_(self.axis.min, 0.0)
        eq_(self.axis.max, 6.0)
        eq_(self.axis.unit, 1.0)

        self.axis.max = 50000
        eq_(self.axis.min, 0.0)
        eq_(self.axis.max, 60000.0)
        eq_(self.axis.unit, 12000.0)

        self.axis.max = 1
        eq_(self.axis.min, 0.0)
        eq_(self.axis.max, 2.0)
        eq_(self.axis.unit, 1.0)

        self.axis.max = 0.9
        eq_(self.axis.min, 0.0)
        eq_(self.axis.max, 1.0)
        eq_(self.axis.unit, 0.2)

        self.axis.max = 0.09
        eq_(self.axis.min, 0.0)
        eq_(self.axis.max, 0.1)
        eq_(self.axis.unit, 0.02)

        self.axis.min = -0.09
        self.axis.max = 0
        eq_(self.axis.min, -0.1)
        eq_(self.axis.max, 0.0)
        eq_(self.axis.unit, 0.02)

        self.axis.min = -2
        self.axis.max = 8
        eq_(self.axis.min, -3.0)
        eq_(self.axis.max, 10.0)
        eq_(self.axis.unit, 2.0)



class TestReference(object):

    def setUp(self):

        wb = Workbook()
        ws = wb.get_active_sheet()
        ws.title = 'reference'
        for i in range(10):
            ws.cell(row=i, column=0).value = i
        self.sheet = ws
        self.cell = Reference(self.sheet, (0, 0))
        self.range = Reference(self.sheet, (0, 0), (9, 0))

    def test_single_cell_ctor(self):
        eq_(self.cell.pos1, (0, 0))
        eq_(self.cell.pos2, None)

    def test_range_ctor(self):
        eq_(self.range.pos1, (0, 0))
        eq_(self.range.pos2, (9, 0))

    def test_type_validation(self):
        pass

    def test_caching_cell(self):
        eq_(self.cell._get_cache(), [0])

    def test_caching_range(self):
        eq_(self.range._get_cache(), [0, 1, 2, 3, 4, 5, 6, 7, 8 , 9])

    def test_ref_cell(self):
        eq_(str(self.cell), "'reference'!$A$1")
        eq_(self.cell._get_ref(), "'reference'!$A$1")

    def test_ref_range(self):
        eq_(str(self.range), "'reference'!$A$1:$A$10")
        eq_(self.range._get_ref(), "'reference'!$A$1:$A$10")

    def test_data_type(self):
        assert_raises(ValueError, setattr, self.cell, 'data_type', 'f')
        eq_(self.cell.data_type, 'n')
        eq_(self.range.data_type, 'n')

    def test_number_format(self):
        assert_raises(ValueError, setattr, self.cell, 'number_format', 'YYYY')
        self.cell.number_format = 'd-mmm'


class TestErrorBar(object):

    def setUp(self):
        wb = Workbook()
        ws = wb.get_active_sheet()
        for i in range(10):
            ws.cell(row=i, column=0).value = i
        self.range = Reference(ws, (0, 0), (9, 0))

    def test_ctor(self):
        from openpyxl.chart import ErrorBar
        assert_raises(TypeError, ErrorBar, None, range(10))


class TestSerie(object):

    def setUp(self):
        wb = Workbook()
        ws = wb.get_active_sheet()
        for i in range(10):
            ws.cell(row=i, column=0).value = i
        for i in range(10):
            ws.cell(row=i, column=1).value = None
        self.cell = Reference(ws, (0, 0))
        self.range = Reference(ws, (0, 0), (9, 0))
        self.empty = Reference(ws, (0, 1), (9, 1))

    def test_ctor(self):
        series = Serie(self.cell)
        eq_(series.values, [0])
        eq_(series.color, None)
        eq_(series.error_bar, None)
        eq_(series.xvalues, None)
        eq_(series.labels, None)
        eq_(series.legend, None)

    def test_color(self):
        series = Serie(self.cell)
        eq_(series.color, None)
        series.color = "blue"
        eq_(series.color, "blue")
        assert_raises(ValueError, setattr, series, 'color', None)

    def test_min_max(self):
        series = Serie(self.cell)
        eq_(series.get_min_max(), (0, 0))
        series = Serie(self.range)
        eq_(series.get_min_max(), (0, 9))
        series = Serie(self.empty)
        eq_(series.get_min_max(), (None, None))

    def test_min(self):
        series = Serie(self.cell)
        eq_(series.min(), 0)
        series = Serie(self.range)
        eq_(series.min(), 0)
        series = Serie(self.empty)
        eq_(series.min(), None)

    def test_max(self):
        series = Serie(self.cell)
        eq_(series.max(), 0)
        series = Serie(self.range)
        eq_(series.max(), 9)
        series = Serie(self.empty)
        eq_(series.max(), None)

    def test_len(self):
        series = Serie(self.cell)
        eq_(len(series), 1)

    def test_error_bar(self):
        series = Serie(self.cell)
        from openpyxl.chart import ErrorBar
        series.error_bar = ErrorBar(None, self.cell)
        eq_(series.get_min_max(), (0, 0))


class TestChart(object):

    def setUp(self):
        wb = Workbook()
        ws = wb.get_active_sheet()
        for i in range(10):
            ws.cell(row=i, column=0).value = 1
        values = Reference(ws, (0, 0), (0, 9))
        self.range = Serie(values=values)

    def make_worksheet(self):
        wb = Workbook()
        return wb.get_active_sheet()

    def test_ctor(self):
        from openpyxl.chart import Axis, Legend
        from openpyxl.drawing import Drawing
        c = Chart(None, None)
        eq_(c.type, None)
        eq_(c.grouping, None)
        assert_true(isinstance(c.x_axis, Axis))
        assert_true(isinstance(c.y_axis, Axis))
        assert_true(isinstance(c.legend, Legend))
        eq_(c.show_legend, True)
        eq_(c.lang, 'en-GB')
        eq_(c.title, '')
        eq_(c.print_margins,
            {'b':.75, 'l':.7, 'r':.7, 't':.75, 'header':0.3, 'footer':.3}
            )
        assert_true(isinstance(c.drawing, Drawing))
        eq_(c.width, .6)
        eq_(c.height, .6)
        eq_(c.margin_top, 0.31)
        #eq_(c.margin_left, 0)
        eq_(c._shapes, [])

    def test_mymax(self):
        c = Chart(None, None)
        eq_(c.mymax(range(10)), 9)
        from string import ascii_letters as letters
        eq_(c.mymax(list(letters)), "z")
        eq_(c.mymax(range(-10, 1)), 0)
        eq_(c.mymax([""]*10), "")

    def test_mymin(self):
        c = Chart(None, None)
        eq_(c.mymin(range(10)), 0)
        from string import ascii_letters as letters
        eq_(c.mymin(list(letters)), "A")
        eq_(c.mymin(range(-10, 1)), -10)
        eq_(c.mymin([""]*10), "")

    def test_get_x_unit(self):
        c = Chart(None, None)
        c._series.append(self.range)
        eq_(c.get_x_units(), 10)

    def test_get_y_unit(self):
        c = Chart(None, None)
        c._series.append(self.range)
        c.y_axis.max = 10
        eq_(c.get_y_units(), 190500.0)

    def test_get_y_char(self):
        c = Chart(None, None)
        c._series.append(self.range)
        eq_(c.get_y_chars(), 1)

    def test_compute_series_extremes(self):
        c = Chart(None, None)
        c._series.append(self.range)
        mini, maxi = c._get_extremes()
        eq_(mini, 0)
        eq_(maxi, 1.0)

    def test_compute_series_max_dates(self):
        ws = self.make_worksheet()
        for i in range(1, 10):
            ws.append([date(2013, i, 1)])
        c = Chart(None, None)
        ref = Reference(ws, (0, 0), (9, 0))
        series = Serie(ref)
        c._series.append(series)
        mini, maxi = c._get_extremes()
        eq_(mini, 0)
        eq_(maxi, 41518.0)

    def test_margin_top(self):
        c = Chart(None, None)
        eq_(c.margin_top, 0.31)

    def test_margin_left(self):
        c = Chart(None, None)
        c._series.append(self.range)
        eq_(c.margin_left, 0.03375)

    def test_set_margin_top(self):
        c = Chart(None, None)
        c.margin_top = 1
        eq_(c.margin_top, 0.31)

    def test_set_margin_left(self):
        c = Chart(None, None)
        c._series.append(self.range)
        c.margin_left = 0
        eq_(c.margin_left , 0.03375)


class TestLineChart(object):

    def test_ctor(self):
        from openpyxl.chart import LineChart
        c = LineChart()
        eq_(c.type, Chart.LINE_CHART)


class TestChartWriter(object):

    def setUp(self):

        wb = Workbook()
        ws = wb.get_active_sheet()
        ws.title = 'data'
        for i in range(10):
            ws.cell(row=i, column=0).value = i
        self.chart = BarChart()
        self.chart.title = 'TITLE'
        self.chart.add_serie(Serie(Reference(ws, (0, 0), (10, 0))))
        self.chart._series[-1].color = Color.GREEN
        self.cw = BarChartWriter(self.chart)
        self.root = Element('test')

    def make_worksheet(self):

        wb = Workbook()
        ws = wb.get_active_sheet()
        ws.title = 'data'
        ws.append(list(range(10)))
        return ws

    def test_write_title(self):
        self.cw._write_title(self.root)
        expected = """<?xml version='1.0' encoding='UTF-8'?><test><c:title><c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr /></a:pPr><a:r><a:rPr lang="en-GB" /><a:t>TITLE</a:t></a:r></a:p></c:rich></c:tx><c:layout /></c:title></test>"""
        assert_equals_string(get_xml(self.root), expected)

    def test_write_xaxis(self):

        self.cw._write_axis(self.root, self.chart.x_axis, 'c:catAx')
        expected = """<?xml version='1.0' encoding='UTF-8'?><test><c:catAx><c:axId val="60871424" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:axPos val="b" /><c:tickLblPos val="nextTo" /><c:crossAx val="60873344" /><c:crosses val="autoZero" /><c:auto val="1" /><c:lblAlgn val="ctr" /><c:lblOffset val="100" /></c:catAx></test>"""
        assert_equals_string(get_xml(self.root), expected)

    def test_write_yaxis(self):

        self.cw._write_axis(self.root, self.chart.y_axis, 'c:valAx')
        expected = """<?xml version='1.0' encoding='UTF-8'?><test><c:valAx><c:axId val="60873344" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="l" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:tickLblPos val="nextTo" /><c:crossAx val="60871424" /><c:crosses val="autoZero" /><c:crossBetween val="between" /><c:majorUnit val="2.0" /></c:valAx></test>"""
        assert_equals_string(get_xml(self.root), expected)

    def test_write_series(self):

        self.cw._write_series(self.root)
        expected = """<?xml version='1.0' encoding='UTF-8'?><test><c:ser><c:idx val="0" /><c:order val="0" /><c:spPr><a:solidFill><a:srgbClr val="00FF00" /></a:solidFill><a:ln><a:solidFill><a:srgbClr val="00FF00" /></a:solidFill></a:ln></c:spPr><c:val><c:numRef><c:f>\'data\'!$A$1:$A$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></test>"""
        assert_equals_string(get_xml(self.root), expected)

    def test_write_legend(self):

        self.cw._write_legend(self.root)
        eq_(get_xml(self.root), """<?xml version='1.0' encoding='UTF-8'?><test><c:legend><c:legendPos val="r" /><c:layout /></c:legend></test>""")

    def test_no_write_legend(self):

        self.chart.show_legend = False
        self.cw._write_legend(self.root)
        children = [e for e in self.root]
        eq_(len(children), 0)

    def test_write_print_settings(self):
        tagnames = ['test', 'c:printSettings', 'c:headerFooter',
                    'c:pageMargins', 'c:pageSetup']
        self.cw._write_print_settings(self.root)
        for e in self.root.iter():
            assert_true(e.tag in tagnames, "%s not found" % e.tag)
            if e.tag == "c:pageMargins":
                eq_(e.keys(), list(self.chart.print_margins.keys()))
                for k, v in e.items():
                    eq_(float(v), self.chart.print_margins[k])
            else:
                eq_(e.text, None)
                eq_(e.attrib, {})

    def test_write_chart(self):
        from openpyxl.namespaces import CHART_NS, A_NS, REL_NS
        from .schema import chart_schema, fromstring
        CHART_NS.update(A_NS)
        root = Element('c:chartSpace', CHART_NS)
        self.cw._write_chart(root)
        tree = fromstring(get_xml(root))
        assert_true(chart_schema.validate(tree))

        # Truncate floats because results differ with Python >= 3.2 and <= 3.1
        expected = """<?xml version='1.0' encoding='UTF-8'?><c:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:title><c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr /></a:pPr><a:r><a:rPr lang="en-GB" /><a:t>TITLE</a:t></a:r></a:p></c:rich></c:tx><c:layout /></c:title><c:plotArea><c:layout><c:manualLayout><c:layoutTarget val="inner" /><c:xMode val="edge" /><c:yMode val="edge" /><c:x val="0.0337" /><c:y val="0.31" /><c:w val="0.6" /><c:h val="0.6" /></c:manualLayout></c:layout><c:barChart><c:barDir val="col" /><c:grouping val="clustered" /><c:ser><c:idx val="0" /><c:order val="0" /><c:spPr><a:solidFill><a:srgbClr val="00FF00" /></a:solidFill><a:ln><a:solidFill><a:srgbClr val="00FF00" /></a:solidFill></a:ln></c:spPr><c:val><c:numRef><c:f>'data'!$A$1:$A$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val="60871424" /><c:axId val="60873344" /></c:barChart><c:catAx><c:axId val="60871424" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:axPos val="b" /><c:tickLblPos val="nextTo" /><c:crossAx val="60873344" /><c:crosses val="autoZero" /><c:auto val="1" /><c:lblAlgn val="ctr" /><c:lblOffset val="100" /></c:catAx><c:valAx><c:axId val="60873344" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="l" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:tickLblPos val="nextTo" /><c:crossAx val="60871424" /><c:crosses val="autoZero" /><c:crossBetween val="between" /><c:majorUnit val="2.0" /></c:valAx></c:plotArea><c:legend><c:legendPos val="r" /><c:layout /></c:legend><c:plotVisOnly val="1" /></c:chart></c:chartSpace>"""


        test_xml = sub('([0-9][.][0-9]{4})[0-9]*', '\\1', get_xml(root))
        eq_(test_xml, expected)

    def test_write_no_ascii(self):

        ws = self.make_worksheet()
        ws.append([u"D\xfcsseldorf"]*10)
        serie = Serie(values=Reference(ws, (0,0), (0,9)),
                      legend=Reference(ws, (1,0), (1,9))
                      )
        c = BarChart()
        c.add_serie(serie)
        cw = ChartWriter(c)

    def test_label_no_number_format(self):
        ws = self.make_worksheet()
        for i in range(10):
            ws.append([i, i])
        labels = Reference(ws, (0,0), (0,9))
        values = Reference(ws, (0,0), (0,9))
        serie = Serie(values=values, labels=labels)
        c = BarChart()
        c.add_serie(serie)
        cw = BarChartWriter(c)
        root = Element('test')
        cw._write_serial(root, c._series[0].labels)
        xml = get_xml(root)
        eq_(xml, """<?xml version='1.0' encoding='UTF-8'?><test><c:numRef><c:f>'data'!$A$1:$J$1</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="10" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt></c:numCache></c:numRef></test>""")


    def test_label_number_format(self):
        ws = self.make_worksheet()
        for i in range(10):
            ws.append([i, i])
        labels = Reference(ws, (0,0), (0,9))
        labels.number_format = 'd-mmm'
        values = Reference(ws, (0,0), (0,9))
        serie = Serie(values=values, labels=labels)
        c = BarChart()
        c.add_serie(serie)
        cw = BarChartWriter(c)
        root = Element('test')
        cw._write_serial(root, c._series[0].labels)
        xml = get_xml(root)
        eq_(xml, """<?xml version='1.0' encoding='UTF-8'?><test><c:numRef><c:f>'data'!$A$1:$J$1</c:f><c:numCache><c:formatCode>d-mmm</c:formatCode><c:ptCount val="10" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt></c:numCache></c:numRef></test>""")



class TestScatterChartWriter(object):

    def setUp(self):

        wb = Workbook()
        ws = wb.get_active_sheet()
        ws.title = 'data'
        for i in range(10):
            ws.cell(row=i, column=0).value = i
            ws.cell(row=i, column=1).value = i
        self.scatterchart = ScatterChart()
        self.scatterchart.add_serie(Serie(Reference(ws, (0, 0), (10, 0)),
                                          xvalues=Reference(ws, (0, 1), (10, 1))))
        self.cw = ScatterChartWriter(self.scatterchart)
        self.root = Element('test')

    def test_write_xaxis(self):

        self.scatterchart.x_axis.title = 'test x axis title'
        self.cw._write_axis(self.root, self.scatterchart.x_axis, 'c:valAx')
        expected = """<?xml version='1.0' encoding='UTF-8'?><test><c:valAx><c:axId val="60871424" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="b" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:title><c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr /></a:pPr><a:r><a:rPr lang="en-GB" /><a:t>test x axis title</a:t></a:r></a:p></c:rich></c:tx><c:layout /></c:title><c:tickLblPos val="nextTo" /><c:crossAx val="60873344" /><c:crosses val="autoZero" /><c:auto val="1" /><c:lblAlgn val="ctr" /><c:lblOffset val="100" /><c:crossBetween val="midCat" /><c:majorUnit val="2.0" /></c:valAx></test>"""
        eq_(get_xml(self.root), expected)

    def test_write_yaxis(self):

        self.scatterchart.y_axis.title = 'test y axis title'
        self.cw._write_axis(self.root, self.scatterchart.y_axis, 'c:valAx')
        expected = """<?xml version='1.0' encoding='UTF-8'?><test><c:valAx><c:axId val="60873344" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="l" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:title><c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr /></a:pPr><a:r><a:rPr lang="en-GB" /><a:t>test y axis title</a:t></a:r></a:p></c:rich></c:tx><c:layout /></c:title><c:tickLblPos val="nextTo" /><c:crossAx val="60871424" /><c:crosses val="autoZero" /><c:crossBetween val="midCat" /><c:majorUnit val="2.0" /></c:valAx></test>"""
        eq_(get_xml(self.root), expected)

    def test_write_series(self):

        self.cw._write_series(self.root)
        expected = """<?xml version='1.0' encoding='UTF-8'?><test><c:ser><c:idx val="0" /><c:order val="0" /><c:xVal><c:numRef><c:f>\'data\'!$B$1:$B$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:xVal><c:yVal><c:numRef><c:f>\'data\'!$A$1:$A$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:yVal></c:ser></test>"""
        eq_(get_xml(self.root), expected)

    def test_write_legend(self):

        self.cw._write_legend(self.root)
        eq_(get_xml(self.root), """<?xml version='1.0' encoding='UTF-8'?><test><c:legend><c:legendPos val="r" /><c:layout /></c:legend></test>""")

    def test_write_print_settings(self):

        self.cw._write_print_settings(self.root)
        eq_(get_xml(self.root), """<?xml version='1.0' encoding='UTF-8'?><test><c:printSettings><c:headerFooter /><c:pageMargins b="0.75" footer="0.3" header="0.3" l="0.7" r="0.7" t="0.75" /><c:pageSetup /></c:printSettings></test>""")

    def test_write_chart(self):

        self.cw._write_chart(self.root)
        xml = get_document_content(self.root)
        # Truncate floats because results differ with Python >= 3.2 and <= 3.1
        xml = sub('([0-9][.][0-9]{4})[0-9]*', '\\1', xml)
        expected = """<?xml version="1.0" encoding="UTF-8"?>
<test>
  <c:chart>
    <c:plotArea>
      <c:layout>
        <c:manualLayout>
          <c:layoutTarget val="inner" />
          <c:xMode val="edge" />
          <c:yMode val="edge" />
          <c:x val="0.0337" />
          <c:y val="0.31" />
          <c:w val="0.6" />
          <c:h val="0.6" />
        </c:manualLayout>
      </c:layout>
      <c:scatterChart>
        <c:scatterStyle val="lineMarker" />
        <c:ser>
          <c:idx val="0" />
          <c:order val="0" />
          <c:xVal>
            <c:numRef>
              <c:f>'data'!$B$1:$B$11</c:f>
              <c:numCache>
                <c:formatCode>General</c:formatCode>
                <c:ptCount val="11" />
                <c:pt idx="0">
                  <c:v>0</c:v>
                </c:pt>
                <c:pt idx="1">
                  <c:v>1</c:v>
                </c:pt>
                <c:pt idx="2">
                  <c:v>2</c:v>
                </c:pt>
                <c:pt idx="3">
                  <c:v>3</c:v>
                </c:pt>
                <c:pt idx="4">
                  <c:v>4</c:v>
                </c:pt>
                <c:pt idx="5">
                  <c:v>5</c:v>
                </c:pt>
                <c:pt idx="6">
                  <c:v>6</c:v>
                </c:pt>
                <c:pt idx="7">
                  <c:v>7</c:v>
                </c:pt>
                <c:pt idx="8">
                  <c:v>8</c:v>
                </c:pt>
                <c:pt idx="9">
                  <c:v>9</c:v>
                </c:pt>
                <c:pt idx="10">
                  <c:v>None</c:v>
                </c:pt>
              </c:numCache>
            </c:numRef>
          </c:xVal>
          <c:yVal>
            <c:numRef>
              <c:f>'data'!$A$1:$A$11</c:f>
              <c:numCache>
                <c:formatCode>General</c:formatCode>
                <c:ptCount val="11" />
                <c:pt idx="0">
                  <c:v>0</c:v>
                </c:pt>
                <c:pt idx="1">
                  <c:v>1</c:v>
                </c:pt>
                <c:pt idx="2">
                  <c:v>2</c:v>
                </c:pt>
                <c:pt idx="3">
                  <c:v>3</c:v>
                </c:pt>
                <c:pt idx="4">
                  <c:v>4</c:v>
                </c:pt>
                <c:pt idx="5">
                  <c:v>5</c:v>
                </c:pt>
                <c:pt idx="6">
                  <c:v>6</c:v>
                </c:pt>
                <c:pt idx="7">
                  <c:v>7</c:v>
                </c:pt>
                <c:pt idx="8">
                  <c:v>8</c:v>
                </c:pt>
                <c:pt idx="9">
                  <c:v>9</c:v>
                </c:pt>
                <c:pt idx="10">
                  <c:v>None</c:v>
                </c:pt>
              </c:numCache>
            </c:numRef>
          </c:yVal>
        </c:ser>
        <c:axId val="60871424" />
        <c:axId val="60873344" />
      </c:scatterChart>
      <c:valAx>
        <c:axId val="60871424" />
        <c:scaling>
          <c:orientation val="minMax" />
          <c:max val="10.0" />
          <c:min val="0.0" />
        </c:scaling>
        <c:axPos val="b" />
        <c:majorGridlines />
        <c:numFmt formatCode="General" sourceLinked="1" />
        <c:tickLblPos val="nextTo" />
        <c:crossAx val="60873344" />
        <c:crosses val="autoZero" />
        <c:auto val="1" />
        <c:lblAlgn val="ctr" />
        <c:lblOffset val="100" />
        <c:crossBetween val="midCat" />
        <c:majorUnit val="2.0" />
      </c:valAx>
      <c:valAx>
        <c:axId val="60873344" />
        <c:scaling>
          <c:orientation val="minMax" />
          <c:max val="10.0" />
          <c:min val="0.0" />
        </c:scaling>
        <c:axPos val="l" />
        <c:majorGridlines />
        <c:numFmt formatCode="General" sourceLinked="1" />
        <c:tickLblPos val="nextTo" />
        <c:crossAx val="60871424" />
        <c:crosses val="autoZero" />
        <c:crossBetween val="midCat" />
        <c:majorUnit val="2.0" />
      </c:valAx>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r" />
      <c:layout />
    </c:legend>
    <c:plotVisOnly val="1" />
  </c:chart>
</test>"""
        #eq_(test_xml, expected)
        assert_equals_string(xml, expected)


    def test_serialised(self):
        return
        xml = self.cw.write()
        fname = os.path.join(DATADIR, "writer", "expected", "ScatterChart.xml")
        with open(fname) as expected:
            assert_equals_string(xml, expected.read())


class TestPieChartWriter(object):

    def setUp(self):
        """Setup a worksheet with one column of data and a pie chart"""
        wb = Workbook()
        ws = wb.get_active_sheet()
        ws.title = 'Pie'
        for i in range(1, 5):
            ws.append([i])
        self.piechart = PieChart()
        values = Reference(ws, (0, 0), (9, 0))
        series = Serie(values, labels=values)
        self.piechart.add_serie(series)
        ws.add_chart(self.piechart)
        self.cw = PieChartWriter(self.piechart)
        self.root = Element('test')

    def test_write_chart(self):
        """check if some characteristic tags of PieChart are there"""
        self.cw._write_chart(self.root)
        tagnames = ['test', 'c:pieChart', 'c:varyColors']
        chart_tags = [e.tag for e in self.root.iter()]
        for tag in tagnames:
            assert_true(tag in chart_tags, tag)

        assert_false('c:catAx' in chart_tags)

    def test_serialised(self):
        """Check the serialised file against sample"""
        xml = self.cw.write()
        pie = open(os.path.join(DATADIR, 'writer', 'expected', 'piechart.xml'))
        compare_xml = pie.read()
        pie.close()
        assert_equals_string(xml, compare_xml)


class TestLineChartWriter(object):

    def setUp(self):
        """Setup a worksheet with one column of data and a line chart"""
        wb = Workbook()
        ws = wb.get_active_sheet()
        ws.title = 'Line'
        for i in range(1, 5):
            ws.append([i])
        self.piechart = LineChart()
        self.piechart.add_serie(Serie(Reference(ws, (0, 0), (4, 0))))
        self.cw = LineChartWriter(self.piechart)
        self.root = Element('test')

    def test_write_chart(self):
        """check if some characteristic tags of LineChart are there"""
        self.cw._write_chart(self.root)
        tagnames = ['test', 'c:lineChart', 'c:valAx', 'c:catAx']
        chart_tags = [e.tag for e in self.root.iter()]
        for tag in tagnames:
            assert_true(tag in chart_tags, tag)

    def test_serialised(self):
        """Check the serialised file against sample"""
        xml = self.cw.write()
        expected_file = os.path.join(DATADIR, "writer", "expected", "LineChart.xml")
        with open(expected_file) as expected:
            assert_equals_string(xml, expected.read())


class TestBarChartWriter(object):
    """"""
    def setUp(self):
        """Setup a worksheet with one column of data and a bar chart"""
        wb = Workbook()
        ws = wb.get_active_sheet()
        ws.title = 'Numbers'
        for i in range(10):
            ws.append([i])
        self.piechart = BarChart()
        self.piechart.add_serie(Serie(Reference(ws, (0, 0), (9, 0))))
        self.cw = BarChartWriter(self.piechart)
        self.root = Element('test')

    def test_write_chart(self):
        """check if some characteristic tags of LineChart are there"""
        self.cw._write_chart(self.root)
        tagnames = ['test', 'c:barChart', 'c:valAx', 'c:catAx']
        chart_tags = [e.tag for e in self.root.iter()]
        for tag in tagnames:
            assert_true(tag in chart_tags, tag)

    def test_serialised(self):
        """Check the serialised file against sample"""
        xml = self.cw.write()
        expected_file = os.path.join(DATADIR, "writer", "expected", "BarChart.xml")
        with open(expected_file) as expected:
            assert_equals_string(xml, expected.read())


class TestAnchoring(object):
    def _get_dummy_class(self):
        class DummyImg(object):
            def __init__(self):
                self.size = (200, 200)

        class DummyImage(Image):
            def _import_image(self, img):
                return DummyImg()

        return DummyImage

    def test_cell_anchor(self):
        wb = Workbook()
        ws = wb.get_active_sheet()

        eq_(ws.cell('A1').anchor, (0, 0))
        eq_(ws.cell('D32').anchor, (210, 620))

    def test_image_anchor(self):
        DummyImage = self._get_dummy_class()
        wb = Workbook()
        ws = wb.get_active_sheet()
        cell = ws.cell('D32')
        img = DummyImage(None)
        img.anchor(cell)
        eq_((img.drawing.top, img.drawing.left), (620, 210))

    def test_image_end(self):
        DummyImage = self._get_dummy_class()
        wb = Workbook()
        ws = wb.get_active_sheet()
        cell = ws.cell('A1')
        img = DummyImage(None)
        img.drawing.width, img.drawing.height = (50, 50)
        end = img.anchor(cell)
        eq_(end[1], ('A', 3))