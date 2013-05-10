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

from nose.tools import eq_, assert_raises, assert_true

from openpyxl.tests.helper import get_xml
from openpyxl.shared.xmltools import Element
from openpyxl.writer.charts import ChartWriter
from openpyxl.workbook import Workbook
from openpyxl.chart import Chart, BarChart, ScatterChart, Serie, Reference
from openpyxl.style import Color
from re import sub
from openpyxl.drawing import Image


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

    def test_get_type(self):
        self.cell.data_type = 'n'
        eq_(self.cell.get_type(), 'n')

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
        self.cell = Reference(ws, (0, 0))
        self.range = Reference(ws, (0, 0), (9, 0))

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
        self.range = Reference(ws, (0, 0), (0, 9))

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
        eq_(c.lang, 'fr-FR')
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
        from string import letters
        eq_(c.mymax(list(letters)), "z")
        eq_(c.mymax(range(-10, 1)), 0)
        eq_(c.mymax([""]*10), "")

    def test_get_x_unit(self):
        c = Chart(None, None)
        c._series.append(self.range)
        eq_(c.get_x_units(), 10)

    def test_get_y_unit(self):
        c = Chart(None, None)
        c._series.append(self.range)
        c.y_axis.max = 10
        eq_(c.get_y_units(), 228600.0)

    def test_get_y_char(self):
        c = Chart(None, None)
        c._series.append(self.range)
        eq_(c.get_y_chars(), 1)

    def test_compute_series_max_numbers(self):
        c = Chart(None, None)
        c._series.append(self.range)
        maxi, unit = c._compute_series_max()
        eq_(maxi, 2.0)
        eq_(unit, 1.0)

    def test_compute_series_max_dates(self):
        ws = self.make_worksheet()
        for i in range(1, 10):
            ws.append([date(2013, i, 1)])
        c = Chart(None, None)
        ref = Reference(ws, (0, 0), (9, 0))
        c._series.append(ref)
        maxi, unit = c._compute_series_max()

    def test_computer_series_max_strings(self):
        ws = self.make_worksheet()
        for i in range(10):
            ws.append(['a'])
        ref = Reference(ws, (0, 0), (9, 0))
        c = Chart(None, None)
        c._series.append(ref)
        #maxi, unit = c._compute_series_max()

    def test_less_than_one(self):
        from openpyxl.chart import less_than_one
        mul = less_than_one(1)
        eq_(mul, None)

        mul = less_than_one(0.9)
        eq_(mul, 10.0)

        mul = less_than_one(0.09)
        eq_(mul, 100.0)

        mul = less_than_one(-0.09)
        eq_(mul, 100.0)

    def test_scale_axis(self):
        from openpyxl.chart import scale_axis
        maxi, unit = scale_axis(10)
        eq_(maxi, 12.0)
        eq_(unit, 2.0)
        assert_true(maxi/unit < 10)

        maxi, unit = scale_axis(5)
        eq_(maxi, 6.0)
        eq_(unit, 1.0)
        assert_true(maxi/unit < 10)

        maxi, unit = scale_axis(50000)
        eq_(maxi, 60000.0)
        eq_(unit, 12000.0)
        assert_true(maxi/unit < 10)

        maxi, unit = scale_axis(1)
        eq_(maxi, 2.0)
        eq_(unit, 1.0)

        maxi, unit = scale_axis(0.9)
        eq_(maxi, 1.0)
        eq_(unit, 0.2)

        maxi, unit = scale_axis(0.09)
        eq_(maxi, 0.1)
        eq_(unit, 0.02)

        maxi, unit = scale_axis(-0.09)
        eq_(maxi, 0.1)
        eq_(unit, 0.02)

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
        self.cw = ChartWriter(self.chart)
        self.root = Element('test')

    def make_worksheet(self):

        wb = Workbook()
        ws = wb.get_active_sheet()
        ws.title = 'data'
        ws.append(range(10))
        return ws

    def test_write_title(self):
        self.cw._write_title(self.root)
        eq_(get_xml(self.root), '<?xml version=\'1.0\' encoding=\'UTF-8\'?><test><c:title><c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr /></a:pPr><a:r><a:rPr lang="fr-FR" /><a:t>TITLE</a:t></a:r></a:p></c:rich></c:tx><c:layout /></c:title></test>')

    def test_write_xaxis(self):

        self.cw._write_axis(self.root, self.chart.x_axis, 'c:catAx')
        eq_(get_xml(self.root), '<?xml version=\'1.0\' encoding=\'UTF-8\'?><test><c:catAx><c:axId val="60871424" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:axPos val="b" /><c:tickLblPos val="nextTo" /><c:crossAx val="60873344" /><c:crosses val="autoZero" /><c:auto val="1" /><c:lblAlgn val="ctr" /><c:lblOffset val="100" /></c:catAx></test>')

    def test_write_yaxis(self):

        self.cw._write_axis(self.root, self.chart.y_axis, 'c:valAx')
        eq_(get_xml(self.root), '<?xml version=\'1.0\' encoding=\'UTF-8\'?><test><c:valAx><c:axId val="60873344" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="l" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:tickLblPos val="nextTo" /><c:crossAx val="60871424" /><c:crosses val="autoZero" /><c:crossBetween val="between" /><c:majorUnit val="2.0" /></c:valAx></test>')

    def test_write_series(self):

        self.cw._write_series(self.root)
        eq_(get_xml(self.root), '<?xml version=\'1.0\' encoding=\'UTF-8\'?><test><c:ser><c:idx val="0" /><c:order val="0" /><c:spPr><a:solidFill><a:srgbClr val="00FF00" /></a:solidFill><a:ln><a:solidFill><a:srgbClr val="00FF00" /></a:solidFill></a:ln></c:spPr><c:marker><c:symbol val="none" /></c:marker><c:val><c:numRef><c:f>\'data\'!$A$1:$A$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></test>')

    def test_write_legend(self):

        self.cw._write_legend(self.root)
        eq_(get_xml(self.root), '<?xml version=\'1.0\' encoding=\'UTF-8\'?><test><c:legend><c:legendPos val="r" /><c:layout /></c:legend></test>')

    def test_no_write_legend(self):

        wb = Workbook()
        ws = wb.get_active_sheet()
        ws.title = 'data'
        for i in range(10):
            ws.cell(row=i, column=0).value = i
            ws.cell(row=i, column=1).value = i
        scatterchart = ScatterChart()
        scatterchart.add_serie(Serie(Reference(ws, (0, 0), (10, 0)),
                         xvalues=Reference(ws, (0, 1), (10, 1))))
        cw = ChartWriter(scatterchart)
        root = Element('test')
        scatterchart.show_legend = False
        cw._write_legend(root)
        eq_(get_xml(root), '<?xml version=\'1.0\' encoding=\'UTF-8\'?><test />')

    def test_write_print_settings(self):

        self.cw._write_print_settings(self.root)
        eq_(get_xml(self.root), '<?xml version=\'1.0\' encoding=\'UTF-8\'?><test><c:printSettings><c:headerFooter /><c:pageMargins b="0.75" footer="0.3" header="0.3" l="0.7" r="0.7" t="0.75" /><c:pageSetup /></c:printSettings></test>')

    def test_write_chart(self):

        self.cw._write_chart(self.root)
        # Truncate floats because results differ with Python >= 3.2 and <= 3.1
        test_xml = sub('([0-9][.][0-9]{4})[0-9]*', '\\1', get_xml(self.root))
        eq_(test_xml, '<?xml version=\'1.0\' encoding=\'UTF-8\'?><test><c:chart><c:title><c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr /></a:pPr><a:r><a:rPr lang="fr-FR" /><a:t>TITLE</a:t></a:r></a:p></c:rich></c:tx><c:layout /></c:title><c:plotArea><c:layout><c:manualLayout><c:layoutTarget val="inner" /><c:xMode val="edge" /><c:yMode val="edge" /><c:x val="0.0337" /><c:y val="0.31" /><c:w val="0.6" /><c:h val="0.6" /></c:manualLayout></c:layout><c:barChart><c:barDir val="col" /><c:grouping val="clustered" /><c:ser><c:idx val="0" /><c:order val="0" /><c:spPr><a:solidFill><a:srgbClr val="00FF00" /></a:solidFill><a:ln><a:solidFill><a:srgbClr val="00FF00" /></a:solidFill></a:ln></c:spPr><c:marker><c:symbol val="none" /></c:marker><c:val><c:numRef><c:f>\'data\'!$A$1:$A$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:marker val="1" /><c:axId val="60871424" /><c:axId val="60873344" /></c:barChart><c:catAx><c:axId val="60871424" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:axPos val="b" /><c:tickLblPos val="nextTo" /><c:crossAx val="60873344" /><c:crosses val="autoZero" /><c:auto val="1" /><c:lblAlgn val="ctr" /><c:lblOffset val="100" /></c:catAx><c:valAx><c:axId val="60873344" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="l" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:tickLblPos val="nextTo" /><c:crossAx val="60871424" /><c:crosses val="autoZero" /><c:crossBetween val="between" /><c:majorUnit val="2.0" /></c:valAx></c:plotArea><c:legend><c:legendPos val="r" /><c:layout /></c:legend><c:plotVisOnly val="1" /></c:chart></test>')

    def test_write_no_ascii(self):

        ws = self.make_worksheet()
        ws.append([u"D\xfcsseldorf"]*10)
        serie = Serie(values=Reference(ws, (0,0), (0,9)),
                      legend=Reference(ws, (1,0), (1,9))
                      )
        c = BarChart()
        c.add_serie(serie)
        cw = ChartWriter(c)


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
        self.cw = ChartWriter(self.scatterchart)
        self.root = Element('test')

    def test_write_xaxis(self):

        self.scatterchart.x_axis.title = 'test x axis title'
        self.cw._write_axis(self.root, self.scatterchart.x_axis, 'c:valAx')
        eq_(get_xml(self.root), '<?xml version=\'1.0\' encoding=\'UTF-8\'?><test><c:valAx><c:axId val="60871424" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="b" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:title><c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr /></a:pPr><a:r><a:rPr lang="fr-FR" /><a:t>test x axis title</a:t></a:r></a:p></c:rich></c:tx><c:layout /></c:title><c:tickLblPos val="nextTo" /><c:crossAx val="60873344" /><c:crosses val="autoZero" /><c:auto val="1" /><c:lblAlgn val="ctr" /><c:lblOffset val="100" /><c:crossBetween val="midCat" /><c:majorUnit val="2.0" /></c:valAx></test>')

    def test_write_yaxis(self):

        self.scatterchart.y_axis.title = 'test y axis title'
        self.cw._write_axis(self.root, self.scatterchart.y_axis, 'c:valAx')
        eq_(get_xml(self.root), '<?xml version=\'1.0\' encoding=\'UTF-8\'?><test><c:valAx><c:axId val="60873344" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="l" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:title><c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr /></a:pPr><a:r><a:rPr lang="fr-FR" /><a:t>test y axis title</a:t></a:r></a:p></c:rich></c:tx><c:layout /></c:title><c:tickLblPos val="nextTo" /><c:crossAx val="60871424" /><c:crosses val="autoZero" /><c:crossBetween val="midCat" /><c:majorUnit val="2.0" /></c:valAx></test>')

    def test_write_series(self):

        self.cw._write_series(self.root)
        eq_(get_xml(self.root), '<?xml version=\'1.0\' encoding=\'UTF-8\'?><test><c:ser><c:idx val="0" /><c:order val="0" /><c:marker><c:symbol val="none" /></c:marker><c:xVal><c:numRef><c:f>\'data\'!$B$1:$B$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:xVal><c:yVal><c:numRef><c:f>\'data\'!$A$1:$A$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:yVal></c:ser></test>')

    def test_write_legend(self):

        self.cw._write_legend(self.root)
        eq_(get_xml(self.root), '<?xml version=\'1.0\' encoding=\'UTF-8\'?><test><c:legend><c:legendPos val="r" /><c:layout /></c:legend></test>')

    def test_write_print_settings(self):

        self.cw._write_print_settings(self.root)
        eq_(get_xml(self.root), '<?xml version=\'1.0\' encoding=\'UTF-8\'?><test><c:printSettings><c:headerFooter /><c:pageMargins b="0.75" footer="0.3" header="0.3" l="0.7" r="0.7" t="0.75" /><c:pageSetup /></c:printSettings></test>')

    def test_write_chart(self):

        self.cw._write_chart(self.root)
        # Truncate floats because results differ with Python >= 3.2 and <= 3.1
        test_xml = sub('([0-9][.][0-9]{4})[0-9]*', '\\1', get_xml(self.root))
        eq_(test_xml, '<?xml version=\'1.0\' encoding=\'UTF-8\'?><test><c:chart><c:plotArea><c:layout><c:manualLayout><c:layoutTarget val="inner" /><c:xMode val="edge" /><c:yMode val="edge" /><c:x val="0.0337" /><c:y val="0.31" /><c:w val="0.6" /><c:h val="0.6" /></c:manualLayout></c:layout><c:scatterChart><c:scatterStyle val="lineMarker" /><c:ser><c:idx val="0" /><c:order val="0" /><c:marker><c:symbol val="none" /></c:marker><c:xVal><c:numRef><c:f>\'data\'!$B$1:$B$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:xVal><c:yVal><c:numRef><c:f>\'data\'!$A$1:$A$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:yVal></c:ser><c:marker val="1" /><c:axId val="60871424" /><c:axId val="60873344" /></c:scatterChart><c:valAx><c:axId val="60871424" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="b" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:tickLblPos val="nextTo" /><c:crossAx val="60873344" /><c:crosses val="autoZero" /><c:auto val="1" /><c:lblAlgn val="ctr" /><c:lblOffset val="100" /><c:crossBetween val="midCat" /><c:majorUnit val="2.0" /></c:valAx><c:valAx><c:axId val="60873344" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="l" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:tickLblPos val="nextTo" /><c:crossAx val="60871424" /><c:crosses val="autoZero" /><c:crossBetween val="midCat" /><c:majorUnit val="2.0" /></c:valAx></c:plotArea><c:legend><c:legendPos val="r" /><c:layout /></c:legend><c:plotVisOnly val="1" /></c:chart></test>')


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
