import pytest
import os

from openpyxl.shared.xmltools import Element, fromstring, safe_iterator
from openpyxl.shared.ooxml import CHART_NS

from openpyxl.writer.charts import (ChartWriter,
                                    PieChartWriter,
                                    LineChartWriter,
                                    BarChartWriter,
                                    ScatterChartWriter,
                                    BaseChartWriter
                                    )
from openpyxl.styles import Color

from .helper import get_xml, DATADIR, compare_xml
from .schema import chart_schema

@pytest.fixture
def bar_chart(ten_row_sheet, BarChart, Series, Reference):
    ws = ten_row_sheet
    chart = BarChart()
    chart.title = "TITLE"
    series = Series(Reference(ws, (0, 0), (10, 0)))
    series.color = Color.GREEN
    chart.add_serie(series)
    return chart


def test_write_serial(ten_row_sheet, LineChart, Series, Reference, root_xml):
    ws = ten_row_sheet
    chart = LineChart()
    for idx, l in enumerate("ABCDEF"):
        ws.cell(row=idx, column=0).value = l
    ref = Reference(ws, (0, 0), (9, 0))
    series = Series(ref)
    chart.add_serie(series)
    cw = BaseChartWriter(chart)
    cw._write_serial(cw.root, ref)
    xml = get_xml(cw.root)
    expected = """ <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:strRef><c:f>'data'!$A$1:$A$10</c:f><c:strCache><c:ptCount val="10"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt><c:pt idx="3"><c:v>D</c:v></c:pt><c:pt idx="4"><c:v>E</c:v></c:pt><c:pt idx="5"><c:v>F</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt></c:strCache></c:strRef></c:chartSpace>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


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

    def test_write_print_settings(self, bar_chart):
        cw = BarChartWriter(bar_chart)
        cw._write_print_settings()
        tagnames = ['test',
                    '{%s}printSettings' % CHART_NS,
                    '{%s}headerFooter' % CHART_NS,
                    '{%s}pageMargins' % CHART_NS,
                    '{%s}pageSetup' % CHART_NS]
        for e in cw.root:
            assert e.tag in tagnames
            if e.tag == "{%s}pageMargins" % CHART_NS:
                assert e.keys() == list(bar_chart.print_margins.keys())
                for k, v in e.items():
                    assert float(v) == bar_chart.print_margins[k]
            else:
                assert e.text == None
                assert e.attrib == {}

    @pytest.mark.lxml_required
    def test_write_chart(self, bar_chart):
        cw = BarChartWriter(bar_chart)
        cw._write_chart()
        assert chart_schema.validate(cw.root)

        expected = """<c:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:title><c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr /></a:pPr><a:r><a:rPr lang="en-GB" /><a:t>TITLE</a:t></a:r></a:p></c:rich></c:tx><c:layout /></c:title><c:plotArea><c:layout><c:manualLayout><c:layoutTarget val="inner" /><c:xMode val="edge" /><c:yMode val="edge" /><c:x val="0.03375" /><c:y val="0.31" /><c:w val="0.6" /><c:h val="0.6" /></c:manualLayout></c:layout><c:barChart><c:barDir val="col" /><c:grouping val="clustered" /><c:ser><c:idx val="0" /><c:order val="0" /><c:spPr><a:solidFill><a:srgbClr val="00FF00" /></a:solidFill><a:ln><a:solidFill><a:srgbClr val="00FF00" /></a:solidFill></a:ln></c:spPr><c:val><c:numRef><c:f>'data'!$A$1:$A$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val="60871424" /><c:axId val="60873344" /></c:barChart><c:catAx><c:axId val="60871424" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:axPos val="b" /><c:tickLblPos val="nextTo" /><c:crossAx val="60873344" /><c:crosses val="autoZero" /><c:auto val="1" /><c:lblAlgn val="ctr" /><c:lblOffset val="100" /></c:catAx><c:valAx><c:axId val="60873344" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="l" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:tickLblPos val="nextTo" /><c:crossAx val="60871424" /><c:crosses val="autoZero" /><c:crossBetween val="between" /><c:majorUnit val="2.0" /></c:valAx></c:plotArea><c:legend><c:legendPos val="r" /><c:layout /></c:legend><c:plotVisOnly val="1" /></c:chart></c:chartSpace>"""

        xml = get_xml(cw.root)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_rels(self, bar_chart):
        cw = BarChartWriter(bar_chart)
        xml = cw.write_rels(1)
        expected = """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes" Target="../drawings/drawing1.xml"/></Relationships>"""
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    @pytest.mark.xfail
    def test_write_no_ascii(self, ten_row_sheet, Series, BarChart, Reference):
        ws = ten_row_sheet
        ws.append(["D\xc3\xbcsseldorf"]*10)
        serie = Series(values=Reference(ws, (0,0), (0,9)),
                      title=(ws.cell(row=1, column=0).value)
                      )
        c = BarChart()
        c.add_serie(serie)
        cw = ChartWriter(c)

    def test_label_no_number_format(self, ten_column_sheet, Reference, Series, BarChart, root_xml):
        ws = ten_column_sheet
        for i in range(10):
            ws.append([i, i])
        labels = Reference(ws, (0,0), (0,9))
        values = Reference(ws, (0,0), (0,9))
        serie = Series(values=values, labels=labels)
        c = BarChart()
        c.add_serie(serie)
        cw = BarChartWriter(c)
        cw._write_serial(root_xml, c.series[0].labels)
        expected = """<test xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:numRef><c:f>'data'!$A$1:$J$1</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="10" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt></c:numCache></c:numRef></test>"""
        xml = get_xml(root_xml)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_label_number_format(self, ten_column_sheet, Reference, Series, BarChart):
        ws = ten_column_sheet
        labels = Reference(ws, (0,0), (0,9))
        labels.number_format = 'd-mmm'
        values = Reference(ws, (0,0), (0,9))
        serie = Series(values=values, labels=labels)
        c = BarChart()
        c.add_serie(serie)
        cw = BarChartWriter(c)
        root = Element('test')
        cw._write_serial(root, c.series[0].labels)

        expected = """<test xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:numRef><c:f>'data'!$A$1:$J$1</c:f><c:numCache><c:formatCode>d-mmm</c:formatCode><c:ptCount val="10" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt></c:numCache></c:numRef></test>"""

        xml = get_xml(root)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

@pytest.fixture()
def scatter_chart(ws, ScatterChart, Reference, Series):
    ws.title = 'Scatter'
    for i in range(10):
        ws.cell(row=i, column=0).value = i
        ws.cell(row=i, column=1).value = i
    chart = ScatterChart()
    chart.add_serie(Series(Reference(ws, (0, 0), (10, 0)),
                                      xvalues=Reference(ws, (0, 1), (10, 1))))
    return chart


class TestScatterChartWriter(object):

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

    def test_write_print_settings(self, scatter_chart):
        cw = ScatterChartWriter(scatter_chart)
        cw._write_print_settings()

        expected = """<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:printSettings><c:headerFooter /><c:pageMargins b="0.75" footer="0.3" header="0.3" l="0.7" r="0.7" t="0.75" /><c:pageSetup /></c:printSettings></c:chartSpace>"""
        xml = get_xml(cw.root)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_chart(self, scatter_chart):
        cw = ScatterChartWriter(scatter_chart)
        cw._write_chart()
        xml = get_xml(cw.root)
        expected = """<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:layout><c:manualLayout><c:layoutTarget val="inner" /><c:xMode val="edge" /><c:yMode val="edge" /><c:x val="0.03375" /><c:y val="0.31" /><c:w val="0.6" /><c:h val="0.6" /></c:manualLayout></c:layout><c:scatterChart><c:scatterStyle val="lineMarker" /><c:ser><c:idx val="0" /><c:order val="0" /><c:xVal><c:numRef><c:f>\'Scatter\'!$B$1:$B$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:xVal><c:yVal><c:numRef><c:f>\'Scatter\'!$A$1:$A$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:yVal></c:ser><c:axId val="60871424" /><c:axId val="60873344" /></c:scatterChart><c:valAx><c:axId val="60871424" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="b" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:tickLblPos val="nextTo" /><c:crossAx val="60873344" /><c:crosses val="autoZero" /><c:auto val="1" /><c:lblAlgn val="ctr" /><c:lblOffset val="100" /><c:crossBetween val="midCat" /><c:majorUnit val="2.0" /></c:valAx><c:valAx><c:axId val="60873344" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="l" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:tickLblPos val="nextTo" /><c:crossAx val="60871424" /><c:crosses val="autoZero" /><c:crossBetween val="midCat" /><c:majorUnit val="2.0" /></c:valAx></c:plotArea><c:legend><c:legendPos val="r" /><c:layout /></c:legend><c:plotVisOnly val="1" /></c:chart></c:chartSpace>"""
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
def pie_chart(ws, Reference, Series, PieChart):
    ws.title = 'Pie'
    for i in range(1, 5):
        ws.append([i])
    chart = PieChart()
    values = Reference(ws, (0, 0), (9, 0))
    series = Series(values, labels=values)
    chart.add_serie(series)
    return chart


class TestPieChartWriter(object):

    def test_write_chart(self, pie_chart):
        """check if some characteristic tags of PieChart are there"""
        cw = PieChartWriter(pie_chart)
        cw._write_chart()

        tagnames = ['{%s}pieChart' % CHART_NS,
                    '{%s}varyColors' % CHART_NS
                    ]
        root = safe_iterator(cw.root)
        chart_tags = [e.tag for e in root]
        for tag in tagnames:
            assert tag in chart_tags

        assert 'c:catAx' not in chart_tags

    @pytest.mark.lxml_required
    def test_serialised(self, pie_chart):
        """Check the serialised file against sample"""
        cw = PieChartWriter(pie_chart)
        xml = cw.write()
        tree = fromstring(xml)
        chart_schema.assertValid(tree)
        expected_file = os.path.join(DATADIR, "writer", "expected", "piechart.xml")
        with open(expected_file) as expected:
            diff = compare_xml(xml, expected.read())
            assert diff is None, diff


@pytest.fixture
def line_chart(ws, Reference, Series, LineChart):
    ws.title = 'Line'
    for i in range(1, 5):
        ws.append([i])
    chart = LineChart()
    chart.add_serie(Series(Reference(ws, (0, 0), (4, 0))))
    return chart


class TestLineChartWriter(object):

    def test_write_chart(self, line_chart):
        """check if some characteristic tags of LineChart are there"""
        cw = LineChartWriter(line_chart)
        cw._write_chart()
        tagnames = ['{%s}lineChart' % CHART_NS,
                    '{%s}valAx' % CHART_NS,
                    '{%s}catAx' % CHART_NS]

        root = safe_iterator(cw.root)
        chart_tags = [e.tag for e in root]
        for tag in tagnames:
            assert tag in chart_tags

    @pytest.mark.lxml_required
    def test_serialised(self, line_chart):
        """Check the serialised file against sample"""
        cw = LineChartWriter(line_chart)
        xml = cw.write()
        tree = fromstring(xml)
        chart_schema.assertValid(tree)
        expected_file = os.path.join(DATADIR, "writer", "expected", "LineChart.xml")
        with open(expected_file) as expected:
            diff = compare_xml(xml, expected.read())
            assert diff is None, diff


@pytest.fixture
def bar_chart_2(ws, BarChart, Reference, Series):
    ws.title = 'Numbers'
    for i in range(10):
        ws.append([i])
    chart = BarChart()
    chart.add_serie(Series(Reference(ws, (0, 0), (9, 0))))
    return chart


class TestBarChartWriter(object):

    def test_write_chart(self, bar_chart_2):
        """check if some characteristic tags of LineChart are there"""
        cw = BarChartWriter(bar_chart_2)
        cw._write_chart()
        tagnames = ['{%s}barChart' % CHART_NS,
                    '{%s}valAx' % CHART_NS,
                    '{%s}catAx' % CHART_NS]
        root = safe_iterator(cw.root)
        chart_tags = [e.tag for e in root]
        for tag in tagnames:
            assert tag in chart_tags

    @pytest.mark.lxml_required
    def test_serialised(self, bar_chart_2):
        """Check the serialised file against sample"""
        cw = BarChartWriter(bar_chart_2)
        xml = cw.write()
        tree = fromstring(xml)
        chart_schema.assertValid(tree)
        expected_file = os.path.join(DATADIR, "writer", "expected", "BarChart.xml")
        with open(expected_file) as expected:
            diff = compare_xml(xml, expected.read())
            assert diff is None, diff
