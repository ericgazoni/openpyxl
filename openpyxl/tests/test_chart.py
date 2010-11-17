'''
Copyright (c) 2010 openpyxl

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.

@license: http://www.opensource.org/licenses/mit-license.php
@author: Eric Gazoni
'''

from nose.tools import eq_

from openpyxl.tests.helper import get_xml
from openpyxl.shared.xmltools import Element
from openpyxl.writer.charts import ChartWriter
from openpyxl.workbook import Workbook
from openpyxl.chart import BarChart, Serie, Reference

class TestChartWriter(object):
    
    def setUp(self):
        
        wb = Workbook()
        ws = wb.get_active_sheet()
        ws.title = u'data'
        for i in range(10):
            ws.cell(row=i, column=0).value = i
        self.chart = BarChart()
        self.chart.add_serie(Serie(Reference(ws, (0,0), (10,0))))
        self.cw = ChartWriter(self.chart)
        self.root = Element('test')
        
    def test_write_xaxis(self):
        
        self.cw._write_axis(self.root, self.chart.x_axis, 'c:catAx')
        eq_(get_xml(self.root), '<?xml version=\'1.0\' encoding=\'UTF-8\'?><test><c:catAx><c:axId val="60871424" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:axPos val="b" /><c:tickLblPos val="nextTo" /><c:crossAx val="60873344" /><c:crosses val="autoZero" /><c:auto val="1" /><c:lblAlgn val="ctr" /><c:lblOffset val="100" /></c:catAx></test>')
        
    def test_write_yaxis(self):
        
        self.cw._write_axis(self.root, self.chart.y_axis, 'c:valAx')
        eq_(get_xml(self.root), '<?xml version=\'1.0\' encoding=\'UTF-8\'?><test><c:valAx><c:axId val="60873344" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0" /></c:scaling><c:axPos val="l" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:tickLblPos val="nextTo" /><c:crossAx val="60871424" /><c:crosses val="autoZero" /><c:crossBetween val="between" /><c:majorUnit val="2.0" /></c:valAx></test>')

    def test_write_series(self):
        
        self.cw._write_series(self.root)
        eq_(get_xml(self.root), '<?xml version=\'1.0\' encoding=\'UTF-8\'?><test><c:ser><c:idx val="0" /><c:order val="0" /><c:marker><c:symbol val="none" /></c:marker><c:val><c:numRef><c:f>data!$A$1:$A$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></test>')
        
    def test_write_legend(self):
        
        self.cw._write_legend(self.root)
        eq_(get_xml(self.root), '<?xml version=\'1.0\' encoding=\'UTF-8\'?><test><c:legend><c:legendPos val="r" /><c:layout /></c:legend></test>')
        
    def test_write_print_settings(self):
        
        self.cw._write_print_settings(self.root)
        eq_(get_xml(self.root), '<?xml version=\'1.0\' encoding=\'UTF-8\'?><test><c:printSettings><c:headerFooter /><c:pageMargins b="0.75" footer="0.3" header="0.3" l="0.7" r="0.7" t="0.75" /><c:pageSetup /></c:printSettings></test>')
        
    def test_write_chart(self):
        
        self.cw._write_chart(self.root)
        eq_(get_xml(self.root), '<?xml version=\'1.0\' encoding=\'UTF-8\'?><test><c:chart><c:plotArea><c:layout><c:manualLayout><c:layoutTarget val="inner" /><c:xMode val="edge" /><c:yMode val="edge" /><c:x val="1.28571428571" /><c:y val="0.2125" /><c:w val="0.6" /><c:h val="0.6" /></c:manualLayout></c:layout><c:barChart><c:barDir val="col" /><c:grouping val="clustered" /><c:ser><c:idx val="0" /><c:order val="0" /><c:marker><c:symbol val="none" /></c:marker><c:val><c:numRef><c:f>data!$A$1:$A$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:marker val="1" /><c:axId val="60871424" /><c:axId val="60873344" /></c:barChart><c:catAx><c:axId val="60871424" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:axPos val="b" /><c:tickLblPos val="nextTo" /><c:crossAx val="60873344" /><c:crosses val="autoZero" /><c:auto val="1" /><c:lblAlgn val="ctr" /><c:lblOffset val="100" /></c:catAx><c:valAx><c:axId val="60873344" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0" /></c:scaling><c:axPos val="l" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:tickLblPos val="nextTo" /><c:crossAx val="60871424" /><c:crosses val="autoZero" /><c:crossBetween val="between" /><c:majorUnit val="2.0" /></c:valAx></c:plotArea><c:legend><c:legendPos val="r" /><c:layout /></c:legend><c:plotVisOnly val="1" /></c:chart></test>')