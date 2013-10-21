# coding=UTF-8
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

from numbers import Number

from openpyxl.shared.xmltools import Element, SubElement, get_document_content
from openpyxl.shared.ooxml import CHART_NS, DRAWING_NS, REL_NS
from openpyxl.shared.compat.itertools import iteritems
from openpyxl.chart import Chart, ErrorBar, BarChart, LineChart, PieChart, ScatterChart

try:
    # Python 2
    basestring
except NameError:
    # Python 3
    basestring = str


def safe_string(value):
    """Safely and consistently format numeric values"""
    if isinstance(value, Number):
        value = "%.15g" % value
    elif not isinstance(value, basestring):
        value = str(value)
    return value

class BaseChartWriter(object):

    def __init__(self, chart):
        self.chart = chart

    def write(self):
        """ write a chart """
        root = Element("{%s}chartSpace" % CHART_NS)

        SubElement(root, '{%s}lang' % CHART_NS, {'val':self.chart.lang})
        self._write_chart(root)
        self._write_print_settings(root)
        self._write_shapes(root)

        return get_document_content(root)

    def _write_chart(self, root):
        ch = SubElement(root, '{%s}chart' % CHART_NS)
        self._write_title(ch)
        self._write_layout(ch)
        self._write_legend(ch)
        SubElement(ch, '{%s}plotVisOnly' % CHART_NS, {'val':'1'})

    def _write_layout(self, element):
        chart = self.chart
        plot_area = SubElement(element, '{%s}plotArea' % CHART_NS)
        layout = SubElement(plot_area, '{%s}layout' % CHART_NS)
        mlayout = SubElement(layout, '{%s}manualLayout' % CHART_NS)
        SubElement(mlayout, '{%s}layoutTarget' % CHART_NS, {'val':'inner'})
        SubElement(mlayout, '{%s}xMode' % CHART_NS, {'val':'edge'})
        SubElement(mlayout, '{%s}yMode' % CHART_NS, {'val':'edge'})
        SubElement(mlayout, '{%s}x' % CHART_NS, {'val':safe_string(chart.margin_left)})
        SubElement(mlayout, '{%s}y' % CHART_NS, {'val':safe_string(chart.margin_top)})
        SubElement(mlayout, '{%s}w' % CHART_NS, {'val':safe_string(chart.width)})
        SubElement(mlayout, '{%s}h' % CHART_NS, {'val':safe_string(chart.height)})

        chart_type = self.chart.TYPE
        subchart = SubElement(plot_area, '{%s}%s' % (CHART_NS, chart_type))
        self._write_options(subchart)
        self._write_series(subchart)
        return subchart, plot_area

    def _write_options(self, subchart):
        pass

    def _write_title(self, chart):
        if self.chart.title != '':
            title = SubElement(chart, '{%s}title' % CHART_NS)
            tx = SubElement(title, '{%s}tx' % CHART_NS)
            rich = SubElement(tx, '{%s}rich' % CHART_NS)
            SubElement(rich, '{%s}bodyPr' % DRAWING_NS)
            SubElement(rich, '{%s}lstStyle' % DRAWING_NS)
            p = SubElement(rich, '{%s}p' % DRAWING_NS)
            pPr = SubElement(p, '{%s}pPr' % DRAWING_NS)
            SubElement(pPr, '{%s}defRPr' % DRAWING_NS)
            r = SubElement(p, '{%s}r' % DRAWING_NS)
            SubElement(r, '{%s}rPr' % DRAWING_NS, {'lang':self.chart.lang})
            t = SubElement(r, '{%s}t' % DRAWING_NS).text = self.chart.title
            SubElement(title, '{%s}layout' % CHART_NS)

    def _write_axis_title(self, axis, ax):

        if axis.title != '':
            title = SubElement(ax, '{%s}title' % CHART_NS)
            tx = SubElement(title, '{%s}tx' % CHART_NS)
            rich = SubElement(tx, '{%s}rich' % CHART_NS)
            SubElement(rich, '{%s}bodyPr' % DRAWING_NS)
            SubElement(rich, '{%s}lstStyle' % DRAWING_NS)
            p = SubElement(rich, '{%s}p' % DRAWING_NS)
            pPr = SubElement(p, '{%s}pPr' % DRAWING_NS)
            SubElement(pPr, '{%s}defRPr' % DRAWING_NS)
            r = SubElement(p, '{%s}r' % DRAWING_NS)
            SubElement(r, '{%s}rPr' % DRAWING_NS, {'lang':self.chart.lang})
            t = SubElement(r, '{%s}t' % DRAWING_NS).text = axis.title
            SubElement(title, '{%s}layout' % CHART_NS)

    def _write_axis(self, plot_area, axis, label):

        self.chart.compute_axes()

        ax = SubElement(plot_area, label)
        SubElement(ax, '{%s}axId' % CHART_NS, {'val':safe_string(axis.id)})

        scaling = SubElement(ax, '{%s}scaling' % CHART_NS)
        SubElement(scaling, '{%s}orientation' % CHART_NS, {'val':axis.orientation})
        if label == '{%s}valAx' % CHART_NS:
            SubElement(scaling, '{%s}max' % CHART_NS, {'val':str(float(axis.max))})
            SubElement(scaling, '{%s}min' % CHART_NS, {'val':str(float(axis.min))})

        SubElement(ax, '{%s}axPos' % CHART_NS, {'val':axis.position})
        if label == '{%s}valAx' % CHART_NS:
            SubElement(ax, '{%s}majorGridlines' % CHART_NS)
            SubElement(ax, '{%s}numFmt' % CHART_NS, {'formatCode':"General", 'sourceLinked':'1'})
        self._write_axis_title(axis, ax)
        SubElement(ax, '{%s}tickLblPos' % CHART_NS, {'val':axis.tick_label_position})
        SubElement(ax, '{%s}crossAx' % CHART_NS, {'val':str(axis.cross)})
        SubElement(ax, '{%s}crosses' % CHART_NS, {'val':axis.crosses})
        if axis.auto:
            SubElement(ax, '{%s}auto' % CHART_NS, {'val':'1'})
        if axis.label_align:
            SubElement(ax, '{%s}lblAlgn' % CHART_NS, {'val':axis.label_align})
        if axis.label_offset:
            SubElement(ax, '{%s}lblOffset' % CHART_NS, {'val':str(axis.label_offset)})
        if label == '{%s}valAx' % CHART_NS:
            if self.chart.TYPE == "scatterChart":
                SubElement(ax, '{%s}crossBetween' % CHART_NS, {'val':'midCat'})
            else:
                SubElement(ax, '{%s}crossBetween' % CHART_NS, {'val':'between'})
            SubElement(ax, '{%s}majorUnit' % CHART_NS, {'val':str(float(axis.unit))})

    def _write_series(self, subchart):

        for i, serie in enumerate(self.chart._series):
            ser = SubElement(subchart, '{%s}ser' % CHART_NS)
            SubElement(ser, '{%s}idx' % CHART_NS, {'val':safe_string(i)})
            SubElement(ser, '{%s}order' % CHART_NS, {'val':safe_string(i)})

            if serie.legend:
                tx = SubElement(ser, '{%s}tx' % CHART_NS)
                self._write_serial(tx, serie.legend)

            if serie.color:
                sppr = SubElement(ser, '{%s}spPr' % CHART_NS)
                if self.chart.TYPE == "barChart":
                    # fill color
                    fillc = SubElement(sppr, '{%s}solidFill' % DRAWING_NS)
                    SubElement(fillc, '{%s}srgbClr' % DRAWING_NS, {'val':serie.color})
                # edge color
                ln = SubElement(sppr, '{%s}ln' % DRAWING_NS)
                fill = SubElement(ln, '{%s}solidFill' % DRAWING_NS)
                SubElement(fill, '{%s}srgbClr' % DRAWING_NS, {'val':serie.color})

            if serie.error_bar:
                self._write_error_bar(ser, serie)

            if serie.labels:
                cat = SubElement(ser, '{%s}cat' % CHART_NS)
                self._write_serial(cat, serie.labels)

            if self.chart.TYPE == "scatterChart":
                if serie.xvalues:
                    xval = SubElement(ser, '{%s}xVal' % CHART_NS)
                    self._write_serial(xval, serie.xreference)

                val = SubElement(ser, '{%s}yVal' % CHART_NS)
            else:
                val = SubElement(ser, '{%s}val' % CHART_NS)
            self._write_serial(val, serie.reference)

    def _write_serial(self, node, reference, literal=False):

        is_ref = hasattr(reference, 'pos1')
        data_type = reference.data_type
        number_format = getattr(reference, 'number_format')

        mapping = {'n':{'ref':'numRef', 'cache':'numCache'},
                   's':{'ref':'strRef', 'cache':'strCache'}}

        if is_ref:
            ref = SubElement(node, '{%s}%s' %(CHART_NS, mapping[data_type]['ref']))
            SubElement(ref, '{%s}f' % CHART_NS).text = str(reference)
            data = SubElement(ref, '{%s}%s' %(CHART_NS, mapping[data_type]['cache']))
            values = reference.values
        else:
            data = SubElement(node, '{%s}numLit' % CHART_NS)
            values = (1,)

        if data_type == 'n':
            SubElement(data, '{%s}formatCode' % CHART_NS).text = number_format or 'General'

        SubElement(data, '{%s}ptCount' % CHART_NS, {'val':str(len(values))})
        for j, val in enumerate(values):
            point = SubElement(data, '{%s}pt' % CHART_NS, {'idx':str(j)})
            if not isinstance(val, basestring):
                val = safe_string(val)
            SubElement(point, '{%s}v' % CHART_NS).text = val

    def _write_error_bar(self, node, serie):

        flag = {ErrorBar.PLUS_MINUS:'both',
                ErrorBar.PLUS:'plus',
                ErrorBar.MINUS:'minus'}

        eb = SubElement(node, '{%s}errBars' % CHART_NS)
        SubElement(eb, '{%s}errBarType' % CHART_NS, {'val':flag[serie.error_bar.type]})
        SubElement(eb, '{%s}errValType' % CHART_NS, {'val':'cust'})

        plus = SubElement(eb, '{%s}plus' % CHART_NS)
        # apart from setting the type of data series the following has
        # no effect on the writer
        self._write_serial(plus, serie.error_bar.reference,
            literal=(serie.error_bar.type == ErrorBar.MINUS))

        minus = SubElement(eb, '{%s}minus' % CHART_NS)
        self._write_serial(minus, serie.error_bar.reference,
            literal=(serie.error_bar.type == ErrorBar.PLUS))

    def _write_legend(self, chart):

        if self.chart.show_legend:
            legend = SubElement(chart, '{%s}legend' % CHART_NS)
            SubElement(legend, '{%s}legendPos' % CHART_NS, {'val':self.chart.legend.position})
            SubElement(legend, '{%s}layout' % CHART_NS)

    def _write_print_settings(self, root):

        settings = SubElement(root, '{%s}printSettings' % CHART_NS)
        SubElement(settings, '{%s}headerFooter' % CHART_NS)
        try:
            # Python 2
            print_margins_items = iteritems(self.chart.print_margins)
        except AttributeError:
            # Python 3
            print_margins_items = self.chart.print_margins.items()

        margins = dict([(k, safe_string(v)) for (k, v) in print_margins_items])
        SubElement(settings, '{%s}pageMargins' % CHART_NS, margins)
        SubElement(settings, '{%s}pageSetup' % CHART_NS)

    def _write_shapes(self, root):

        if self.chart._shapes:
            SubElement(root, '{%s}userShapes' % CHART_NS, {'{%s}id' % REL_NS:'rId1'})

    def write_rels(self, drawing_id):
        root = Element("{%s}relationships" % REL_NS)

        attrs = {'Id' : 'rId1',
            'Type' : '%s/chartUserShapes' % REL_NS,
            'Target' : '../drawings/drawing%s.xml' % drawing_id }
        SubElement(root, 'Relationship', attrs)
        return get_document_content(root)


class PieChartWriter(BaseChartWriter):

    def _write_options(self, subchart):
        SubElement(subchart, '{%s}varyColors' % CHART_NS, {'val':'1'})

    def _write_axis(self, plot_area, axis, label):
        """Pie Charts have no axes, do nothing"""
        pass


class LineChartWriter(BaseChartWriter):

    def _write_options(self, subchart):
        SubElement(subchart, '{%s}grouping' % CHART_NS, {'val':self.chart.GROUPING})

    def _write_layout(self, root):
        subchart, plotarea = super(LineChartWriter, self)._write_layout(root)
        chart = self.chart

        SubElement(subchart, '{%s}axId' % CHART_NS, {'val':safe_string(chart.x_axis.id)})
        SubElement(subchart, '{%s}axId' % CHART_NS, {'val':safe_string(chart.y_axis.id)})
        super(LineChartWriter, self)._write_axis(plotarea, chart.x_axis, '{%s}catAx' % CHART_NS)
        super(LineChartWriter, self)._write_axis(plotarea, chart.y_axis, '{%s}valAx' % CHART_NS)


class BarChartWriter(LineChartWriter):

    def _write_options(self, subchart):
        SubElement(subchart, '{%s}barDir' % CHART_NS, {'val':'col'})
        SubElement(subchart, '{%s}grouping' % CHART_NS, {'val':self.chart.GROUPING})


class ScatterChartWriter(BaseChartWriter):

    def _write_options(self, subchart):
        SubElement(subchart, '{%s}scatterStyle' % CHART_NS, {'val':'lineMarker'})

    def _write_layout(self, root):
        subchart, plotarea = super(ScatterChartWriter, self)._write_layout(root)
        chart = self.chart
        SubElement(subchart, '{%s}axId' % CHART_NS, {'val':safe_string(chart.x_axis.id)})
        SubElement(subchart, '{%s}axId' % CHART_NS, {'val':safe_string(chart.y_axis.id)})

        self._write_axis(plotarea, chart.x_axis, '{%s}valAx' % CHART_NS)
        self._write_axis(plotarea, chart.y_axis, '{%s}valAx' % CHART_NS)

    def _write_axis(self, plot_area, axis, label):
        self.chart.compute_axes()

        ax = SubElement(plot_area, label)
        SubElement(ax, '{%s}axId' % CHART_NS, {'val':safe_string(axis.id)})

        scaling = SubElement(ax, '{%s}scaling' % CHART_NS)
        SubElement(scaling, '{%s}orientation' % CHART_NS, {'val':axis.orientation})
        SubElement(scaling, '{%s}max' % CHART_NS, {'val':str(float(axis.max))})
        SubElement(scaling, '{%s}min' % CHART_NS, {'val':str(float(axis.min))})

        SubElement(ax, '{%s}axPos' % CHART_NS, {'val':axis.position})
        SubElement(ax, '{%s}majorGridlines' % CHART_NS)
        SubElement(ax, '{%s}numFmt' % CHART_NS, {'formatCode':"General", 'sourceLinked':'1'})
        self._write_axis_title(axis, ax)

        SubElement(ax, '{%s}tickLblPos' % CHART_NS, {'val':axis.tick_label_position})
        SubElement(ax, '{%s}crossAx' % CHART_NS, {'val':str(axis.cross)})
        SubElement(ax, '{%s}crosses' % CHART_NS, {'val':axis.crosses})
        if axis.auto:
            SubElement(ax, '{%s}auto' % CHART_NS, {'val':'1'})
        if axis.label_align:
            SubElement(ax, '{%s}lblAlgn' % CHART_NS, {'val':axis.label_align})
        if axis.label_offset:
            SubElement(ax, '{%s}lblOffset' % CHART_NS, {'val':str(axis.label_offset)})
        SubElement(ax, '{%s}crossBetween' % CHART_NS, {'val':'midCat'})
        SubElement(ax, '{%s}majorUnit' % CHART_NS, {'val':str(float(axis.unit))})


class ChartWriter(object):
    """
    Preserve interface for chart writer
    """

    def __init__(self, chart):
        if isinstance(chart, PieChart):
            self.cw = PieChartWriter(chart)
        elif isinstance(chart, LineChart):
            self.cw = LineChartWriter(chart)
        elif isinstance(chart, BarChart):
            self.cw = BarChartWriter(chart)
        elif isinstance(chart, ScatterChart):
            self.cw = ScatterChartWriter(chart)
        else:
            raise ValueError("Don't know how to handle %s", chart.__class__.__name__)

    def write(self):
        return self.cw.write()
