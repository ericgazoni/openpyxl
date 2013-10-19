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
from openpyxl.shared.compat.itertools import iteritems
from openpyxl.chart import Chart, ErrorBar

try:
    # Python 2
    basestring
except NameError:
    # Python 3
    basestring = str

CHART_TAG = "http://schemas.openxmlformats.org/drawingml/2006/chart"
DRAWING_TAG = "http://schemas.openxmlformats.org/drawingml/2006/main"
REL_TAG = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

def safe_string(value):
    """Safely and consistently format numeric values"""
    if isinstance(value, Number):
        value = "%.15g" % value
    elif not isinstance(value, basestring):
        value = str(value)
    return value

class ChartWriter(object):

    def __init__(self, chart):
        self.chart = chart

    def write(self):
        """ write a chart """
        root = Element("{%s}chartSpace" % CHART_TAG)

        SubElement(root, '{%s}langc' % CHART_TAG, {'val':self.chart.lang})
        self._write_chart(root)
        self._write_print_settings(root)
        self._write_shapes(root)

        return get_document_content(root)

    def _write_chart(self, root):

        chart = self.chart

        ch = SubElement(root, '{%s}chart' % CHART_TAG)
        self._write_title(ch)
        plot_area = SubElement(ch, '{%s}plotArea' % CHART_TAG)
        layout = SubElement(plot_area, '{%s}layout' % CHART_TAG)
        mlayout = SubElement(layout, '{%s}manualLayout' % CHART_TAG)
        SubElement(mlayout, '{%s}layoutTarget' % CHART_TAG, {'val':'inner'})
        SubElement(mlayout, '{%s}xMode' % CHART_TAG, {'val':'edge'})
        SubElement(mlayout, '{%s}yMode' % CHART_TAG, {'val':'edge'})
        SubElement(mlayout, '{%s}x' % CHART_TAG, {'val':safe_string(chart.margin_left)})
        SubElement(mlayout, '{%s}y' % CHART_TAG, {'val':safe_string(chart.margin_top)})
        SubElement(mlayout, '{%s}w' % CHART_TAG, {'val':safe_string(chart.width)})
        SubElement(mlayout, '{%s}h' % CHART_TAG, {'val':safe_string(chart.height)})

        if chart.type == Chart.SCATTER_CHART:
            subchart = SubElement(plot_area, '{%s}scatterChart' % CHART_TAG)
            SubElement(subchart, '{%s}scatterStyle' % CHART_TAG, {'val':'lineMarker'})
        else:
            if chart.type == Chart.BAR_CHART:
                subchart = SubElement(plot_area, '{%s}barChart' % CHART_TAG)
                SubElement(subchart, '{%s}barDir' % CHART_TAG, {'val':'col'})
            else:
                subchart = SubElement(plot_area, '{%s}lineChart' % CHART_TAG)

            SubElement(subchart, '{%s}grouping' % CHART_TAG, {'val':chart.grouping})

        self._write_series(subchart)

        SubElement(subchart, '{%s}axId' % CHART_TAG, {'val':safe_string(chart.x_axis.id)})
        SubElement(subchart, '{%s}axId' % CHART_TAG, {'val':safe_string(chart.y_axis.id)})

        if chart.type == Chart.SCATTER_CHART:
            self._write_axis(plot_area, chart.x_axis, '{%s}valAx' % CHART_TAG)
        else:
            self._write_axis(plot_area, chart.x_axis, '{%s}catAx' % CHART_TAG)
        self._write_axis(plot_area, chart.y_axis, '{%s}valAx' % CHART_TAG)

        self._write_legend(ch)

        SubElement(ch, '{%s}plotVisOnly' % CHART_TAG, {'val':'1'})

    def _write_title(self, chart):
        if self.chart.title != '':
            title = SubElement(chart, '{%s}title' % CHART_TAG)
            tx = SubElement(title, '{%s}tx' % CHART_TAG)
            rich = SubElement(tx, '{%s}rich' % CHART_TAG)
            SubElement(rich, '{%s}bodyPr' % DRAWING_TAG)
            SubElement(rich, '{%s}lstStyle' % DRAWING_TAG)
            p = SubElement(rich, '{%s}p' % DRAWING_TAG)
            pPr = SubElement(p, '{%s}pPr' % DRAWING_TAG)
            SubElement(pPr, '{%s}defRPr' % DRAWING_TAG)
            r = SubElement(p, '{%s}r' % DRAWING_TAG)
            SubElement(r, '{%s}rPr' % DRAWING_TAG, {'lang':self.chart.lang})
            t = SubElement(r, '{%s}t' % DRAWING_TAG).text = self.chart.title
            SubElement(title, '{%s}layout' % CHART_TAG)

    def _write_axis(self, plot_area, axis, label):
        self.chart.compute_axes()

        ax = SubElement(plot_area, label)
        SubElement(ax, '{%s}axId' % CHART_TAG, {'val':safe_string(axis.id)})

        scaling = SubElement(ax, '{%s}scaling' % CHART_TAG)
        SubElement(scaling, '{%s}orientation' % CHART_TAG, {'val':axis.orientation})
        if label == '{%s}valAx' % CHART_TAG:
            SubElement(scaling, '{%s}max' % CHART_TAG, {'val':str(float(axis.max))})
            SubElement(scaling, '{%s}min' % CHART_TAG, {'val':str(float(axis.min))})

        SubElement(ax, '{%s}axPos' % CHART_TAG, {'val':axis.position})
        if label == '{%s}valAx' % CHART_TAG:
            SubElement(ax, '{%s}majorGridlines' % CHART_TAG)
            SubElement(ax, '{%s}numFmt' % CHART_TAG, {'formatCode':"General", 'sourceLinked':'1'})
        if axis.title != '':
            title = SubElement(ax, '{%s}title' % CHART_TAG)
            tx = SubElement(title, '{%s}tx' % CHART_TAG)
            rich = SubElement(tx, '{%s}rich' % CHART_TAG)
            SubElement(rich, '{%s}bodyPr' % DRAWING_TAG)
            SubElement(rich, '{%s}lstStyle' % DRAWING_TAG)
            p = SubElement(rich, '{%s}p' % DRAWING_TAG)
            pPr = SubElement(p, '{%s}pPr' % DRAWING_TAG)
            SubElement(pPr, '{%s}defRPr' % DRAWING_TAG)
            r = SubElement(p, '{%s}r' % DRAWING_TAG)
            SubElement(r, '{%s}rPr' % DRAWING_TAG, {'lang':self.chart.lang})
            t = SubElement(r, '{%s}t' % DRAWING_TAG).text = axis.title
            SubElement(title, '{%s}layout' % CHART_TAG)
        SubElement(ax, '{%s}tickLblPos' % CHART_TAG, {'val':axis.tick_label_position})
        SubElement(ax, '{%s}crossAx' % CHART_TAG, {'val':str(axis.cross)})
        SubElement(ax, '{%s}crosses' % CHART_TAG, {'val':axis.crosses})
        if axis.auto:
            SubElement(ax, '{%s}auto' % CHART_TAG, {'val':'1'})
        if axis.label_align:
            SubElement(ax, '{%s}lblAlgn' % CHART_TAG, {'val':axis.label_align})
        if axis.label_offset:
            SubElement(ax, '{%s}lblOffset' % CHART_TAG, {'val':str(axis.label_offset)})
        if label == '{%s}valAx' % CHART_TAG:
            if self.chart.type == Chart.SCATTER_CHART:
                SubElement(ax, '{%s}crossBetween' % CHART_TAG, {'val':'midCat'})
            else:
                SubElement(ax, '{%s}crossBetween' % CHART_TAG, {'val':'between'})
            SubElement(ax, '{%s}majorUnit' % CHART_TAG, {'val':str(float(axis.unit))})

    def _write_series(self, subchart):

        for i, serie in enumerate(self.chart._series):
            ser = SubElement(subchart, '{%s}ser' % CHART_TAG)
            SubElement(ser, '{%s}idx' % CHART_TAG, {'val':safe_string(i)})
            SubElement(ser, '{%s}order' % CHART_TAG, {'val':safe_string(i)})

            if serie.legend:
                tx = SubElement(ser, '{%s}tx' % CHART_TAG)
                self._write_serial(tx, serie.legend)

            if serie.color:
                sppr = SubElement(ser, '{%s}spPr' % CHART_TAG)
                if self.chart.type == Chart.BAR_CHART:
                    # fill color
                    fillc = SubElement(sppr, '{%s}solidFill' % DRAWING_TAG)
                    SubElement(fillc, '{%s}srgbClr' % DRAWING_TAG, {'val':serie.color})
                # edge color
                ln = SubElement(sppr, '{%s}ln' % DRAWING_TAG)
                fill = SubElement(ln, '{%s}solidFill' % DRAWING_TAG)
                SubElement(fill, '{%s}srgbClr' % DRAWING_TAG, {'val':serie.color})

            if serie.error_bar:
                self._write_error_bar(ser, serie)

            if serie.labels:
                cat = SubElement(ser, '{%s}cat' % CHART_TAG)
                self._write_serial(cat, serie.labels)

            if self.chart.type == Chart.SCATTER_CHART:
                if serie.xvalues:
                    xval = SubElement(ser, '{%s}xVal' % CHART_TAG)
                    self._write_serial(xval, serie.xreference)

                val = SubElement(ser, '{%s}yVal' % CHART_TAG)
            else:
                val = SubElement(ser, '{%s}val' % CHART_TAG)
            self._write_serial(val, serie.reference)

    def _write_serial(self, node, reference, literal=False):

        is_ref = hasattr(reference, 'pos1')
        data_type = reference.data_type
        number_format = getattr(reference, 'number_format')

        mapping = {'n':{'ref':'numRef', 'cache':'numCache'},
                   's':{'ref':'strRef', 'cache':'strCache'}}

        if is_ref:
            ref = SubElement(node, '{%s}%s' %(CHART_TAG, mapping[data_type]['ref']))
            SubElement(ref, '{%s}f' % CHART_TAG).text = str(reference)
            data = SubElement(ref, '{%s}%s' %(CHART_TAG, mapping[data_type]['cache']))
            values = reference.values
        else:
            data = SubElement(node, '{%s}numLit' % CHART_TAG)
            values = (1,)

        if data_type == 'n':
            SubElement(data, '{%s}formatCode' % CHART_TAG).text = number_format or 'General'

        SubElement(data, '{%s}ptCount' % CHART_TAG, {'val':str(len(values))})
        for j, val in enumerate(values):
            point = SubElement(data, '{%s}pt' % CHART_TAG, {'idx':str(j)})
            if not isinstance(val, basestring):
                val = safe_string(val)
            SubElement(point, '{%s}v' % CHART_TAG).text = val

    def _write_error_bar(self, node, serie):

        flag = {ErrorBar.PLUS_MINUS:'both',
                ErrorBar.PLUS:'plus',
                ErrorBar.MINUS:'minus'}

        eb = SubElement(node, '{%s}errBars' % CHART_TAG)
        SubElement(eb, '{%s}errBarType' % CHART_TAG, {'val':flag[serie.error_bar.type]})
        SubElement(eb, '{%s}errValType' % CHART_TAG, {'val':'cust'})

        plus = SubElement(eb, '{%s}plus' % CHART_TAG)
        # apart from setting the type of data series the following has
        # no effect on the writer
        self._write_serial(plus, serie.error_bar.reference,
            literal=(serie.error_bar.type == ErrorBar.MINUS))

        minus = SubElement(eb, '{%s}minus' % CHART_TAG)
        self._write_serial(minus, serie.error_bar.reference,
            literal=(serie.error_bar.type == ErrorBar.PLUS))

    def _write_legend(self, chart):

        if self.chart.show_legend:
            legend = SubElement(chart, '{%s}legend' % CHART_TAG)
            SubElement(legend, '{%s}legendPos' % CHART_TAG, {'val':self.chart.legend.position})
            SubElement(legend, '{%s}layout' % CHART_TAG)

    def _write_print_settings(self, root):

        settings = SubElement(root, '{%s}printSettings' % CHART_TAG)
        SubElement(settings, '{%s}headerFooter' % CHART_TAG)
        try:
            # Python 2
            print_margins_items = iteritems(self.chart.print_margins)
        except AttributeError:
            # Python 3
            print_margins_items = self.chart.print_margins.items()

        margins = dict([(k, safe_string(v)) for (k, v) in print_margins_items])
        SubElement(settings, '{%s}pageMargins' % CHART_TAG, margins)
        SubElement(settings, '{%s}pageSetup' % CHART_TAG)

    def _write_shapes(self, root):

        if self.chart._shapes:
            SubElement(root, '{%s}userShapes' % CHART_TAG, {'{%s}id' % REL_TAG:'rId1'})

    def write_rels(self, drawing_id):
        root = Element("{%s}relationships" % REL_TAG)

        attrs = {'Id' : 'rId1',
            'Type' : '%s/chartUserShapes' % REL_TAG,
            'Target' : '../drawings/drawing%s.xml' % drawing_id }
        SubElement(root, 'Relationship', attrs)
        return get_document_content(root)
