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

        root = Element('{http://schemas.openxmlformats.org/drawingml/2006/chart}chartSpace',
            {'xmlns:c':"http://schemas.openxmlformats.org/drawingml/2006/chart",
             'xmlns:a':"http://schemas.openxmlformats.org/drawingml/2006/main",
             'xmlns:r':"http://schemas.openxmlformats.org/officeDocument/2006/relationships"})

        root = Element("{http://schemas.openxmlformats.org/drawingml/2006/chart}chartSpace")

        SubElement(root, '{http://schemas.openxmlformats.org/drawingml/2006/chart}langc', {'val':self.chart.lang})
        self._write_chart(root)
        self._write_print_settings(root)
        self._write_shapes(root)

        return get_document_content(root)

    def _write_chart(self, root):

        chart = self.chart

        ch = SubElement(root, '{http://schemas.openxmlformats.org/drawingml/2006/chart}chart')
        self._write_title(ch)
        plot_area = SubElement(ch, '{http://schemas.openxmlformats.org/drawingml/2006/chart}plotArea')
        layout = SubElement(plot_area, '{http://schemas.openxmlformats.org/drawingml/2006/chart}layout')
        mlayout = SubElement(layout, '{http://schemas.openxmlformats.org/drawingml/2006/chart}manualLayout')
        SubElement(mlayout, '{http://schemas.openxmlformats.org/drawingml/2006/chart}layoutTarget', {'val':'inner'})
        SubElement(mlayout, '{http://schemas.openxmlformats.org/drawingml/2006/chart}xMode', {'val':'edge'})
        SubElement(mlayout, '{http://schemas.openxmlformats.org/drawingml/2006/chart}yMode', {'val':'edge'})
        SubElement(mlayout, '{http://schemas.openxmlformats.org/drawingml/2006/chart}x', {'val':safe_string(chart.margin_left)})
        SubElement(mlayout, '{http://schemas.openxmlformats.org/drawingml/2006/chart}y', {'val':safe_string(chart.margin_top)})
        SubElement(mlayout, '{http://schemas.openxmlformats.org/drawingml/2006/chart}w', {'val':safe_string(chart.width)})
        SubElement(mlayout, '{http://schemas.openxmlformats.org/drawingml/2006/chart}h', {'val':safe_string(chart.height)})

        if chart.type == Chart.SCATTER_CHART:
            subchart = SubElement(plot_area, '{http://schemas.openxmlformats.org/drawingml/2006/chart}scatterChart')
            SubElement(subchart, '{http://schemas.openxmlformats.org/drawingml/2006/chart}scatterStyle', {'val':'lineMarker'})
        else:
            if chart.type == Chart.BAR_CHART:
                subchart = SubElement(plot_area, '{http://schemas.openxmlformats.org/drawingml/2006/chart}barChart')
                SubElement(subchart, '{http://schemas.openxmlformats.org/drawingml/2006/chart}barDir', {'val':'col'})
            else:
                subchart = SubElement(plot_area, '{http://schemas.openxmlformats.org/drawingml/2006/chart}lineChart')

            SubElement(subchart, '{http://schemas.openxmlformats.org/drawingml/2006/chart}grouping', {'val':chart.grouping})

        self._write_series(subchart)

        SubElement(subchart, '{http://schemas.openxmlformats.org/drawingml/2006/chart}axId', {'val':safe_string(chart.x_axis.id)})
        SubElement(subchart, '{http://schemas.openxmlformats.org/drawingml/2006/chart}axId', {'val':safe_string(chart.y_axis.id)})

        if chart.type == Chart.SCATTER_CHART:
            self._write_axis(plot_area, chart.x_axis, '{http://schemas.openxmlformats.org/drawingml/2006/chart}valAx')
        else:
            self._write_axis(plot_area, chart.x_axis, '{http://schemas.openxmlformats.org/drawingml/2006/chart}catAx')
        self._write_axis(plot_area, chart.y_axis, '{http://schemas.openxmlformats.org/drawingml/2006/chart}valAx')

        self._write_legend(ch)

        SubElement(ch, '{http://schemas.openxmlformats.org/drawingml/2006/chart}plotVisOnly', {'val':'1'})

    def _write_title(self, chart):
        if self.chart.title != '':
            title = SubElement(chart, '{http://schemas.openxmlformats.org/drawingml/2006/chart}title')
            tx = SubElement(title, '{http://schemas.openxmlformats.org/drawingml/2006/chart}tx')
            rich = SubElement(tx, '{http://schemas.openxmlformats.org/drawingml/2006/chart}rich')
            SubElement(rich, '{http://schemas.openxmlformats.org/drawingml/2006/main}bodyPr')
            SubElement(rich, '{http://schemas.openxmlformats.org/drawingml/2006/main}lstStyle')
            p = SubElement(rich, '{http://schemas.openxmlformats.org/drawingml/2006/main}p')
            pPr = SubElement(p, '{http://schemas.openxmlformats.org/drawingml/2006/main}pPr')
            SubElement(pPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}defRPr')
            r = SubElement(p, '{http://schemas.openxmlformats.org/drawingml/2006/main}r')
            SubElement(r, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr', {'lang':self.chart.lang})
            t = SubElement(r, '{http://schemas.openxmlformats.org/drawingml/2006/main}t').text = self.chart.title
            SubElement(title, '{http://schemas.openxmlformats.org/drawingml/2006/chart}layout')

    def _write_axis(self, plot_area, axis, label):
        self.chart.compute_axes()

        ax = SubElement(plot_area, label)
        SubElement(ax, '{http://schemas.openxmlformats.org/drawingml/2006/chart}axId', {'val':safe_string(axis.id)})

        scaling = SubElement(ax, '{http://schemas.openxmlformats.org/drawingml/2006/chart}scaling')
        SubElement(scaling, '{http://schemas.openxmlformats.org/drawingml/2006/chart}orientation', {'val':axis.orientation})
        if label == '{http://schemas.openxmlformats.org/drawingml/2006/chart}valAx':
            SubElement(scaling, '{http://schemas.openxmlformats.org/drawingml/2006/chart}max', {'val':str(float(axis.max))})
            SubElement(scaling, '{http://schemas.openxmlformats.org/drawingml/2006/chart}min', {'val':str(float(axis.min))})

        SubElement(ax, '{http://schemas.openxmlformats.org/drawingml/2006/chart}axPos', {'val':axis.position})
        if label == '{http://schemas.openxmlformats.org/drawingml/2006/chart}valAx':
            SubElement(ax, '{http://schemas.openxmlformats.org/drawingml/2006/chart}majorGridlines')
            SubElement(ax, '{http://schemas.openxmlformats.org/drawingml/2006/chart}numFmt', {'formatCode':"General", 'sourceLinked':'1'})
        if axis.title != '':
            title = SubElement(ax, '{http://schemas.openxmlformats.org/drawingml/2006/chart}title')
            tx = SubElement(title, '{http://schemas.openxmlformats.org/drawingml/2006/chart}tx')
            rich = SubElement(tx, '{http://schemas.openxmlformats.org/drawingml/2006/chart}rich')
            SubElement(rich, '{http://schemas.openxmlformats.org/drawingml/2006/main}bodyPr')
            SubElement(rich, '{http://schemas.openxmlformats.org/drawingml/2006/main}lstStyle')
            p = SubElement(rich, '{http://schemas.openxmlformats.org/drawingml/2006/main}p')
            pPr = SubElement(p, '{http://schemas.openxmlformats.org/drawingml/2006/main}pPr')
            SubElement(pPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}defRPr')
            r = SubElement(p, '{http://schemas.openxmlformats.org/drawingml/2006/main}r')
            SubElement(r, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr', {'lang':self.chart.lang})
            t = SubElement(r, '{http://schemas.openxmlformats.org/drawingml/2006/main}t').text = axis.title
            SubElement(title, '{http://schemas.openxmlformats.org/drawingml/2006/chart}layout')
        SubElement(ax, '{http://schemas.openxmlformats.org/drawingml/2006/chart}tickLblPos', {'val':axis.tick_label_position})
        SubElement(ax, '{http://schemas.openxmlformats.org/drawingml/2006/chart}crossAx', {'val':str(axis.cross)})
        SubElement(ax, '{http://schemas.openxmlformats.org/drawingml/2006/chart}crosses', {'val':axis.crosses})
        if axis.auto:
            SubElement(ax, '{http://schemas.openxmlformats.org/drawingml/2006/chart}auto', {'val':'1'})
        if axis.label_align:
            SubElement(ax, '{http://schemas.openxmlformats.org/drawingml/2006/chart}lblAlgn', {'val':axis.label_align})
        if axis.label_offset:
            SubElement(ax, '{http://schemas.openxmlformats.org/drawingml/2006/chart}lblOffset', {'val':str(axis.label_offset)})
        if label == '{http://schemas.openxmlformats.org/drawingml/2006/chart}valAx':
            if self.chart.type == Chart.SCATTER_CHART:
                SubElement(ax, '{http://schemas.openxmlformats.org/drawingml/2006/chart}crossBetween', {'val':'midCat'})
            else:
                SubElement(ax, '{http://schemas.openxmlformats.org/drawingml/2006/chart}crossBetween', {'val':'between'})
            SubElement(ax, '{http://schemas.openxmlformats.org/drawingml/2006/chart}majorUnit', {'val':str(float(axis.unit))})

    def _write_series(self, subchart):

        for i, serie in enumerate(self.chart._series):
            ser = SubElement(subchart, '{http://schemas.openxmlformats.org/drawingml/2006/chart}ser')
            SubElement(ser, '{http://schemas.openxmlformats.org/drawingml/2006/chart}idx', {'val':safe_string(i)})
            SubElement(ser, '{http://schemas.openxmlformats.org/drawingml/2006/chart}order', {'val':safe_string(i)})

            if serie.legend:
                tx = SubElement(ser, '{http://schemas.openxmlformats.org/drawingml/2006/chart}tx')
                self._write_serial(tx, serie.legend)

            if serie.color:
                sppr = SubElement(ser, '{http://schemas.openxmlformats.org/drawingml/2006/chart}spPr')
                if self.chart.type == Chart.BAR_CHART:
                    # fill color
                    fillc = SubElement(sppr, '{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill')
                    SubElement(fillc, '{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr', {'val':serie.color})
                # edge color
                ln = SubElement(sppr, '{http://schemas.openxmlformats.org/drawingml/2006/main}ln')
                fill = SubElement(ln, '{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill')
                SubElement(fill, '{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr', {'val':serie.color})

            if serie.error_bar:
                self._write_error_bar(ser, serie)

            if serie.labels:
                cat = SubElement(ser, '{http://schemas.openxmlformats.org/drawingml/2006/chart}cat')
                self._write_serial(cat, serie.labels)

            if self.chart.type == Chart.SCATTER_CHART:
                if serie.xvalues:
                    xval = SubElement(ser, '{http://schemas.openxmlformats.org/drawingml/2006/chart}xVal')
                    self._write_serial(xval, serie.xreference)

                val = SubElement(ser, '{http://schemas.openxmlformats.org/drawingml/2006/chart}yVal')
            else:
                val = SubElement(ser, '{http://schemas.openxmlformats.org/drawingml/2006/chart}val')
            self._write_serial(val, serie.reference)

    def _write_serial(self, node, reference, literal=False):

        is_ref = hasattr(reference, 'pos1')
        data_type = reference.data_type
        number_format = getattr(reference, 'number_format')

        mapping = {'n':{'ref':'numRef', 'cache':'numCache'},
                   's':{'ref':'strRef', 'cache':'strCache'}}

        if is_ref:
            ref = SubElement(node, '{http://schemas.openxmlformats.org/drawingml/2006/chart}%s' % mapping[data_type]['ref'])
            SubElement(ref, '{http://schemas.openxmlformats.org/drawingml/2006/chart}f').text = str(reference)
            data = SubElement(ref, '{http://schemas.openxmlformats.org/drawingml/2006/chart}%s' % mapping[data_type]['cache'])
            values = reference.values
        else:
            data = SubElement(node, '{http://schemas.openxmlformats.org/drawingml/2006/chart}numLit')
            values = (1,)

        if data_type == 'n':
            SubElement(data, '{http://schemas.openxmlformats.org/drawingml/2006/chart}formatCode').text = number_format or 'General'

        SubElement(data, '{http://schemas.openxmlformats.org/drawingml/2006/chart}ptCount', {'val':str(len(values))})
        for j, val in enumerate(values):
            point = SubElement(data, '{http://schemas.openxmlformats.org/drawingml/2006/chart}pt', {'idx':str(j)})
            if not isinstance(val, basestring):
                val = safe_string(val)
            SubElement(point, '{http://schemas.openxmlformats.org/drawingml/2006/chart}v').text = val

    def _write_error_bar(self, node, serie):

        flag = {ErrorBar.PLUS_MINUS:'both',
                ErrorBar.PLUS:'plus',
                ErrorBar.MINUS:'minus'}

        eb = SubElement(node, '{http://schemas.openxmlformats.org/drawingml/2006/chart}errBars')
        SubElement(eb, '{http://schemas.openxmlformats.org/drawingml/2006/chart}errBarType', {'val':flag[serie.error_bar.type]})
        SubElement(eb, '{http://schemas.openxmlformats.org/drawingml/2006/chart}errValType', {'val':'cust'})

        plus = SubElement(eb, '{http://schemas.openxmlformats.org/drawingml/2006/chart}plus')
        # apart from setting the type of data series the following has
        # no effect on the writer
        self._write_serial(plus, serie.error_bar.reference,
            literal=(serie.error_bar.type == ErrorBar.MINUS))

        minus = SubElement(eb, '{http://schemas.openxmlformats.org/drawingml/2006/chart}minus')
        self._write_serial(minus, serie.error_bar.reference,
            literal=(serie.error_bar.type == ErrorBar.PLUS))

    def _write_legend(self, chart):

        if self.chart.show_legend:
            legend = SubElement(chart, '{http://schemas.openxmlformats.org/drawingml/2006/chart}legend')
            SubElement(legend, '{http://schemas.openxmlformats.org/drawingml/2006/chart}legendPos', {'val':self.chart.legend.position})
            SubElement(legend, '{http://schemas.openxmlformats.org/drawingml/2006/chart}layout')

    def _write_print_settings(self, root):

        settings = SubElement(root, '{http://schemas.openxmlformats.org/drawingml/2006/chart}printSettings')
        SubElement(settings, '{http://schemas.openxmlformats.org/drawingml/2006/chart}headerFooter')
        try:
            # Python 2
            print_margins_items = iteritems(self.chart.print_margins)
        except AttributeError:
            # Python 3
            print_margins_items = self.chart.print_margins.items()

        margins = dict([(k, safe_string(v)) for (k, v) in print_margins_items])
        SubElement(settings, '{http://schemas.openxmlformats.org/drawingml/2006/chart}pageMargins', margins)
        SubElement(settings, '{http://schemas.openxmlformats.org/drawingml/2006/chart}pageSetup')

    def _write_shapes(self, root):

        if self.chart._shapes:
            SubElement(root, '{http://schemas.openxmlformats.org/drawingml/2006/chart}userShapes', {'{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id':'rId1'})

    def write_rels(self, drawing_id):

        root = Element('Relationships', {'xmlns' : 'http://schemas.openxmlformats.org/package/2006/relationships'})
        attrs = {'Id' : 'rId1',
            'Type' : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes',
            'Target' : '../drawings/drawing%s.xml' % drawing_id }
        SubElement(root, 'Relationship', attrs)
        return get_document_content(root)
