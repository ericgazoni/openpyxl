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

from openpyxl.shared.xmltools import Element, SubElement, get_document_content
from openpyxl.shared.ooxml import (
    SHEET_MAIN_NS,
    DRAWING_NS,
    SHEET_DRAWING_NS,
    CHART_NS,
    REL_NS,
    CHART_DRAWING_NS,
    PKG_REL_NS
)


class DrawingWriter(object):
    """ one main drawing file per sheet """

    def __init__(self, sheet):
        self._sheet = sheet

    def write(self):
        """ write drawings for one sheet in one file """

        root = Element("{%s}wsDr" % SHEET_DRAWING_NS)

        for idx, chart in enumerate(self._sheet._charts):
            self._write_chart(root, chart, idx+1)

        for idx, img in enumerate(self._sheet._images):
            self._write_image(root, img, idx+1)

        return get_document_content(root)

    def _write_chart(self, node, chart, idx):
        """Add a chart"""
        drawing = chart.drawing

        #anchor = SubElement(root, 'xdr:twoCellAnchor')
        #(start_row, start_col), (end_row, end_col) = drawing.coordinates
        ## anchor coordinates
        #_from = SubElement(anchor, 'xdr:from')
        #x = SubElement(_from, 'xdr:col').text = str(start_col)
        #x = SubElement(_from, 'xdr:colOff').text = '0'
        #x = SubElement(_from, 'xdr:row').text = str(start_row)
        #x = SubElement(_from, 'xdr:rowOff').text = '0'

        #_to = SubElement(anchor, 'xdr:to')
        #x = SubElement(_to, 'xdr:col').text = str(end_col)
        #x = SubElement(_to, 'xdr:colOff').text = '0'
        #x = SubElement(_to, 'xdr:row').text = str(end_row)
        #x = SubElement(_to, 'xdr:rowOff').text = '0'

        # we only support absolute anchor atm (TODO: oneCellAnchor, twoCellAnchor
        x, y, w, h = drawing.get_emu_dimensions()
        anchor = SubElement(node, '{%s}absoluteAnchor' % SHEET_DRAWING_NS)
        SubElement(anchor, '{%s}pos' % SHEET_DRAWING_NS, {'x':str(x), 'y':str(y)})
        SubElement(anchor, '{%s}ext' % SHEET_DRAWING_NS, {'cx':str(w), 'cy':str(h)})

        # graph frame
        frame = SubElement(anchor, '{%s}graphicFrame' % SHEET_DRAWING_NS, {'macro':''})

        name = SubElement(frame, '{%s}nvGraphicFramePr' % SHEET_DRAWING_NS)
        SubElement(name, '{%s}cNvPr'% SHEET_DRAWING_NS, {'id':'%s' % (idx + 1), 'name':'Chart %s' % idx})
        SubElement(name, '{%s}cNvGraphicFramePr' % SHEET_DRAWING_NS)

        frm = SubElement(frame, '{%s}xfrm'  % SHEET_DRAWING_NS)
        # no transformation
        SubElement(frm, '{%s}off' % DRAWING_NS, {'x':'0', 'y':'0'})
        SubElement(frm, '{%s}ext' % DRAWING_NS, {'cx':'0', 'cy':'0'})

        graph = SubElement(frame, '{%s}graphic' % DRAWING_NS)
        data = SubElement(graph, '{%s}graphicData' % DRAWING_NS, {'uri':CHART_NS})
        SubElement(data, '{%s}chart' % CHART_NS, {'{%s}id' % REL_NS:'rId%s' % idx })

        SubElement(anchor, '{%s}clientData' % SHEET_DRAWING_NS)
        return node

    def _write_image(self, node, img, idx):
        drawing = img.drawing

        x, y, w, h = drawing.get_emu_dimensions()
        anchor = SubElement(node, '{%s}absoluteAnchor' % SHEET_DRAWING_NS)
        SubElement(anchor, '{%s}pos' % SHEET_DRAWING_NS, {'x':str(x), 'y':str(y)})
        SubElement(anchor, '{%s}ext' % SHEET_DRAWING_NS, {'cx':str(w), 'cy':str(h)})

        pic = SubElement(anchor, '{%s}pic' % SHEET_DRAWING_NS)
        name = SubElement(pic, '{%s}nvPicPr' % SHEET_DRAWING_NS)
        SubElement(name, '{%s}cNvPr' % SHEET_DRAWING_NS,
                   {'id':'%s' % (idx + 1),
                    'name':'Picture %s' % idx})
        cNvPicPr = SubElement(name, '{%s}cNvPicPr' % SHEET_DRAWING_NS)
        paras = {"noChangeAspect": "0"}
        if img.nochangeaspect:
            paras["noChangeAspect"] = "1"
        if img.nochangearrowheads:
            paras["noChangeArrowheads"] = "1"

        SubElement(cNvPicPr, '{%s}picLocks' % DRAWING_NS, paras)

        blipfill = SubElement(pic, '{%s}blipFill' % SHEET_DRAWING_NS)
        SubElement(blipfill, '{%s}blip' % DRAWING_NS, {
            '{%s}embed' % REL_NS: 'rId%s' % idx,
            'cstate':'print'
        })
        SubElement(blipfill, '{%s}srcRect' % DRAWING_NS)
        stretch = SubElement(blipfill, '{%s}stretch' % DRAWING_NS)
        SubElement(stretch, '{%s}fillRect' % DRAWING_NS)

        sppr = SubElement(pic, '{%s}spPr' % SHEET_DRAWING_NS, {'bwMode':'auto'})
        frm = SubElement(sppr, '{%s}xfrm' % DRAWING_NS)
        # no transformation
        SubElement(frm, '{%s}off' % DRAWING_NS, {'x':'0', 'y':'0'})
        SubElement(frm, '{%s}ext' % DRAWING_NS, {'cx':'0', 'cy':'0'})
        prstGeom = SubElement(sppr, '{%s}prstGeom' % DRAWING_NS, {'prst':'rect'})
        SubElement(prstGeom, '{%s}avLst' % DRAWING_NS)

        SubElement(sppr, '{%s}noFill' % DRAWING_NS)

        ln = SubElement(sppr, '{%s}ln' % DRAWING_NS, {'w':'1'})
        SubElement(ln, '{%s}noFill' % DRAWING_NS)
        SubElement(ln, '{%s}miter' % DRAWING_NS, {'lim':'800000'})
        SubElement(ln, '{%s}headEnd' % DRAWING_NS)
        SubElement(ln, '{%s}tailEnd' % DRAWING_NS, {'type':'none', 'w':'med', 'len':'med'})
        SubElement(sppr, '{%s}effectLst' % DRAWING_NS)

        SubElement(anchor, '{%s}clientData' % SHEET_DRAWING_NS)

    def write_rels(self, chart_id, image_id):

        root = Element("{%s}Relationships" % PKG_REL_NS)
        i = 0
        for i, chart in enumerate(self._sheet._charts):
            attrs = {'Id' : 'rId%s' % (i + 1),
                'Type' : '%s/chart' % REL_NS,
                'Target' : '../charts/chart%s.xml' % (chart_id + i) }
            SubElement(root, '{%s}Relationship' % PKG_REL_NS, attrs)
        for j, img in enumerate(self._sheet._images):
            attrs = {'Id' : 'rId%s' % (i + j + 1),
                'Type' : '%s/image' % REL_NS,
                'Target' : '../media/image%s.png' % (image_id + j) }
            SubElement(root, '{%s}Relationship' % PKG_REL_NS, attrs)
        return get_document_content(root)


class ShapeWriter(object):
    """ one file per shape """

    def __init__(self, shapes):

        self._shapes = shapes

    def write(self, shape_id):

        root = Element('{%s}userShapes' % CHART_NS)

        for shape in self._shapes:
            anchor = SubElement(root, '{%s}relSizeAnchor' % CHART_DRAWING_NS)

            xstart, ystart, xend, yend = shape.coordinates

            _from = SubElement(anchor, '{%s}from' % CHART_DRAWING_NS)
            SubElement(_from, '{%s}x' % CHART_DRAWING_NS).text = str(xstart)
            SubElement(_from, '{%s}y' % CHART_DRAWING_NS).text = str(ystart)

            _to = SubElement(anchor, '{%s}to' % CHART_DRAWING_NS)
            SubElement(_to, '{%s}x' % CHART_DRAWING_NS).text = str(xend)
            SubElement(_to, '{%s}y' % CHART_DRAWING_NS).text = str(yend)

            sp = SubElement(anchor, '{%s}sp' % CHART_DRAWING_NS, {'macro':'', 'textlink':''})
            nvspr = SubElement(sp, '{%s}nvSpPr' % CHART_DRAWING_NS)
            SubElement(nvspr, '{%s}cNvPr' % CHART_DRAWING_NS, {'id':str(shape_id), 'name':'shape %s' % shape_id})
            SubElement(nvspr, '{%s}cNvSpPr' % CHART_DRAWING_NS)

            sppr = SubElement(sp, '{%s}spPr' % CHART_DRAWING_NS)
            frm = SubElement(sppr, '{%s}xfrm' % DRAWING_NS,)
            # no transformation
            SubElement(frm, '{%s}off' % DRAWING_NS, {'x':'0', 'y':'0'})
            SubElement(frm, '{%s}ext' % DRAWING_NS, {'cx':'0', 'cy':'0'})

            prstgeom = SubElement(sppr, '{%s}prstGeom' % DRAWING_NS, {'prst':str(shape.style)})
            SubElement(prstgeom, '{%s}avLst' % DRAWING_NS)

            fill = SubElement(sppr, '{%s}solidFill' % DRAWING_NS, )
            SubElement(fill, '{%s}srgbClr' % DRAWING_NS, {'val':shape.color})

            border = SubElement(sppr, '{%s}ln' % DRAWING_NS, {'w':str(shape._border_width)})
            sf = SubElement(border, '{%s}solidFill' % DRAWING_NS)
            SubElement(sf, '{%s}srgbClr' % DRAWING_NS, {'val':shape.border_color})

            self._write_style(sp)
            self._write_text(sp, shape)

            shape_id += 1

        return get_document_content(root)

    def _write_text(self, node, shape):
        """ write text in the shape """

        tx_body = SubElement(node, '{%s}txBody' % CHART_DRAWING_NS)
        SubElement(tx_body, '{%s}bodyPr' % DRAWING_NS, {'vertOverflow':'clip'})
        SubElement(tx_body, '{%s}lstStyle' % DRAWING_NS)
        p = SubElement(tx_body, '{%s}p' % DRAWING_NS)
        if shape.text:
            r = SubElement(p, '{%s}r' % DRAWING_NS)
            rpr = SubElement(r, '{%s}rPr' % DRAWING_NS, {'lang':'en-US'})
            fill = SubElement(rpr, '{%s}solidFill' % DRAWING_NS)
            SubElement(fill, '{%s}srgbClr' % DRAWING_NS, {'val':shape.text_color})

            SubElement(r, '{%s}t' % DRAWING_NS).text = shape.text
        else:
            SubElement(p, '{%s}endParaRPr' % DRAWING_NS, {'lang':'en-US'})

    def _write_style(self, node):
        """ write style theme """

        style = SubElement(node, '{%s}style' % CHART_DRAWING_NS)

        ln_ref = SubElement(style, '{%s}lnRef' % DRAWING_NS, {'idx':'2'})
        scheme_clr = SubElement(ln_ref, '{%s}schemeClr' % DRAWING_NS, {'val':'accent1'})
        SubElement(scheme_clr, '{%s}shade' % DRAWING_NS, {'val':'50000'})

        fill_ref = SubElement(style, '{%s}fillRef' % DRAWING_NS, {'idx':'1'})
        SubElement(fill_ref, '{%s}schemeClr' % DRAWING_NS, {'val':'accent1'})

        effect_ref = SubElement(style, '{%s}effectRef' % DRAWING_NS, {'idx':'0'})
        SubElement(effect_ref, '{%s}schemeClr' % DRAWING_NS, {'val':'accent1'})

        font_ref = SubElement(style, '{%s}fontRef' % DRAWING_NS, {'idx':'minor'})
        SubElement(font_ref, '{%s}schemeClr' % DRAWING_NS, {'val':'lt1'})