from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl
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

"""Write the shared style table."""

# package imports
from openpyxl.xml.functions import (
    Element,
    SubElement,
    get_document_content
    )
from openpyxl.xml.constants import SHEET_MAIN_NS

from openpyxl.styles import DEFAULTS, Protection


class StyleWriter(object):

    def __init__(self, workbook):
        self._style_list = self._get_style_list(workbook)
        self._style_properties = workbook.style_properties
        self._root = Element('styleSheet', {'xmlns':SHEET_MAIN_NS})

    def _get_style_list(self, workbook):
        crc = {}
        for worksheet in workbook.worksheets:
            uniqueStyles = dict((id(style), style) for style in worksheet._styles.values()).values()
            for style in uniqueStyles:
                crc[hash(style)] = style
        self.style_table = dict([(style, i + 1) \
            for i, style in enumerate(crc.values())])
        sorted_styles = sorted(self.style_table.items(), \
            key=lambda pair:pair[1])
        return [s[0] for s in sorted_styles]

    def get_style_by_hash(self):
        return dict([(hash(style), id) \
            for style, id in self.style_table.items()])

    def write_table(self):
        number_format_table = self._write_number_formats()
        fonts_table = self._write_fonts()
        fills_table = self._write_fills()
        borders_table = self._write_borders()
        self._write_cell_style_xfs()
        self._write_cell_xfs(number_format_table, fonts_table, fills_table, borders_table)
        self._write_cell_style()
        self._write_dxfs()
        self._write_table_styles()

        return get_document_content(xml_node=self._root)


    def _unpack_color(self, node, color, key='color'):
        """Convert colors encoded as RGB, theme or tints
        Possible values
        RGB: #4F81BD
        Theme: theme:9
        Tint: theme:9:7 # guess work
        """
        if color is None:
            return
        if not ":" in color:
            SubElement(node, key, {'rgb':color})
        else:
            _, theme, tint = color.split(":")
            if tint == '':
                SubElement(node, key, {'theme':theme})
            else:
                SubElement(node, key, {'theme':theme, 'tint':tint})


    def _write_fonts(self):
        """ add fonts part to root
            return {font.crc => index}
        """

        fonts = SubElement(self._root, 'fonts')

        # default
        font_node = SubElement(fonts, 'font')
        SubElement(font_node, 'sz', {'val':'11'})
        SubElement(font_node, 'color', {'theme':'1'})
        SubElement(font_node, 'name', {'val':'Calibri'})
        SubElement(font_node, 'family', {'val':'2'})
        SubElement(font_node, 'scheme', {'val':'minor'})

        # others
        table = {}
        index = 1
        for st in self._style_list:
            if st.font != DEFAULTS.font and st.font not in table:
                table[st.font] = index
                font_node = SubElement(fonts, 'font')
                SubElement(font_node, 'sz', {'val':str(st.font.size)})
                self._unpack_color(font_node, st.font.color.index)
                SubElement(font_node, 'name', {'val':st.font.name})
                SubElement(font_node, 'family', {'val':'2'})
                # Don't write the 'scheme' element because it appears to prevent
                # the font name from being applied in Excel.
                #SubElement(font_node, 'scheme', {'val':'minor'})
                if st.font.bold:
                    SubElement(font_node, 'b')
                if st.font.italic:
                    SubElement(font_node, 'i')
                if st.font.underline == 'single':
                    SubElement(font_node, 'u')

                index += 1

        fonts.attrib["count"] = str(index)
        return table

    def _write_fills(self):
        fills = SubElement(self._root, 'fills', {'count':'2'})
        fill = SubElement(fills, 'fill')
        SubElement(fill, 'patternFill', {'patternType':'none'})
        fill = SubElement(fills, 'fill')
        SubElement(fill, 'patternFill', {'patternType':'gray125'})

        table = {}
        index = 2
        for st in self._style_list:
            if st.fill != DEFAULTS.fill and st.fill not in table:
                table[st.fill] = index
                fill = SubElement(fills, 'fill')
                if st.fill.fill_type != DEFAULTS.fill.fill_type:
                    node = SubElement(fill, 'patternFill', {'patternType':st.fill.fill_type})
                    if st.fill.start_color != DEFAULTS.fill.start_color:
                        self._unpack_color(node, st.fill.start_color.index, 'fgColor')

                    if st.fill.end_color != DEFAULTS.fill.end_color:
                        self._unpack_color(node, st.fill.end_color.index, 'bgColor')

                index += 1

        fills.attrib["count"] = str(index)
        return table

    def _write_borders(self):
        borders = SubElement(self._root, 'borders')

        # default
        border = SubElement(borders, 'border')
        SubElement(border, 'left')
        SubElement(border, 'right')
        SubElement(border, 'top')
        SubElement(border, 'bottom')
        SubElement(border, 'diagonal')

        # others
        table = {}
        index = 1
        for st in self._style_list:
            if st.borders != DEFAULTS.borders and st.borders not in table:
                table[st.borders] = index
                border = SubElement(borders, 'border')
                # caution: respect this order
                for side in ('left', 'right', 'top', 'bottom', 'diagonal'):
                    obj = getattr(st.borders, side)
                    if obj.border_style is None or obj.border_style == 'none':
                        node = SubElement(border, side)
                    else:
                        node = SubElement(border, side, {'style':obj.border_style})
                        self._unpack_color(node, obj.color.index)

                index += 1

        borders.attrib["count"] = str(index)
        return table

    def _write_cell_style_xfs(self):
        cell_style_xfs = SubElement(self._root, 'cellStyleXfs', {'count':'1'})
        xf = SubElement(cell_style_xfs, 'xf',
            {'numFmtId':"0", 'fontId':"0", 'fillId':"0", 'borderId':"0"})

    def _write_cell_xfs(self, number_format_table, fonts_table, fills_table, borders_table):
        """ write styles combinations based on ids found in tables """

        # writing the cellXfs
        cell_xfs = SubElement(self._root, 'cellXfs',
            {'count':'%d' % (len(self._style_list) + 1)})

        # default
        def _get_default_vals():
            return dict(numFmtId='0', fontId='0', fillId='0',
                xfId='0', borderId='0')

        SubElement(cell_xfs, 'xf', _get_default_vals())

        for st in self._style_list:
            vals = _get_default_vals()

            if st.font != DEFAULTS.font:
                vals['fontId'] = str(fonts_table[st.font])
                vals['applyFont'] = '1'

            if st.borders != DEFAULTS.borders:
                vals['borderId'] = str(borders_table[st.borders])
                vals['applyBorder'] = '1'

            if st.fill != DEFAULTS.fill:
                vals['fillId'] = str(fills_table[st.fill])
                vals['applyFill'] = '1'

            if st.number_format != DEFAULTS.number_format:
                vals['numFmtId'] = '%d' % number_format_table[st.number_format]
                vals['applyNumberFormat'] = '1'

            if st.alignment != DEFAULTS.alignment:
                vals['applyAlignment'] = '1'

            if st.protection != DEFAULTS.protection:
                vals['applyProtection'] = '1'

            node = SubElement(cell_xfs, 'xf', vals)

            if st.alignment != DEFAULTS.alignment:
                alignments = {}

                for align_attr in ['horizontal', 'vertical']:
                    if getattr(st.alignment, align_attr) != getattr(DEFAULTS.alignment, align_attr):
                        alignments[align_attr] = getattr(st.alignment, align_attr)

                    if st.alignment.wrap_text != DEFAULTS.alignment.wrap_text:
                        alignments['wrapText'] = '1'

                    if st.alignment.shrink_to_fit != DEFAULTS.alignment.shrink_to_fit:
                        alignments['shrinkToFit'] = '1'

                    if st.alignment.indent > 0:
                        alignments['indent'] = '%s' % st.alignment.indent

                    if st.alignment.text_rotation > 0:
                        alignments['textRotation'] = '%s' % st.alignment.text_rotation
                    elif st.alignment.text_rotation < 0:
                        alignments['textRotation'] = '%s' % (90 - st.alignment.text_rotation)

                SubElement(node, 'alignment', alignments)

            if st.protection != DEFAULTS.protection:
                protections = {}

                if st.protection.locked == Protection.PROTECTION_PROTECTED:
                    protections['locked'] = '1'
                elif st.protection.locked == Protection.PROTECTION_UNPROTECTED:
                    protections['locked'] = '0'

                if st.protection.hidden == Protection.PROTECTION_PROTECTED:
                    protections['hidden'] = '1'
                elif st.protection.hidden == Protection.PROTECTION_UNPROTECTED:
                    protections['hidden'] = '0'

                SubElement(node, 'protection', protections)

    def _write_cell_style(self):
        cell_styles = SubElement(self._root, 'cellStyles', {'count':'1'})
        cell_style = SubElement(cell_styles, 'cellStyle',
            {'name':"Normal", 'xfId':"0", 'builtinId':"0"})

    def _write_dxfs(self):
        if self._style_properties and 'dxf_list' in self._style_properties:
            dxfs = SubElement(self._root, 'dxfs', {'count': str(len(self._style_properties['dxf_list']))})
            for d in self._style_properties['dxf_list']:
                dxf = SubElement(dxfs, 'dxf')
                if 'font' in d and d['font'] is not None:
                    font_node = SubElement(dxf, 'font')
                    if d['font'].color is not None:
                        self._unpack_color(font_node, d['font'].color.index)
                    if d['font'].bold:
                        SubElement(font_node, 'b', {'val': '1'})
                    if d['font'].italic:
                        SubElement(font_node, 'i', {'val': '1'})
                    if d['font'].underline != 'none':
                        SubElement(font_node, 'u', {'val': d['font'].underline})
                    if d['font'].strikethrough:
                        SubElement(font_node, 'strike')

                if 'fill' in d:
                    f = d['fill']
                    fill = SubElement(dxf, 'fill')
                    if f.fill_type:
                        node = SubElement(fill, 'patternFill', {'patternType': f.fill_type})
                    else:
                        node = SubElement(fill, 'patternFill')
                    if f.start_color != DEFAULTS.fill.start_color:
                        self._unpack_color(node, f.start_color.index, 'fgColor')

                    if f.end_color != DEFAULTS.fill.end_color:
                        self._unpack_color(node, f.end_color.index, 'bgColor')

                if 'border' in d:
                    borders = d['border']
                    border = SubElement(dxf, 'border')
                    # caution: respect this order
                    for side in ('left', 'right', 'top', 'bottom'):
                        obj = getattr(borders, side)
                        if obj.border_style is None or obj.border_style == 'none':
                            node = SubElement(border, side)
                        else:
                            node = SubElement(border, side, {'style': obj.border_style})
                            self._unpack_color(node, obj.color.index)
        else:
            dxfs = SubElement(self._root, 'dxfs', {'count': '0'})
        return dxfs

    def _write_table_styles(self):

        table_styles = SubElement(self._root, 'tableStyles',
            {'count':'0', 'defaultTableStyle':'TableStyleMedium9',
            'defaultPivotStyle':'PivotStyleLight16'})

    def _write_number_formats(self):

        number_format_table = {}

        number_format_list = []
        exceptions_list = []
        num_fmt_id = 165 # start at a greatly higher value as any builtin can go
        num_fmt_offset = 0

        for style in self._style_list:

            if not style.number_format in number_format_list  :
                number_format_list.append(style.number_format)

        for number_format in number_format_list:

            if number_format.is_builtin():
                btin = number_format.builtin_format_id(number_format.format_code)
                number_format_table[number_format] = btin
            else:
                number_format_table[number_format] = num_fmt_id + num_fmt_offset
                num_fmt_offset += 1
                exceptions_list.append(number_format)

        num_fmts = SubElement(self._root, 'numFmts',
            {'count':'%d' % len(exceptions_list)})

        for number_format in exceptions_list :
            SubElement(num_fmts, 'numFmt',
                {'numFmtId':'%d' % number_format_table[number_format],
                'formatCode':'%s' % number_format.format_code})

        return number_format_table
