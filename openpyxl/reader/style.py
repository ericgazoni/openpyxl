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

"""Read shared style definitions"""

# package imports
from openpyxl.xml.functions import fromstring, safe_iterator
from openpyxl.exceptions import MissingNumberFormat
from openpyxl.styles import Style, NumberFormat, Font, Fill, Borders, Protection
from openpyxl.styles.colors import COLOR_INDEX, Color
from openpyxl.xml.constants import SHEET_MAIN_NS
from copy import deepcopy


class SharedStylesParser(object):

    def __init__(self, xml_source):
        self.root = fromstring(xml_source)
        self.style_prop = {'table': {}}
        self.color_index = COLOR_INDEX

    def parse(self):
        self.parse_custom_num_formats()
        self.parse_color_index()
        self.style_prop['color_index'] = self.color_index
        self.font_list = list(self.parse_fonts())
        self.fill_list = list(self.parse_fills())
        self.border_list = list(self.parse_borders())
        self.parse_dxfs()
        self.parse_cell_xfs()


    def parse_custom_num_formats(self):
        """Read in custom numeric formatting rules from the shared style table"""
        custom_formats = {}
        num_fmts = self.root.find('{%s}numFmts' % SHEET_MAIN_NS)
        if num_fmts is not None:
            num_fmt_nodes = safe_iterator(num_fmts, '{%s}numFmt' % SHEET_MAIN_NS)
            for num_fmt_node in num_fmt_nodes:
                custom_formats[int(num_fmt_node.get('numFmtId'))] = \
                        num_fmt_node.get('formatCode').lower()
        self.custom_num_formats = custom_formats


    def parse_color_index(self):
        """Read in the list of indexed colors"""
        colors = self.root.find('{%s}colors' % SHEET_MAIN_NS)
        if colors is not None:
            indexedColors = colors.find('{%s}indexedColors' % SHEET_MAIN_NS)
            if indexedColors is not None:
                color_nodes = safe_iterator(indexedColors, '{%s}rgbColor' % SHEET_MAIN_NS)
                self.color_index = [node.get('rgb') for node in color_nodes]


    def parse_dxfs(self):
        """Read in the dxfs effects - used by conditional formatting."""
        dxf_list = []
        dxfs = self.root.find('{%s}dxfs' % SHEET_MAIN_NS)
        if dxfs is not None:
            nodes = dxfs.findall('{%s}dxf' % SHEET_MAIN_NS)
            for dxf in nodes:
                dxf_item = {}
                font_node = dxf.find('{%s}font' % SHEET_MAIN_NS)
                if font_node is not None:
                    dxf_item['font'] = self.parse_font(font_node)
                fill_node = dxf.find('{%s}fill' % SHEET_MAIN_NS)
                if fill_node is not None:
                    dxf_item['fill'] =  self.parse_fill(fill_node)
                border_node = dxf.find('{%s}border' % SHEET_MAIN_NS)
                if border_node is not None:
                    dxf_item['border'] = self.parse_border(border_node)
                dxf_list.append(dxf_item)
        self.style_prop['dxf_list'] = dxf_list


    def _get_relevant_color(self, color):
        """Utility method for getting the color from different attributes"""
        value = None
        idx = int(color.get('indexed', -1))
        theme = color.get('theme')
        tint = color.get('tint')
        rgb = color.get('rgb')
        if 0 <= idx < len(self.color_index):
            value = self.color_index[idx]
        elif theme is not None:
            if tint is not None:
                value = 'theme:%s:%s' % (theme, tint)
            else:
                value = 'theme:%s:' % theme  # prefix color with theme
        elif rgb is not None:
            value = rgb
        return value


    def parse_fonts(self):
        """Read in the fonts"""
        fonts = self.root.find('{%s}fonts' % SHEET_MAIN_NS)
        if fonts is not None:
            for node in safe_iterator(fonts, '{%s}font' % SHEET_MAIN_NS):
                yield self.parse_font(node)


    def parse_font(self, font_node):
        """Read individual font"""
        font = Font()
        fontSizeEl = font_node.find('{%s}sz' % SHEET_MAIN_NS)
        if fontSizeEl is not None:
            font.size = fontSizeEl.get('val')
        fontNameEl = font_node.find('{%s}name' % SHEET_MAIN_NS)
        if fontNameEl is not None:
            font.name = fontNameEl.get('val')
        bold = font_node.find('{%s}b' % SHEET_MAIN_NS)
        if bold is not None:
            font.bold = bool(bold.get('val', True))
        italic = font_node.find('{%s}i' % SHEET_MAIN_NS)
        if italic is not None:
            font.italic = bool(italic.get('val', True))
        underline = font_node.find('{%s}u' % SHEET_MAIN_NS)
        if underline is not None:
            font.underline = underline.get('val', 'single')
        strikethrough = font_node.find('{%s}strike' % SHEET_MAIN_NS)
        if strikethrough is not None:
            font.strikethrough = True
        color = font_node.find('{%s}color' % SHEET_MAIN_NS)
        if color is not None:
            font.color.index = self._get_relevant_color(color)
        return font


    def parse_fills(self):
        """Read in the list of fills"""
        fills = self.root.find('{%s}fills' % SHEET_MAIN_NS)
        if fills is not None:
            for fill_node in safe_iterator(fills, '{%s}fill' % SHEET_MAIN_NS):
                yield self.parse_fill(fill_node)


    def parse_fill(self, fill_node):
        """Read individual fill"""
        patternFill = fill_node.find('{%s}patternFill' % SHEET_MAIN_NS)
        if patternFill is not None:
            newFill = Fill()
            newFill.fill_type = patternFill.get('patternType')
            fgColor = patternFill.find('{%s}fgColor' % SHEET_MAIN_NS)
            if fgColor is not None:
                newFill.start_color.index = self._get_relevant_color(fgColor)
            bgColor = patternFill.find('{%s}bgColor' % SHEET_MAIN_NS)
            if bgColor is not None:
                newFill.end_color.index = self._get_relevant_color(bgColor)
            return newFill


    def parse_borders(self):
        """Read in the boarders"""
        borders = self.root.find('{%s}borders' % SHEET_MAIN_NS)
        if borders is not None:
            for border_node in safe_iterator(borders, '{%s}border' % SHEET_MAIN_NS):
                yield self.parse_border(border_node)


    def parse_border(self, border_node):
        """Read individual border"""
        newBorder = Borders()
        if bool(border_node.get('diagonalup')):
            newBorder.diagonal_direction = newBorder.DIAGONAL_UP
        if bool(border_node.get('diagonalDown')):
            if newBorder.diagonal_direction == newBorder.DIAGONAL_UP:
                newBorder.diagonal_direction = newBorder.DIAGONAL_BOTH
            else:
                newBorder.diagonal_direction = newBorder.DIAGONAL_DOWN

        for side in ('left', 'right', 'top', 'bottom', 'diagonal'):
            node = border_node.find('{%s}%s' % (SHEET_MAIN_NS, side))
            if node is not None:
                borderSide = getattr(newBorder,side)
                if node.get('style') is not None:
                    borderSide.border_style = node.get('style')
                color = node.find('{%s}color' % SHEET_MAIN_NS)
                if color is not None:
                    borderSide.color.index = self._get_relevant_color(color)
                    # Ignore 'auto'
        return newBorder


    def parse_cell_xfs(self):
        """Read styles from the shared style table"""
        cell_xfs = self.root.find('{%s}cellXfs' % SHEET_MAIN_NS)

        if cell_xfs is None:  # can happen on bad OOXML writers (e.g. Gnumeric)
            return

        builtin_formats = NumberFormat._BUILTIN_FORMATS
        cell_xfs_nodes = safe_iterator(cell_xfs, '{%s}xf' % SHEET_MAIN_NS)
        for index, cell_xfs_node in enumerate(cell_xfs_nodes):
            new_style = Style(static=True)
            number_format_id = int(cell_xfs_node.get('numFmtId'))
            if number_format_id < 164:
                new_style.number_format.format_code = \
                        builtin_formats.get(number_format_id, 'General')
            else:

                if number_format_id in self.custom_num_formats:
                    new_style.number_format.format_code = \
                            self.custom_num_formats[number_format_id]
                else:
                    raise MissingNumberFormat('%s' % number_format_id)

            if bool(cell_xfs_node.get('applyAlignment')):
                alignment = cell_xfs_node.find('{%s}alignment' % SHEET_MAIN_NS)
                if alignment is not None:
                    for key in ('horizontal', 'vertical', 'indent'):
                        _value = alignment.get(key)
                        if _value is not None:
                            setattr(new_style.alignment, key, _value)
                    new_style.alignment.wrap_text = bool(alignment.get('wrapText'))
                    new_style.alignment.shrink_to_fit = bool(alignment.get('shrinkToFit'))
                    text_rotation = alignment.get('textRotation')
                    if text_rotation is not None:
                        new_style.alignment.text_rotation = int(text_rotation)
                    # ignore justifyLastLine option when horizontal = distributed

            if bool(cell_xfs_node.get('applyFont')):
                new_style.font = deepcopy(self.font_list[int(cell_xfs_node.get('fontId'))])

            if bool(cell_xfs_node.get('applyFill')):
                new_style.fill = deepcopy(self.fill_list[int(cell_xfs_node.get('fillId'))])

            if bool(cell_xfs_node.get('applyBorder')):
                new_style.borders = deepcopy(self.border_list[int(cell_xfs_node.get('borderId'))])

            if bool(cell_xfs_node.get('applyProtection')):
                protection = cell_xfs_node.find('{%s}protection' % SHEET_MAIN_NS)
                # Ignore if there are no protection sub-nodes
                if protection is not None:
                    new_style.protection.locked = bool(protection.get('locked'))
                    new_style.protection.hidden = bool(protection.get('hidden'))

            self.style_prop['table'][index] = new_style


def read_style_table(xml_source):
    p = SharedStylesParser(xml_source)
    p.parse()
    return p.style_prop
