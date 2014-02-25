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


def read_style_table(xml_source):
    """Read styles from the shared style table"""
    style_prop = {'table': {}}
    root = fromstring(xml_source)
    custom_num_formats = parse_custom_num_formats(root)
    style_prop['color_index'] = parse_color_index(root)
    font_list = parse_fonts(root, style_prop['color_index'])
    fill_list = parse_fills(root, style_prop['color_index'])
    border_list = parse_borders(root, style_prop['color_index'])
    style_prop['dxf_list'] = parse_dxfs(root, style_prop['color_index'])
    builtin_formats = NumberFormat._BUILTIN_FORMATS
    cell_xfs = root.find('{%s}cellXfs' % SHEET_MAIN_NS)
    if cell_xfs is not None:  # can happen on bad OOXML writers (e.g. Gnumeric)
        cell_xfs_nodes = safe_iterator(cell_xfs, '{%s}xf' % SHEET_MAIN_NS)
        for index, cell_xfs_node in enumerate(cell_xfs_nodes):
            new_style = Style(static=True)
            number_format_id = int(cell_xfs_node.get('numFmtId'))
            if number_format_id < 164:
                new_style.number_format.format_code = \
                        builtin_formats.get(number_format_id, 'General')
            else:

                if number_format_id in custom_num_formats:
                    new_style.number_format.format_code = \
                            custom_num_formats[number_format_id]
                else:
                    raise MissingNumberFormat('%s' % number_format_id)

            if cell_xfs_node.get('applyAlignment') == '1':
                alignment = cell_xfs_node.find('{%s}alignment' % SHEET_MAIN_NS)
                if alignment is not None:
                    for key in ('horizontal', 'vertical', 'indent'):
                        _value = alignment.get(key)
                        if _value is not None:
                            setattr(new_style.alignment, key, _value)
                    if alignment.get('wrapText'):
                        new_style.alignment.wrap_text = True
                    if alignment.get('shrinkToFit'):
                        new_style.alignment.shrink_to_fit = True
                    if alignment.get('textRotation') is not None:
                        new_style.alignment.text_rotation = int(alignment.get('textRotation'))
                    # ignore justifyLastLine option when horizontal = distributed

            if cell_xfs_node.get('applyFont') == '1':
                new_style.font = deepcopy(font_list[int(cell_xfs_node.get('fontId'))])

            if cell_xfs_node.get('applyFill') == '1':
                new_style.fill = deepcopy(fill_list[int(cell_xfs_node.get('fillId'))])

            if cell_xfs_node.get('applyBorder') == '1':
                new_style.borders = deepcopy(border_list[int(cell_xfs_node.get('borderId'))])

            if cell_xfs_node.get('applyProtection') == '1':
                protection = cell_xfs_node.find('{%s}protection' % SHEET_MAIN_NS)
                # Ignore if there are no protection sub-nodes
                if protection is not None:
                    _protected = protection.get('locked')
                    if _protected is not None:
                        new_style.protection.locked = bool(_protected)
                    _hidden = protection.get('hidden')
                    if _hidden is not None:
                        new_style.protection.hidden = bool(_hidden)

            style_prop['table'][index] = new_style
    return style_prop


def parse_custom_num_formats(root):
    """Read in custom numeric formatting rules from the shared style table"""
    custom_formats = {}
    num_fmts = root.find('{%s}numFmts' % SHEET_MAIN_NS)
    if num_fmts is not None:
        num_fmt_nodes = num_fmts.findall('{%s}numFmt' % SHEET_MAIN_NS)
        for num_fmt_node in num_fmt_nodes:
            custom_formats[int(num_fmt_node.get('numFmtId'))] = \
                    num_fmt_node.get('formatCode').lower()
    return custom_formats


def parse_color_index(root):
    """Read in the list of indexed colors"""
    color_index = []
    colors = root.find('{%s}colors' % SHEET_MAIN_NS)
    if colors is not None:
        indexedColors = colors.find('{%s}indexedColors' % SHEET_MAIN_NS)
        if indexedColors is not None:
            color_nodes = safe_iterator(indexedColors, '{%s}rgbColor' % SHEET_MAIN_NS)
            color_index = [node.get('rgb') for node in color_nodes]
    return color_index or COLOR_INDEX


def parse_dxfs(root, color_index):
    """Read in the dxfs effects - used by conditional formatting."""
    dxf_list = []
    dxfs = root.find('{%s}dxfs' % SHEET_MAIN_NS)
    if dxfs is not None:
        nodes = dxfs.findall('{%s}dxf' % SHEET_MAIN_NS)
        for dxf in nodes:
            dxf_item = {}
            font_list = parse_fonts(dxf, color_index, True)
            if len(font_list):
                dxf_item['font'] = font_list[0]
            fill_list = parse_fills(dxf, color_index, True)
            if len(fill_list):
                dxf_item['fill'] = fill_list[0]
            border_list = parse_borders(dxf, color_index, True)
            if len(border_list):
                dxf_item['border'] = border_list[0]
            dxf_list.append(dxf_item)
    return dxf_list


def parse_fonts(root, color_index, parse_dxf=False):
    """Read in the fonts"""
    font_list = []
    if parse_dxf:
        fonts = root
    else:
        fonts = root.find('{%s}fonts' % SHEET_MAIN_NS)
    if fonts is not None:
        font_nodes = safe_iterator(fonts, '{%s}font' % SHEET_MAIN_NS)
        for font_node in font_nodes:
            font = Font()
            if not parse_dxf:
                fontSizeEl = font_node.find('{%s}sz' % SHEET_MAIN_NS)
                if fontSizeEl is not None:
                    font.size = fontSizeEl.get('val')
                fontNameEl = font_node.find('{%s}name' % SHEET_MAIN_NS)
                if fontNameEl is not None:
                    font.name = fontNameEl.get('val')
            bold = font_node.find('{%s}b' % SHEET_MAIN_NS)
            if bold is not None and 'val' in bold.attrib:
                font.bold = bool(bold.get('val'))
            else:
                font.bold = True if bold is not None else False
            italic = font_node.find('{%s}i' % SHEET_MAIN_NS)
            if italic is not None:
                font.italic = bool(italic.get('val'))
            if italic is not None and 'val' in italic.attrib:
                font.italic = bool(italic.get('val'))
            else:
                font.italic = True if italic is not None else False
            if len(font_node.findall('{%s}u' % SHEET_MAIN_NS)):
                underline = font_node.find('{%s}u' % SHEET_MAIN_NS).get('val')
                font.underline = underline if underline else 'single'
            font.strikethrough = True if len(font_node.findall('{%s}strike' % SHEET_MAIN_NS)) else False
            color = font_node.find('{%s}color' % SHEET_MAIN_NS)
            if color is not None:
                if color.get('indexed') is not None and 0 <= int(color.get('indexed')) < len(color_index):
                    font.color.index = color_index[int(color.get('indexed'))]
                elif color.get('theme') is not None:
                    if color.get('tint') is not None:
                        font.color.index = 'theme:%s:%s' % (color.get('theme'), color.get('tint'))
                    else:
                        font.color.index = 'theme:%s:' % color.get('theme')  # prefix color with theme
                elif color.get('rgb'):
                    font.color.index = color.get('rgb')
            elif parse_dxf:
                font.color = None
            font_list.append(font)
    return font_list


def parse_fills(root, color_index, skip_find=False):
    """Read in the list of fills"""
    fill_list = []
    if skip_find:
        fills = root
    else:
        fills = root.find('{%s}fills' % SHEET_MAIN_NS)
    if fills is not None:
        fillNodes = safe_iterator(fills, '{%s}fill' % SHEET_MAIN_NS)
        for fill in fillNodes:
            # Rotation is unset
            patternFill = fill.find('{%s}patternFill' % SHEET_MAIN_NS)
            if patternFill is not None:
                newFill = Fill()
                newFill.fill_type = patternFill.get('patternType')

                fgColor = patternFill.find('{%s}fgColor' % SHEET_MAIN_NS)
                if fgColor is not None:
                    if fgColor.get('indexed') is not None and 0 <= int(fgColor.get('indexed')) < len(color_index):
                        newFill.start_color.index = color_index[int(fgColor.get('indexed'))]
                    elif fgColor.get('indexed') is not None:
                        # Invalid color - out of range of color_index, set to white
                        newFill.start_color.index = 'FFFFFFFF'
                    elif fgColor.get('theme') is not None:
                        if fgColor.get('tint') is not None:
                            newFill.start_color.index = 'theme:%s:%s' % (fgColor.get('theme'), fgColor.get('tint'))
                        else:
                            newFill.start_color.index = 'theme:%s:' % fgColor.get('theme')  # prefix color with theme
                    else:
                        newFill.start_color.index = fgColor.get('rgb')

                bgColor = patternFill.find('{%s}bgColor' % SHEET_MAIN_NS)
                if bgColor is not None:
                    if bgColor.get('indexed') is not None and 0 <= int(bgColor.get('indexed')) < len(color_index):
                        newFill.end_color.index = color_index[int(bgColor.get('indexed'))]
                    elif bgColor.get('indexed') is not None:
                        # Invalid color - out of range of color_index, set to white
                        newFill.end_color.index = 'FFFFFFFF'
                    elif bgColor.get('theme') is not None:
                        if bgColor.get('tint') is not None:
                            newFill.end_color.index = 'theme:%s:%s' % (bgColor.get('theme'), bgColor.get('tint'))
                        else:
                            newFill.end_color.index = 'theme:%s:' % bgColor.get('theme')  # prefix color with theme
                    elif bgColor.get('rgb'):
                        newFill.end_color.index = bgColor.get('rgb')
                fill_list.append(newFill)
    return fill_list


def parse_borders(root, color_index, skip_find=False):
    """Read in the boarders"""
    border_list = []
    if skip_find:
        borders = root
    else:
        borders = root.find('{%s}borders' % SHEET_MAIN_NS)
    if borders is not None:
        boarderNodes = safe_iterator(borders, '{%s}border' % SHEET_MAIN_NS)
        for border in boarderNodes:
            newBorder = Borders()
            if border.get('diagonalup') == 1:
                newBorder.diagonal_direction = newBorder.DIAGONAL_UP
            if border.get('diagonalDown') == 1:
                if newBorder.diagonal_direction == newBorder.DIAGONAL_UP:
                    newBorder.diagonal_direction = newBorder.DIAGONAL_BOTH
                else:
                    newBorder.diagonal_direction = newBorder.DIAGONAL_DOWN

            for side in ('left', 'right', 'top', 'bottom', 'diagonal'):
                node = border.find('{%s}%s' % (SHEET_MAIN_NS, side))
                if node is not None:
                    borderSide = getattr(newBorder,side)
                    if node.get('style') is not None:
                        borderSide.border_style = node.get('style')
                    color = node.find('{%s}color' % SHEET_MAIN_NS)
                    if color is not None:
                        # Ignore 'auto'
                        if color.get('indexed') is not None and 0 <= int(color.get('indexed')) < len(color_index):
                            borderSide.color.index = color_index[int(color.get('indexed'))]
                        elif color.get('theme') is not None:
                            if color.get('tint') is not None:
                                borderSide.color.index = 'theme:%s:%s' % (color.get('theme'), color.get('tint'))
                            else:
                                borderSide.color.index = 'theme:%s:' % color.get('theme')  # prefix color with theme
                        elif color.get('rgb'):
                            borderSide.color.index = color.get('rgb')
            border_list.append(newBorder)

    return border_list
