# file openpyxl/reader/style.py

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

"""Read shared style definitions"""

# package imports
from openpyxl.shared.exc import MissingNumberFormat
from openpyxl.shared.xmltools import fromstring, QName
from openpyxl.style import (Style, NumberFormat, Font, Fill, Borders, Protection,
                            Color, Border, Alignment)


def read_style_table(xml_source):
    """Read styles from the shared style table"""
    table = {}
    xmlns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    root = fromstring(xml_source)
    custom_num_formats = parse_custom_num_formats(root, xmlns)
    color_index = parse_color_index(root, xmlns)
    font_list = parse_fonts(root, xmlns, color_index)
    fill_list = parse_fills(root, xmlns, color_index)
    border_list = parse_borders(root, xmlns, color_index)
    builtin_formats = NumberFormat._BUILTIN_FORMATS
    cell_xfs = root.find(QName(xmlns, 'cellXfs').text)
    if cell_xfs is not None:  # can happen on bad OOXML writers (e.g. Gnumeric)
        cell_xfs_nodes = cell_xfs.findall(QName(xmlns, 'xf').text)
        for index, cell_xfs_node in enumerate(cell_xfs_nodes):
            # new_style = Style()
            style_attrs = {}
            number_format_id = int(cell_xfs_node.get('numFmtId'))
            if number_format_id < 164:
                style_attrs['number_format'] = NumberFormat(builtin_formats.get(number_format_id, 'General'))
            else:
                if number_format_id in custom_num_formats:
                    style_attrs['number_format'] = NumberFormat(custom_num_formats[number_format_id])
                else:
                    raise MissingNumberFormat('%s' % number_format_id)

            if cell_xfs_node.get('applyAlignment') == '1':
                alignment = cell_xfs_node.find(QName(xmlns, 'alignment').text)
                if alignment is not None:
                    alignment_attrs = {}
                    if alignment.get('horizontal') is not None:
                        alignment_attrs['horizontal'] = alignment.get('horizontal')
                    if alignment.get('vertical') is not None:
                        alignment_attrs['vertical'] = alignment.get('vertical')
                    if alignment.get('wrapText'):
                        alignment_attrs['wrap_text'] = True
                    if alignment.get('shrinkToFit'):
                        alignment_attrs['shrink_to_fit'] = True
                    if alignment.get('indent') is not None:
                        alignment_attrs['indent'] = int(alignment.get('indent'))
                    if alignment.get('textRotation') is not None:
                        alignment_attrs['text_rotation'] = int(alignment.get('textRotation'))
                    # ignore justifyLastLine option when horizontal = distributed
                    style_attrs['alignment'] = Alignment(**alignment_attrs)

            if cell_xfs_node.get('applyFont') == '1':
                style_attrs['font'] = font_list[int(cell_xfs_node.get('fontId'))]

            if cell_xfs_node.get('applyFill') == '1':
                style_attrs['fill'] = fill_list[int(cell_xfs_node.get('fillId'))]

            if cell_xfs_node.get('applyBorder') == '1':
                style_attrs['borders'] = border_list[int(cell_xfs_node.get('borderId'))]

            if cell_xfs_node.get('applyProtection') == '1':
                protection = cell_xfs_node.find(QName(xmlns, 'protection').text)
                # Ignore if there are no protection sub-nodes
                if protection is not None:
                    protection_attrs = {}
                    if protection.get('locked') is not None:
                        if protection.get('locked') == '1':
                            protection_attrs['locked'] = Protection.PROTECTION_PROTECTED
                        else:
                            protection_attrs['locked'] = Protection.PROTECTION_UNPROTECTED
                    if protection.get('hidden') is not None:
                        if protection.get('hidden') == '1':
                            protection_attrs['hidden'] = Protection.PROTECTION_PROTECTED
                        else:
                            protection_attrs['hidden'] = Protection.PROTECTION_UNPROTECTED
                    style_attrs['protection'] = Protection(**protection_attrs)

            table[index] = Style(**style_attrs)
    return table

def parse_custom_num_formats(root, xmlns):
    """Read in custom numeric formatting rules from the shared style table"""
    custom_formats = {}
    num_fmts = root.find(QName(xmlns, 'numFmts').text)
    if num_fmts is not None:
        num_fmt_nodes = num_fmts.findall(QName(xmlns, 'numFmt').text)
        for num_fmt_node in num_fmt_nodes:
            custom_formats[int(num_fmt_node.get('numFmtId'))] = \
                    num_fmt_node.get('formatCode').lower()
    return custom_formats

def parse_color_index(root, xmlns):
    """Read in the list of indexed colors"""
    color_index = []
    colors = root.find(QName(xmlns, 'colors').text)
    if colors is not None:
        indexedColors = colors.find(QName(xmlns, 'indexedColors').text)
        if indexedColors is not None:
            color_nodes = indexedColors.findall(QName(xmlns, 'rgbColor').text)
            for color_node in color_nodes:
                color_index.append(color_node.get('rgb'))
    if not color_index:
        # Default Color Index as per http://dmcritchie.mvps.org/excel/colors.htm
        color_index = ['FF000000', 'FFFFFFFF', 'FFFF0000', 'FF00FF00', 'FF0000FF', 'FFFFFF00', 'FFFF00FF', 'FF00FFFF',
                       'FF800000', 'FF008000', 'FF000080', 'FF808000', 'FF800080', 'FF008080', 'FFC0C0C0', 'FF808080',
                       'FF9999FF', 'FF993366', 'FFFFFFCC', 'FFCCFFFF', 'FF660066', 'FFFF8080', 'FF0066CC', 'FFCCCCFF',
                       'FF000080', 'FFFF00FF', 'FFFFFF00', 'FF00FFFF', 'FF800080', 'FF800000', 'FF008080', 'FF0000FF',
                       'FF00CCFF', 'FFCCFFFF', 'FFCCFFCC', 'FFFFFF99', 'FF99CCFF', 'FFFF99CC', 'FFCC99FF', 'FFFFCC99',
                       'FF3366FF', 'FF33CCCC', 'FF99CC00', 'FFFFCC00', 'FFFF9900', 'FFFF6600', 'FF666699', 'FF969696',
                       'FF003366', 'FF339966', 'FF003300', 'FF333300', 'FF993300', 'FF993366', 'FF333399', 'FF333333']
    return color_index

def parse_fonts(root, xmlns, color_index):
    """Read in the fonts"""
    font_list = []
    fonts = root.find(QName(xmlns, 'fonts').text)
    if fonts is not None:
        font_nodes = fonts.findall(QName(xmlns, 'font').text)
        for font_node in font_nodes:
            attrs = {}
            attrs['size'] = font_node.find(QName(xmlns, 'sz').text).get('val')
            attrs['name'] = font_node.find(QName(xmlns, 'name').text).get('val')
            attrs['bold'] = True if len(font_node.findall(QName(xmlns, 'b').text)) else False
            attrs['italic'] = True if len(font_node.findall(QName(xmlns, 'i').text)) else False
            if len(font_node.findall(QName(xmlns, 'u').text)):
                underline = font_node.find(QName(xmlns, 'u').text).get('val')
                attrs['underline'] = underline if underline else 'single'
            color = font_node.find(QName(xmlns, 'color').text)
            if color is not None:
                if color.get('indexed') is not None and 0 <= int(color.get('indexed')) < len(color_index):
                    index = color_index[int(color.get('indexed'))]
                elif color.get('theme') is not None:
                    if color.get('tint') is not None:
                        index = 'theme:%s:%s' % (color.get('theme'), color.get('tint'))
                    else:
                        index = 'theme:%s:' % color.get('theme')  # prefix color with theme
                elif color.get('rgb'):
                    index = color.get('rgb')
                attrs['color'] = Color(index=index)
            font_list.append(Font(**attrs))
    return font_list

def parse_fills(root, xmlns, color_index):
    """Read in the list of fills"""
    fill_list = []
    fills = root.find(QName(xmlns, 'fills').text)
    count = 0
    if fills is not None:
        fillNodes = fills.findall(QName(xmlns, 'fill').text)
        for fill in fillNodes:
            # Rotation is unset
            patternFill = fill.find(QName(xmlns, 'patternFill').text)
            if patternFill is not None:
                attrs = {}
                attrs['fill_type'] = patternFill.get('patternType')

                fgColor = patternFill.find(QName(xmlns, 'fgColor').text)
                if fgColor is not None:
                    if fgColor.get('indexed') is not None and 0 <= int(fgColor.get('indexed')) < len(color_index):
                        index = color_index[int(fgColor.get('indexed'))]
                    elif fgColor.get('indexed') is not None:
                        # Invalid color - out of range of color_index, set to white
                        index = 'FFFFFFFF'
                    elif fgColor.get('theme') is not None:
                        if fgColor.get('tint') is not None:
                            index = 'theme:%s:%s' % (fgColor.get('theme'), fgColor.get('tint'))
                        else:
                            index = 'theme:%s:' % fgColor.get('theme')  # prefix color with theme
                    else:
                        index = fgColor.get('rgb')
                    attrs['start_color'] = Color(index=index)

                bgColor = patternFill.find(QName(xmlns, 'bgColor').text)
                if bgColor is not None:
                    if bgColor.get('indexed') is not None and 0 <= int(bgColor.get('indexed')) < len(color_index):
                        index = color_index[int(bgColor.get('indexed'))]
                    elif bgColor.get('indexed') is not None:
                        # Invalid color - out of range of color_index, set to white
                        index = 'FFFFFFFF'
                    elif bgColor.get('theme') is not None:
                        if bgColor.get('tint') is not None:
                            index = 'theme:%s:%s' % (bgColor.get('theme'), bgColor.get('tint'))
                        else:
                            index = 'theme:%s:' % bgColor.get('theme')  # prefix color with theme
                    elif bgColor.get('rgb'):
                        index = bgColor.get('rgb')
                    attrs['end_color'] = Color(index=index)
                count += 1
                fill_list.append(Fill(**attrs))
    return fill_list

def parse_borders(root, xmlns, color_index):
    """Read in the boarders"""
    border_list = []
    borders = root.find(QName(xmlns, 'borders').text)
    if borders is not None:
        border_nodes = borders.findall(QName(xmlns, 'border').text)
        count = 0
        for border in border_nodes:
            borders_attrs = {}
            if border.get('diagonalup') == 1:
                borders_attrs['diagonal_direction'] = Borders.DIAGONAL_UP
            if border.get('diagonalDown') == 1:
                if borders_attrs['diagonal_direction'] == Borders.DIAGONAL_UP:
                    borders_attrs['diagonal_direction'] = Borders.DIAGONAL_BOTH
                else:
                    borders_attrs['diagonal_direction'] = Borders.DIAGONAL_DOWN

            for side in ('left', 'right', 'top', 'bottom', 'diagonal'):
                node = border.find(QName(xmlns, side).text)
                if node is not None:
                    side_attrs = {}
                    if node.get('style') is not None:
                        side_attrs['border_style'] = node.get('style')
                    color = node.find(QName(xmlns, 'color').text)
                    if color is not None:
                        # Ignore 'auto'
                        if color.get('indexed') is not None and 0 <= int(color.get('indexed')) < len(color_index):
                            index = color_index[int(color.get('indexed'))]
                        elif color.get('theme') is not None:
                            if color.get('tint') is not None:
                                index = 'theme:%s:%s' % (color.get('theme'), color.get('tint'))
                            else:
                                index = 'theme:%s:' % color.get('theme')  # prefix color with theme
                        elif color.get('rgb'):
                            index = color.get('rgb')
                        side_attrs['color'] = Color(index=index)
                    borders_attrs[side] = Border(**side_attrs)

            count += 1
            border_list.append(Borders(**borders_attrs))

    return border_list
