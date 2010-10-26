# file openpyxl/writer/styles.py

# Copyright (c) 2010 openpyxl
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
# @author: Eric Gazoni

"""Write the shared style table."""

# package imports
from openpyxl.shared.xmltools import Element, SubElement
from openpyxl.shared.xmltools import get_document_content


def create_style_table(workbook):
    """Compile the style table for this workbook."""
    styles_by_crc = {}
    for worksheet in workbook.worksheets:
        for style in worksheet._styles.values():
            styles_by_crc[hash(style)] = style
    return dict([(style, i + 1) for i, style in
            enumerate(styles_by_crc.values())])


def write_fonts(root_node, style_list):
    """Write the font xml definitions."""
    fonts = SubElement(root_node, 'fonts', {'count': '1'})
    font_node = SubElement(fonts, 'font')
    SubElement(font_node, 'sz', {'val': '11'})
    SubElement(font_node, 'color', {'theme': '1'})
    SubElement(font_node, 'name', {'val': 'Calibri'})
    SubElement(font_node, 'family', {'val': '2'})
    SubElement(font_node, 'scheme', {'val': 'minor'})


def write_fills(root_node, style_list):
    """Write the fill xml definitions."""
    fills = SubElement(root_node, 'fills', {'count': '2'})
    fill = SubElement(fills, 'fill')
    SubElement(fill, 'patternFill', {'patternType': 'none'})
    fill = SubElement(fills, 'fill')
    SubElement(fill, 'patternFill', {'patternType': 'gray125'})


def write_borders(root_node, style_list):
    """Write the border xml definitions."""
    borders = SubElement(root_node, 'borders', {'count': '1'})
    border = SubElement(borders, 'border')
    SubElement(border, 'left')
    SubElement(border, 'right')
    SubElement(border, 'top')
    SubElement(border, 'bottom')
    SubElement(border, 'diagonal')


def write_cell_style_xfs(root_node, style_list):
    """Write the cell style xml."""
    cell_style_xfs = SubElement(root_node, 'cellStyleXfs', {'count': '1'})
    SubElement(cell_style_xfs, 'xf', {'numFmtId': '0', 'fontId': '0',
            'fillId': '0', 'borderId': '0'})


def write_cell_style(root_node, style_list):
    """Write the cell style xml."""
    cell_styles = SubElement(root_node, 'cellStyles', {'count': '1'})
    SubElement(cell_styles, 'cellStyle', {'name': 'Normal',
            'xfId': '0', 'builtinId': '0'})


def write_dxfs(root_node, style_list):
    """Write the differential cell formatting xml."""
    SubElement(root_node, 'dxfs', {'count': '0'})


def write_table_styles(root_node, style_list):
    """Write the xml table of shared styles."""
    SubElement(root_node, 'tableStyles', {'count': '0',
            'defaultTableStyle': 'TableStyleMedium9',
            'defaultPivotStyle': 'PivotStyleLight16'})


def write_number_formats(root_node, style_list):
    """Write the xlm table of shared numeric formats."""
    number_format_table = {}
    number_format_list = []
    exceptions_list = []
    num_fmt_id = 165  # start at a larger value than builtin styles
    num_fmt_offset = 0
    for style in style_list:
        if not style.number_format in number_format_list:
            number_format_list.append(style.number_format)
    for number_format in number_format_list:
        if number_format.is_builtin():
            number_format_table[number_format] = \
                    number_format.builtin_format_id(number_format.format_code)
        else:
            number_format_table[number_format] = num_fmt_id + num_fmt_offset
            num_fmt_offset += 1
            exceptions_list.append(number_format)
    num_fmts = SubElement(root_node, 'numFmts',
            {'count': '%d' % len(exceptions_list)})
    for number_format in exceptions_list:
        SubElement(num_fmts, 'numFmt',
                {'numFmtId': '%d' % number_format_table[number_format],
                'formatCode': '%s' % number_format.format_code})
    return number_format_table


def write_style_table(style_table):
    """Write the style table xml."""
    root_node = Element('styleSheet', {'xmlns':
            'http://schemas.openxmlformats.org/spreadsheetml/2006/main'})
    sorted_styles = sorted(style_table.iteritems(), key = lambda pair: pair[1])
    style_list = [s[0] for s in sorted_styles]
    number_format_table = write_number_formats(root_node, style_list)
    write_fonts(root_node, style_list)
    write_fills(root_node, style_list)
    write_borders(root_node, style_list)
    write_cell_style_xfs(root_node, style_list)

    # writing the cellXfs
    cell_xfs = SubElement(root_node, 'cellXfs',
            {'count': '%d' % (len(style_list) + 1)})
    SubElement(cell_xfs, 'xf', {'numFmtId': '0', 'fontId': '0', 'fillId': '0',
            'xfId': '0', 'borderId': '0'})
    for style in style_list:
        SubElement(cell_xfs, 'xf', {
                'numFmtId': '%d' % number_format_table[style.number_format],
                'applyNumberFormat': '1', 'fontId': '0', 'fillId': '0',
                'xfId': '0', 'borderId': '0'})
    write_cell_style(root_node, style_list)
    write_dxfs(root_node, style_list)
    write_table_styles(root_node, style_list)
    return get_document_content(root_node)
