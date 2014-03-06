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

"""Reader for a single worksheet."""

# Python stdlib imports
from warnings import warn

# compatibility imports
from openpyxl.compat import BytesIO
from openpyxl.xml.functions import iterparse

# package imports
from openpyxl import LXML
from openpyxl.cell import get_column_letter
from openpyxl.cell import Cell, coordinate_from_string
from openpyxl.worksheet import Worksheet, ColumnDimension, RowDimension
from openpyxl.worksheet.iter_worksheet import IterableWorksheet
from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.xml.functions import safe_iterator
from openpyxl.styles import Color
from openpyxl.formatting import ConditionalFormatting
from openpyxl.formatting.rules import FormatRule, CellIsRule, ColorScaleRule, FormatRule


def _get_xml_iter(xml_source):

    if not hasattr(xml_source, 'read'):
        if hasattr(xml_source, 'decode'):
            return BytesIO(xml_source)
        else:
            return BytesIO(xml_source.encode('utf-8'))
    else:
        try:
            xml_source.seek(0)
        except:
            pass
        return xml_source


class WorkSheetParser(object):

    COL_TAG = '{%s}col' % SHEET_MAIN_NS
    ROW_TAG = '{%s}row' % SHEET_MAIN_NS
    CELL_TAG = '{%s}c' % SHEET_MAIN_NS
    VALUE_TAG = '{%s}v' % SHEET_MAIN_NS
    FORMULA_TAG = '{%s}f' % SHEET_MAIN_NS
    MERGE_TAG = '{%s}mergeCell' % SHEET_MAIN_NS

    def __init__(self, ws, xml_source, string_table, style_table, color_index=None):
        self.ws = ws
        self.source = xml_source
        self.string_table = string_table
        self.style_table = style_table
        self.color_index = color_index
        self.guess_types = ws.parent._guess_types
        self.data_only = ws.parent.data_only

    def parse(self):
        stream = _get_xml_iter(self.source)
        it = iterparse(stream)

        dispatcher = {
            '{%s}mergeCells' % SHEET_MAIN_NS: self.parse_merge,
            '{%s}col' % SHEET_MAIN_NS: self.parse_column_dimensions,
            '{%s}row' % SHEET_MAIN_NS: self.parse_row_dimensions,
            '{%s}printOptions' % SHEET_MAIN_NS: self.parse_print_options,
            '{%s}pageMargins' % SHEET_MAIN_NS: self.parse_margins,
            '{%s}pageSetup' % SHEET_MAIN_NS: self.parse_page_setup,
            '{%s}headerFooter' % SHEET_MAIN_NS: self.parse_header_footer,
            '{%s}conditionalFormatting' % SHEET_MAIN_NS: self.parser_conditional_formatting,
            '{%s}autoFilter' % SHEET_MAIN_NS: self.parse_auto_filter
                      }
        tags = dispatcher.keys()
        stream = _get_xml_iter(self.source)
        it = iterparse(stream, tag=tags)

        for event, element in it:
            tag_name = element.tag
            if tag_name in dispatcher:
                dispatcher[tag_name](element)
                element.clear()

        # Handle parsed conditional formatting rules together.
        if len(self.ws.conditional_formatting.parse_rules):
            self.ws.conditional_formatting.update(self.ws.conditional_formatting.parse_rules)

    def parse_cell(self, element):
        value = element.findtext(self.VALUE_TAG)
        formula = element.find(self.FORMULA_TAG)

        coordinate = element.get('r')
        style_id = element.get('s')
        if style_id is not None:
            self.ws._styles[coordinate] = self.style_table.get(int(style_id))

        if value is not None and value is not '':
            data_type = element.get('t', 'n')
            if data_type == Cell.TYPE_STRING:
                value = self.string_table.get(int(value))
            elif data_type == Cell.TYPE_BOOL:
                value = bool(int(value))
            elif data_type == 'n':
                value = float(value)
            if formula is not None and not self.data_only:
                if formula.text:
                    value = "=" + formula.text
                else:
                    value = "="
                formula_type = formula.get('t')
                if formula_type:
                    self.ws.formula_attributes[coordinate] = {'t': formula_type}
                    if formula.get('si'):  # Shared group index for shared formulas
                        self.ws.formula_attributes[coordinate]['si'] = formula.get('si')
                    if formula.get('ref'):  # Range for shared formulas
                        self.ws.formula_attributes[coordinate]['ref'] = formula.get('ref')
            if not self.guess_types and formula is None:
                self.ws.cell(coordinate).set_explicit_value(value=value, data_type=data_type)
            else:
                self.ws.cell(coordinate).value = value


    def parse_merge(self, element):
        for mergeCell in safe_iterator(element, ('{%s}mergeCell' % SHEET_MAIN_NS)):
            self.ws.merge_cells(mergeCell.get('ref'))


    def parse_column_dimensions(self, col):
        min = int(col.get('min')) if col.get('min') else 1
        max = int(col.get('max')) if col.get('max') else 1
        # Ignore ranges that go up to the max column 16384.  Columns need to be extended to handle
        # ranges without creating an entry for every single one.
        if max != 16384:
            for colId in range(min, max + 1):
                column = get_column_letter(colId)
                width = col.get("width")
                auto_size = col.get('bestFit') == '1'
                visible = col.get('hidden') != '1'
                outline = col.get('outlineLevel') or 0
                collapsed = col.get('collapsed') == '1'
                style_index = col.get('style')
                if style_index is not None:
                    self.ws._styles[column] = self.style_table.get(int(style_index))
                if column not in self.ws.column_dimensions:
                    new_dim = ColumnDimension( index=column, width=width,
                                               auto_size=auto_size, visible=visible,
                                               outline_level=outline, collapsed=collapsed)
                    self.ws.column_dimensions[column] = new_dim


    def parse_row_dimensions(self, row):
        rowId = int(row.get('r'))
        ht = row.get('ht', -1)
        if rowId not in self.ws.row_dimensions:
            self.ws.row_dimensions[rowId] = RowDimension(rowId, height=ht)
        style_index = row.get('s')
        if row.get('customFormat') and style_index:
            self.ws._styles[rowId] = self.style_table.get(int(style_index))
        for cell in safe_iterator(row, self.CELL_TAG):
            self.parse_cell(cell)


    def parse_print_options(self, element):
        hc = element.get('horizontalCentered')
        if hc is not None:
            self.ws.page_setup.horizontalCentered = hc
        vc = element.get('verticalCentered')
        if vc is not None:
            self.ws.page_setup.verticalCentered = vc


    def parse_margins(self, element):
        for key in ("left", "right", "top", "bottom", "header", "footer"):
            value = element.get(key)
            if value is not None:
                setattr(self.ws.page_margins, key, float(value))


    def parse_page_setup(self, element):
        for key in ("orientation", "paperSize", "scale", "fitToPage",
                    "fitToHeight", "fitToWidth", "firstPageNumber",
                    "useFirstPageNumber"):
            value = element.get(key)
            if value is not None:
                setattr(self.ws.page_setup, key, value)


    def parse_header_footer(self, element):
        oddHeader = element.find('{%s}oddHeader' % SHEET_MAIN_NS)
        if oddHeader is not None and oddHeader.text is not None:
            self.ws.header_footer.setHeader(oddHeader.text)
        oddFooter = element.find('{%s}oddFooter' % SHEET_MAIN_NS)
        if oddFooter is not None and oddFooter.text is not None:
            self.ws.header_footer.setFooter(oddFooter.text)


    def parser_conditional_formatting(self, element):
        for cf in safe_iterator(element, '{%s}conditionalFormatting' % SHEET_MAIN_NS):
            if not cf.get('sqref'):
                # Potentially flag - this attribute should always be present.
                continue
            range_string = cf.get('sqref')
            cfRules = cf.findall('{%s}cfRule' % SHEET_MAIN_NS)
            if range_string not in self.ws.conditional_formatting.parse_rules:
                self.ws.conditional_formatting.parse_rules[range_string] = []
            for cfRule in cfRules:
                if not cfRule.get('type') or cfRule.get('type') == 'dataBar':
                    # dataBar conditional formatting isn't supported, as it relies on the complex <extLst> tag
                    continue
                rule = {'type': cfRule.get('type')}
                for attr in ConditionalFormatting.rule_attributes:
                    if cfRule.get(attr) is not None:
                        if attr == 'priority':
                            rule[attr] = int(cfRule.get(attr))
                        else:
                            rule[attr] = cfRule.get(attr)

                formula = cfRule.findall('{%s}formula' % SHEET_MAIN_NS)
                for f in formula:
                    if 'formula' not in rule:
                        rule['formula'] = []
                    rule['formula'].append(f.text)

                colorScale = cfRule.find('{%s}colorScale' % SHEET_MAIN_NS)
                if colorScale is not None:
                    rule['colorScale'] = {'cfvo': [], 'color': []}
                    cfvoNodes = colorScale.findall('{%s}cfvo' % SHEET_MAIN_NS)
                    for node in cfvoNodes:
                        cfvo = {}
                        if node.get('type') is not None:
                            cfvo['type'] = node.get('type')
                        if node.get('val') is not None:
                            cfvo['val'] = node.get('val')
                        rule['colorScale']['cfvo'].append(cfvo)
                    colorNodes = colorScale.findall('{%s}color' % SHEET_MAIN_NS)
                    for color in colorNodes:
                        c = Color(Color.BLACK)
                        if self.color_index\
                           and color.get('indexed') is not None\
                           and 0 <= int(color.get('indexed')) < len(self.color_index):
                            c.index = self.color_index[int(color.get('indexed'))]
                        if color.get('theme') is not None:
                            if color.get('tint') is not None:
                                c.index = 'theme:%s:%s' % (color.get('theme'), color.get('tint'))
                            else:
                                c.index = 'theme:%s:' % color.get('theme')  # prefix color with theme
                        elif color.get('rgb'):
                            c.index = color.get('rgb')
                        rule['colorScale']['color'].append(c)

                iconSet = cfRule.find('{%s}iconSet' % SHEET_MAIN_NS)
                if iconSet is not None:
                    rule['iconSet'] = {'cfvo': []}
                    for iconAttr in ConditionalFormatting.icon_attributes:
                        if iconSet.get(iconAttr) is not None:
                            rule['iconSet'][iconAttr] = iconSet.get(iconAttr)
                    cfvoNodes = iconSet.findall('{%s}cfvo' % SHEET_MAIN_NS)
                    for node in cfvoNodes:
                        cfvo = {}
                        if node.get('type') is not None:
                            cfvo['type'] = node.get('type')
                        if node.get('val') is not None:
                            cfvo['val'] = node.get('val')
                        rule['iconSet']['cfvo'].append(cfvo)

                self.ws.conditional_formatting.parse_rules[range_string].append(rule)

    def parse_auto_filter(self, element):
        self.ws.auto_filter.ref = element.get("ref")
        for fc in safe_iterator(element, '{%s}filterColumn' % SHEET_MAIN_NS):
            filters = fc.find('{%s}filters' % SHEET_MAIN_NS)
            if filters is None:
                continue
            vals = [f.get("val") for f in safe_iterator(filters, '{%s}filter' % SHEET_MAIN_NS)]
            blank = filters.get("blank")
            self.ws.auto_filter.add_filter_column(fc.get("colId"), vals, blank=blank)
        for sc in safe_iterator(element, '{%s}sortCondition' % SHEET_MAIN_NS):
            self.ws.auto_filter.add_sort_condition(sc.get("ref"), sc.get("descending"))

def fast_parse(ws, xml_source, string_table, style_table, color_index=None):

    parser = WorkSheetParser(ws, xml_source, string_table, style_table, color_index)
    parser.parse()
    del parser


def read_worksheet(xml_source, parent, preset_title, string_table,
                   style_table, color_index=None, worksheet_path=None, keep_vba=False):
    """Read an xml worksheet"""
    if worksheet_path:
        ws = IterableWorksheet(parent, preset_title,
                worksheet_path, xml_source, string_table, style_table)
    else:
        ws = Worksheet(parent, preset_title)
        fast_parse(ws, xml_source, string_table, style_table, color_index)
    if keep_vba:
        ws.xml_source = xml_source
    return ws
