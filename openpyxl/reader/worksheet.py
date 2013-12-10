# file openpyxl/reader/worksheet.py

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

"""Reader for a single worksheet."""

# Python stdlib imports

# compatibility imports
from openpyxl.shared.compat import BytesIO, StringIO
from openpyxl.shared.compat import iterparse

# package imports
from openpyxl.cell import get_column_letter
from openpyxl.shared.xmltools import fromstring
from openpyxl.cell import Cell, coordinate_from_string
from openpyxl.worksheet import Worksheet, ColumnDimension, RowDimension
from openpyxl.shared.ooxml import SHEET_MAIN_NS
from openpyxl.style import Color
from openpyxl.styles.formatting import ConditionalFormatting

def _get_xml_iter(xml_source):

    if not hasattr(xml_source, 'name'):
        if hasattr(xml_source, 'decode'):
            return BytesIO(xml_source)
        else:
            return BytesIO(xml_source.encode('utf-8'))
    else:
        xml_source.seek(0)
        return xml_source

def read_dimension(xml_source):

    source = _get_xml_iter(xml_source)

    it = iterparse(source)

    smax_col = None
    smax_row = None
    smin_col = None
    smin_row = None

    for event, element in it:

        if element.tag == '{%s}dimension' % SHEET_MAIN_NS:
            ref = element.get('ref')

            if ':' in ref:
                min_range, max_range = ref.split(':')
            else:
                min_range = max_range = ref

            min_col, min_row = coordinate_from_string(min_range)
            max_col, max_row = coordinate_from_string(max_range)

            return min_col, min_row, max_col, max_row

        if element.tag == '{%s}c' % SHEET_MAIN_NS:
            # Supposedly the dimension is mandatory, but in practice it can be
            # left off sometimes, if so, observe the max/min extants and return
            # those instead.
            col, row = coordinate_from_string(element.get('r'))
            if smin_row is None:
                # initialize the observed max/min values
                smin_col = smax_col = col
                smin_row = smax_row = row
            else:
                # Keep track of the seen max and min (fallback if there's no dimension)
                smin_col = min(smin_col, col)
                smin_row = min(smin_row, row)
                smax_col = max(smax_col, col)
                smax_row = max(smax_row, row)
        else:
            element.clear()

    return smin_col, smin_row, smax_col, smax_row

def filter_cells(pair):
    (event, element) = pair

    return element.tag == '{%s}c' % SHEET_MAIN_NS

def fast_parse(ws, xml_source, string_table, style_table, color_index=None):

    guess_types = ws.parent._guess_types
    data_only = ws.parent.data_only

    source = _get_xml_iter(xml_source)
    it = iterparse(source) # parses sheet tag by tag
    for event, element in filter(filter_cells, it):

        value = element.findtext('{%s}v' % SHEET_MAIN_NS)
        formula = element.find('{%s}f' % SHEET_MAIN_NS)

        coordinate = element.get('r')
        style_id = element.get('s')
        if style_id is not None:
            ws._styles[coordinate] = style_table.get(int(style_id))

        if value is not None:
            data_type = element.get('t', 'n')
            if data_type == Cell.TYPE_STRING:
                value = string_table.get(int(value))
            if formula is not None and not data_only:
                if formula.text:
                    value = "=" + str(formula.text)
                else:
                    value = "="
                formula_type = formula.get('t')
                if formula_type:
                    ws.formula_attributes[coordinate] = {'t': formula_type}
                    if formula.get('si'):  # Shared group index for shared formulas
                        ws.formula_attributes[coordinate]['si'] = formula.get('si')
                    if formula.get('ref'):  # Range for shared formulas
                        ws.formula_attributes[coordinate]['ref'] = formula.get('ref')
            if not guess_types and formula is None:
                ws.cell(coordinate).set_explicit_value(value=value, data_type=data_type)
            else:
                ws.cell(coordinate).value = value

        # to avoid memory exhaustion, clear the item after use
        element.clear()


    for event, element in it:
        mergeCells = element.find('{%s}mergeCells' % SHEET_MAIN_NS)
        if mergeCells is not None:
            for mergeCell in mergeCells.findall('{%s}mergeCell' % SHEET_MAIN_NS):
                ws.merge_cells(mergeCell.get('ref'))


    for event, element in it:
        cols = element.find('{%s}cols' % SHEET_MAIN_NS)
        if cols is not None:
            colNodes = cols.findall('{%s}col' % SHEET_MAIN_NS)
            for col in colNodes:
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
                        outline = col.get('outlineLevel')
                        collapsed = col.get('collapsed') == '1'
                        style_index =  style_table.get(int(col.get('style', 0)))
                        if column not in ws.column_dimensions:
                            new_column = ColumnDimension(index=column,
                                                         width=width, auto_size=auto_size,
                                                         visible=visible, outline_level=outline,
                                                         collapsed=collapsed, style_index=style_index)
                            ws.column_dimensions[column] = new_column


    for event, element in it:
        sheetData = element.find('{%s}sheetData' % SHEET_MAIN_NS)
        if sheetData is not None:
            rowNodes = sheetData.findall('{%s}row' % SHEET_MAIN_NS)
            for row in rowNodes:
                rowId = int(row.get('r'))
                if rowId not in ws.row_dimensions:
                    ws.row_dimensions[rowId] = RowDimension(rowId)
                ht = row.get('ht')
                if ht is not None:
                    ws.row_dimensions[rowId].height = float(ht)


    for event, element in it:
        printOptions = element.find('{%s}printOptions' % SHEET_MAIN_NS)
        if printOptions is not None:
            hc = printOptions.get('horizontalCentered')
            if hc is not None:
                ws.page_setup.horizontalCentered = hc
            vc = printOptions.get('verticalCentered')
            if vc is not None:
                ws.page_setup.verticalCentered = vc


    for event, element in it:
        pageMargins = element.find('{%s}pageMargins' % SHEET_MAIN_NS)
        if pageMargins is not None:
            for key in ("left", "right", "top", "bottom", "header", "footer"):
                value = pageMargins.get(key)
                if value is not None:
                    setattr(ws.page_margins, key, value)



    for event, element in it:
        pageSetup = element.find('{%s}pageSetup' % SHEET_MAIN_NS)
        if pageSetup is not None:
            for key in ("orientation", "paperSize", "scale", "fitToPage",
                        "fitToHeight", "fitToWidth", "firstPageNumber",
                        "useFirstPageNumber"):
                value = pageSetup.get(key, value)
                if value is not None:
                    setattr(ws.page_setup, key, value)


    for event, element in it:
        headerFooter = element.find('{%s}headerFooter' % SHEET_MAIN_NS)
        if headerFooter is not None:
            oddHeader = headerFooter.find('{%s}oddHeader' % SHEET_MAIN_NS)
            if oddHeader is not None and oddHeader.text is not None:
                ws.header_footer.setHeader(oddHeader.text)
            oddFooter = headerFooter.find('{%s}oddFooter' % SHEET_MAIN_NS)
            if oddFooter is not None and oddFooter.text is not None:
                ws.header_footer.setFooter(oddFooter.text)


    for event, element in it:
        conditionalFormattingNodes = element.findall('{%s}conditionalFormatting' % SHEET_MAIN_NS)
        rules = {}
        for cf in conditionalFormattingNodes:
            if not cf.get('sqref'):
                # Potentially flag - this attribute should always be present.
                continue
            range_string = cf.get('sqref')
            cfRules = cf.findall('{%s}cfRule' % SHEET_MAIN_NS)
            rules[range_string] = []
            for cfRule in cfRules:
                if not cfRule.get('type') or cfRule.get('type') == 'dataBar':
                    # dataBar conditional formatting isn't supported, as it relies on the complex <extLst> tag
                    continue
                rule = {'type': cfRule.get('type')}
                for attr in ConditionalFormatting.rule_attributes:
                    if cfRule.get(attr) is not None:
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
                        if color_index and color.get('indexed') is not None and 0 <= int(color.get('indexed')) < len(color_index):
                            c.index = color_index[int(color.get('indexed'))]
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

                rules[range_string].append(rule)
        if len(rules):
            ws.conditional_formatting.setRules(rules)

from openpyxl.reader.iter_worksheet import IterableWorksheet

def read_worksheet(xml_source, parent, preset_title, string_table,
                   style_table, color_index=None, workbook_name=None, sheet_codename=None, keep_vba=False):
    """Read an xml worksheet"""
    if workbook_name and sheet_codename:
        ws = IterableWorksheet(parent, preset_title, workbook_name,
                sheet_codename, xml_source, string_table)
    else:
        ws = Worksheet(parent, preset_title)
        fast_parse(ws, xml_source, string_table, style_table, color_index)
    if keep_vba:
        ws.xml_source = xml_source
    return ws
