# file openpyxl/tests/test_style.py

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

# Python stdlib imports
import os.path
import datetime

# compatibility imports
from openpyxl.shared.compat import BytesIO, StringIO

# package imports
from openpyxl.reader.excel import load_workbook
from openpyxl.reader.style import read_style_table
from openpyxl.workbook import Workbook
from openpyxl.writer.worksheet import write_worksheet
from openpyxl.writer.excel import save_virtual_workbook
from openpyxl.writer.styles import StyleWriter
from openpyxl.style import NumberFormat, Border, Color, Font, HashableObject

# test imports
from nose.tools import eq_, ok_, assert_false
from openpyxl.tests.helper import DATADIR, assert_equals_file_content, get_xml, compare_xml

class TestCreateStyle(object):

    @classmethod
    def setup_class(cls):
        now = datetime.datetime.now()
        cls.workbook = Workbook()
        cls.worksheet = cls.workbook.create_sheet()
        cls.worksheet.cell(coordinate='A1').value = '12.34%'
        cls.worksheet.cell(coordinate='B4').value = now
        cls.worksheet.cell(coordinate='B5').value = now
        cls.worksheet.cell(coordinate='C14').value = 'This is a test'
        cls.worksheet.cell(coordinate='D9').value = '31.31415'
        cls.worksheet.cell(coordinate='D9').style.number_format.format_code = NumberFormat.FORMAT_NUMBER_00
        cls.writer = StyleWriter(cls.workbook)

    def test_create_style_table(self):
        eq_(3, len(self.writer.style_table))

    def test_write_style_table(self):
        reference_file = os.path.join(DATADIR, 'writer', 'expected', 'simple-styles.xml')
        #assert_equals_file_content(reference_file, self.writer.write_table())


class TestStyleWriter(object):

    def setup(self):

        self.workbook = Workbook()
        self.worksheet = self.workbook.create_sheet()

    def test_no_style(self):

        w = StyleWriter(self.workbook)
        eq_(0, len(w.style_table))

    def test_nb_style(self):

        for i in range(1, 6):
            self.worksheet.cell(row=1, column=i).style.font.size += i
        w = StyleWriter(self.workbook)
        eq_(5, len(w.style_table))

        self.worksheet.cell('A10').style.borders.top = Border.BORDER_THIN
        w = StyleWriter(self.workbook)
        eq_(6, len(w.style_table))

    def test_style_unicity(self):

        for i in range(1, 6):
            self.worksheet.cell(row=1, column=i).style.font.bold = True
        w = StyleWriter(self.workbook)
        eq_(1, len(w.style_table))

    def test_fonts(self):

        self.worksheet.cell('A1').style.font.size = 12
        self.worksheet.cell('A1').style.font.bold = True
        w = StyleWriter(self.workbook)
        w._write_fonts()
        expected = '<?xml version=\'1.0\'?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11" /><color theme="1" /><name val="Calibri" /><family val="2" /><scheme val="minor" /></font><font><sz val="12" /><color rgb="FF000000" /><name val="Calibri" /><family val="2" /><b /></font></fonts></styleSheet>'
        xml = get_xml(w._root)
        diff = compare_xml(xml, expected)
        assert_false(diff)

    def test_fonts_with_underline(self):
        self.worksheet.cell('A1').style.font.size = 12
        self.worksheet.cell('A1').style.font.bold = True
        self.worksheet.cell('A1').style.font.underline = Font.UNDERLINE_SINGLE
        w = StyleWriter(self.workbook)
        w._write_fonts()
        expected = '<?xml version=\'1.0\'?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11" /><color theme="1" /><name val="Calibri" /><family val="2" /><scheme val="minor" /></font><font><sz val="12" /><color rgb="FF000000" /><name val="Calibri" /><family val="2" /><b /><u /></font></fonts></styleSheet>'
        xml = get_xml(w._root)
        diff = compare_xml(xml, expected)
        assert_false(diff)

    def test_fills(self):

        self.worksheet.cell('A1').style.fill.fill_type = 'solid'
        self.worksheet.cell('A1').style.fill.start_color.index = Color.DARKYELLOW
        w = StyleWriter(self.workbook)
        w._write_fills()
        expected = '<?xml version=\'1.0\'?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fills count="3"><fill><patternFill patternType="none" /></fill><fill><patternFill patternType="gray125" /></fill><fill><patternFill patternType="solid"><fgColor rgb="FF808000" /></patternFill></fill></fills></styleSheet>'
        xml = get_xml(w._root)
        diff = compare_xml(xml, expected)
        assert_false(diff)

    def test_borders(self):

        self.worksheet.cell('A1').style.borders.top.border_style = Border.BORDER_THIN
        self.worksheet.cell('A1').style.borders.top.color.index = Color.DARKYELLOW
        w = StyleWriter(self.workbook)
        w._write_borders()
        expected = '<?xml version=\'1.0\'?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><borders count="2"><border><left /><right /><top /><bottom /><diagonal /></border><border><left /><right /><top style="thin"><color rgb="FF808000" /></top><bottom /><diagonal /></border></borders></styleSheet>'
        xml = get_xml(w._root)
        diff = compare_xml(xml, expected)
        assert_false(diff)

    def test_write_cell_xfs_1(self):

        self.worksheet.cell('A1').style.font.size = 12
        w = StyleWriter(self.workbook)
        ft = w._write_fonts()
        nft = w._write_number_formats()
        w._write_cell_xfs(nft, ft, {}, {})
        xml = get_xml(w._root)
        ok_('applyFont="1"' in xml)
        ok_('applyFill="1"' not in xml)
        ok_('applyBorder="1"' not in xml)
        ok_('applyAlignment="1"' not in xml)

    def test_alignment(self):
        self.worksheet.cell('A1').style.alignment.horizontal = 'center'
        self.worksheet.cell('A1').style.alignment.vertical = 'center'
        w = StyleWriter(self.workbook)
        nft = w._write_number_formats()
        w._write_cell_xfs(nft, {}, {}, {})
        xml = get_xml(w._root)
        ok_('applyAlignment="1"' in xml)
        ok_('horizontal="center"' in xml)
        ok_('vertical="center"' in xml)

    def test_alignment_rotation(self):
        self.worksheet.cell('A1').style.alignment.vertical = 'center'
        self.worksheet.cell('A1').style.alignment.text_rotation = 90
        self.worksheet.cell('A2').style.alignment.vertical = 'center'
        self.worksheet.cell('A2').style.alignment.text_rotation = 135
        self.worksheet.cell('A3').style.alignment.text_rotation = -34
        w = StyleWriter(self.workbook)
        nft = w._write_number_formats()
        w._write_cell_xfs(nft, {}, {}, {})
        xml = get_xml(w._root)
        ok_('textRotation="90"' in xml)
        ok_('textRotation="135"' in xml)
        ok_('textRotation="124"' in xml)

    def test_alignment_indent(self):
        self.worksheet.cell('A1').style.alignment.indent = 1
        self.worksheet.cell('A2').style.alignment.indent = 4
        self.worksheet.cell('A3').style.alignment.indent = 0
        self.worksheet.cell('A3').style.alignment.indent = -1
        w = StyleWriter(self.workbook)
        nft = w._write_number_formats()
        w._write_cell_xfs(nft, {}, {}, {})
        xml = get_xml(w._root)
        ok_('indent="1"' in xml)
        ok_('indent="4"' in xml)
        #Indents not greater than zero are ignored when writing
        ok_('indent="0"' not in xml)
        ok_('indent="-1"' not in xml)

    def test_conditional_formatting_write(self):
        self.worksheet.conditional_formatting.add2ColorScale('A1:A10', 'min', None, 'FFAA0000', 'max', None, 'FF00AA00')
        self.worksheet.conditional_formatting.add3ColorScale('B1:B10', 'percentile', 10, 'FFAA0000', 'percentile', 50,
                                                             'FF0000AA', 'percentile', 90, 'FF00AA00')
        xml = write_worksheet(self.worksheet, None, None)
        ok_('<conditionalFormatting sqref="A1:A10"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"></cfvo><cfvo type="max"></cfvo><color rgb="FFAA0000"></color><color rgb="FF00AA00"></color></colorScale></cfRule></conditionalFormatting>' in xml)
        ok_('<conditionalFormatting sqref="B1:B10"><cfRule type="colorScale" priority="2"><colorScale><cfvo type="percentile" val="10"></cfvo><cfvo type="percentile" val="50"></cfvo><cfvo type="percentile" val="90"></cfvo><color rgb="FFAA0000"></color><color rgb="FF0000AA"></color><color rgb="FF00AA00"></color></colorScale></cfRule></conditionalFormatting>' in xml)


#def test_format_comparisions():
#    format1 = NumberFormat()
#    format2 = NumberFormat()
#    format3 = NumberFormat()
#    format1.format_code = 'm/d/yyyy'
#    format2.format_code = 'm/d/yyyy'
#    format3.format_code = 'mm/dd/yyyy'
#    assert not format1 < format2
#    assert format1 < format3
#    assert format1 == format2
#    assert format1 != format3


def test_builtin_format():
    nFormat = NumberFormat()
    nFormat.format_code = '0.00'
    eq_(nFormat.builtin_format_code(2), nFormat._format_code)


def test_read_style():
    reference_file = os.path.join(DATADIR, 'reader', 'simple-styles.xml')

    handle = open(reference_file, 'r')
    try:
        content = handle.read()
    finally:
        handle.close()
    style_properties = read_style_table(content)
    style_table = style_properties['table']
    eq_(4, len(style_table))
    eq_(NumberFormat._BUILTIN_FORMATS[9], style_table[1].number_format.format_code)
    eq_('yyyy-mm-dd', style_table[2].number_format.format_code)


def test_read_complex_style():
    reference_file = os.path.join(DATADIR, 'reader', 'complex-styles.xlsx')
    wb = load_workbook(reference_file)
    ws = wb.get_active_sheet()
    eq_(ws.column_dimensions['A'].width, 31.1640625)
    eq_(ws.cell('A2').style.font.name, 'Arial')
    eq_(ws.cell('A2').style.font.size, '10')
    eq_(ws.cell('A2').style.font.bold, False)
    eq_(ws.cell('A2').style.font.italic, False)
    eq_(ws.cell('A3').style.font.name, 'Arial')
    eq_(ws.cell('A3').style.font.size, '12')
    eq_(ws.cell('A3').style.font.bold, True)
    eq_(ws.cell('A3').style.font.italic, False)
    eq_(ws.cell('A4').style.font.name, 'Arial')
    eq_(ws.cell('A4').style.font.size, '14')
    eq_(ws.cell('A4').style.font.bold, False)
    eq_(ws.cell('A4').style.font.italic, True)
    eq_(ws.cell('A5').style.font.color.index, 'FF3300FF')
    eq_(ws.cell('A6').style.font.color.index, 'theme:9:')
    eq_(ws.cell('A7').style.fill.start_color.index, 'FFFFFF66')
    eq_(ws.cell('A8').style.fill.start_color.index, 'theme:8:')
    eq_(ws.cell('A9').style.alignment.horizontal, 'left')
    eq_(ws.cell('A10').style.alignment.horizontal, 'right')
    eq_(ws.cell('A11').style.alignment.horizontal, 'center')
    eq_(ws.cell('A12').style.alignment.vertical, 'top')
    eq_(ws.cell('A13').style.alignment.vertical, 'center')
    eq_(ws.cell('A14').style.alignment.vertical, 'bottom')
    eq_(ws.cell('A15').style.number_format._format_code, '0.00')
    eq_(ws.cell('A16').style.number_format._format_code, 'mm-dd-yy')
    eq_(ws.cell('A17').style.number_format._format_code, '0.00%')
    eq_('A18:B18' in ws._merged_cells, True)
    eq_(ws.cell('B18').merged, True)
    eq_(ws.cell('A19').style.borders.top.color.index, 'FF006600')
    eq_(ws.cell('A19').style.borders.bottom.color.index, 'FF006600')
    eq_(ws.cell('A19').style.borders.left.color.index, 'FF006600')
    eq_(ws.cell('A19').style.borders.right.color.index, 'FF006600')
    eq_(ws.cell('A21').style.borders.top.color.index, 'theme:7:')
    eq_(ws.cell('A21').style.borders.bottom.color.index, 'theme:7:')
    eq_(ws.cell('A21').style.borders.left.color.index, 'theme:7:')
    eq_(ws.cell('A21').style.borders.right.color.index, 'theme:7:')
    eq_(ws.cell('A23').style.fill.start_color.index, 'FFCCCCFF')
    eq_(ws.cell('A23').style.borders.top.color.index, 'theme:6:')
    eq_('A23:B24' in ws._merged_cells, True)
    eq_(ws.cell('A24').merged, True)
    eq_(ws.cell('B23').merged, True)
    eq_(ws.cell('B24').merged, True)
    eq_(ws.cell('A25').style.alignment.wrap_text, True)
    eq_(ws.cell('A26').style.alignment.shrink_to_fit, True)


def compare_complex(a, b):
    if isinstance(a, list):
        if not isinstance(b, list):
            return False
        else:
            for i, v in enumerate(a):
                if not compare_complex(v, b[i]):
                    return False
    elif isinstance(a, dict):
        if not isinstance(b, dict):
            return False
        else:
            for k in a.iterkeys():
                if isinstance(a[k], (list, dict)):
                    if not compare_complex(a[k], b[k]):
                        return False
                elif a[k] != b[k]:
                    return False
    elif isinstance(a, HashableObject) or isinstance(b, HashableObject):
        if repr(a) != repr(b):
            return False
    elif a != b:
        return False
    return True


def test_conditional_formatting_read():
    reference_file = os.path.join(DATADIR, 'reader', 'conditional-formatting.xlsx')
    wb = load_workbook(reference_file)
    ws = wb.get_active_sheet()

    # First test the conditional formatting rules read
    ok_(compare_complex(ws.conditional_formatting.cf_rules['A1:A1048576'], [{'priority': '27', 'type': 'colorScale', 'colorScale': {'color': [Color('FFFF7128'), 'FFFFEF9C'], 'cfvo': [{'type': 'min'}, {'type': 'max'}]}}]))
    ok_(compare_complex(ws.conditional_formatting.cf_rules['B1:B10'], [{'priority': '26', 'type': 'colorScale', 'colorScale': {'color': ['theme:6:', 'theme:4:'], 'cfvo': [{'type': 'num', 'val': '3'}, {'type': 'num', 'val': '7'}]}}]))
    ok_(compare_complex(ws.conditional_formatting.cf_rules['C1:C10'], [{'priority': '25', 'type': 'colorScale', 'colorScale': {'color': ['FFFF7128', 'FFFFEF9C'], 'cfvo': [{'type': 'percent', 'val': '10'}, {'type': 'percent', 'val': '90'}]}}]))
    ok_(compare_complex(ws.conditional_formatting.cf_rules['D1:D10'], [{'priority': '24', 'type': 'colorScale', 'colorScale': {'color': ['theme:6:', 'theme:5:'], 'cfvo': [{'type': 'formula', 'val': '2'}, {'type': 'formula', 'val': '4'}]}}]))
    ok_(compare_complex(ws.conditional_formatting.cf_rules['E1:E10'], [{'priority': '23', 'type': 'colorScale', 'colorScale': {'color': ['FFFF7128', 'FFFFEF9C'], 'cfvo': [{'type': 'percentile', 'val': '10'}, {'type': 'percentile', 'val': '90'}]}}]))
    ok_(compare_complex(ws.conditional_formatting.cf_rules['F1:F10'], [{'priority': '22', 'type': 'colorScale', 'colorScale': {'color': ['FFFF7128', 'FFFFEB84', 'FF63BE7B'], 'cfvo': [{'type': 'min'}, {'type': 'percentile', 'val': '50'}, {'type': 'max'}]}}]))
    ok_(compare_complex(ws.conditional_formatting.cf_rules['G1:G10'], [{'priority': '21', 'type': 'colorScale', 'colorScale': {'color': ['theme:4:', 'FFFFEB84', 'theme:5:'], 'cfvo': [{'type': 'num', 'val': '0'}, {'type': 'percentile', 'val': '50'}, {'type': 'num', 'val': '10'}]}}]))
    ok_(compare_complex(ws.conditional_formatting.cf_rules['H1:H10'], [{'priority': '20', 'type': 'colorScale', 'colorScale': {'color': ['FFFF7128', 'FFFFEB84', 'FF63BE7B'], 'cfvo': [{'type': 'percent', 'val': '0'}, {'type': 'percent', 'val': '50'}, {'type': 'percent', 'val': '100'}]}}]))
    ok_(compare_complex(ws.conditional_formatting.cf_rules['I1:I10'], [{'priority': '19', 'type': 'colorScale', 'colorScale': {'color': ['FF0000FF', 'FFFF6600', 'FF008000'], 'cfvo': [{'type': 'formula', 'val': '2'}, {'type': 'formula', 'val': '7'}, {'type': 'formula', 'val': '9'}]}}]))
    ok_(compare_complex(ws.conditional_formatting.cf_rules['J1:J10'], [{'priority': '18', 'type': 'colorScale', 'colorScale': {'color': ['FFFF7128', 'FFFFEB84', 'FF63BE7B'], 'cfvo': [{'type': 'percentile', 'val': '10'}, {'type': 'percentile', 'val': '50'}, {'type': 'percentile', 'val': '90'}]}}]))
    ok_(compare_complex(ws.conditional_formatting.cf_rules['K1:K10'], []))  # K - M are dataBar conditional formatting, which are not
    ok_(compare_complex(ws.conditional_formatting.cf_rules['L1:L10'], []))  # handled at the moment, and should not load, but also
    ok_(compare_complex(ws.conditional_formatting.cf_rules['M1:M10'], []))  # should not interfere with the loading / saving of the file.
    ok_(compare_complex(ws.conditional_formatting.cf_rules['N1:N10'], [{'priority': '17', 'iconSet': {'cfvo': [{'type': 'percent', 'val': '0'}, {'type': 'percent', 'val': '33'}, {'type': 'percent', 'val': '67'}]}, 'type': 'iconSet'}]))
    ok_(compare_complex(ws.conditional_formatting.cf_rules['O1:O10'], [{'priority': '16', 'iconSet': {'cfvo': [{'type': 'percent', 'val': '0'}, {'type': 'num', 'val': '2'}, {'type': 'num', 'val': '4'}, {'type': 'num', 'val': '6'}], 'showValue': '0', 'iconSet': '4ArrowsGray', 'reverse': '1'}, 'type': 'iconSet'}]))
    ok_(compare_complex(ws.conditional_formatting.cf_rules['P1:P10'], [{'priority': '15', 'iconSet': {'cfvo': [{'type': 'percent', 'val': '0'}, {'type': 'percentile', 'val': '20'}, {'type': 'percentile', 'val': '40'}, {'type': 'percentile', 'val': '60'}, {'type': 'percentile', 'val': '80'}], 'iconSet': '5Rating'}, 'type': 'iconSet'}]))
    ok_(compare_complex(ws.conditional_formatting.cf_rules['Q1:Q10'], [{'text': '3', 'priority': '14', 'dxfId': '27', 'operator': 'containsText', 'formula': ['NOT(ISERROR(SEARCH("3",Q1)))'], 'type': 'containsText'}]))
    ok_(compare_complex(ws.conditional_formatting.cf_rules['R1:R10'], [{'operator': 'between', 'dxfId': '26', 'type': 'cellIs', 'formula': ['2', '7'], 'priority': '13'}]))
    ok_(compare_complex(ws.conditional_formatting.cf_rules['S1:S10'], [{'priority': '12', 'dxfId': '25', 'percent': '1', 'type': 'top10', 'rank': '10'}]))
    ok_(compare_complex(ws.conditional_formatting.cf_rules['T1:T10'], [{'priority': '11', 'dxfId': '24', 'type': 'top10', 'rank': '4', 'bottom': '1'}]))
    ok_(compare_complex(ws.conditional_formatting.cf_rules['U1:U10'], [{'priority': '10', 'dxfId': '23', 'type': 'aboveAverage'}]))
    ok_(compare_complex(ws.conditional_formatting.cf_rules['V1:V10'], [{'aboveAverage': '0', 'dxfId': '22', 'type': 'aboveAverage', 'priority': '9'}]))
    ok_(compare_complex(ws.conditional_formatting.cf_rules['W1:W10'], [{'priority': '8', 'dxfId': '21', 'type': 'aboveAverage', 'equalAverage': '1'}]))
    ok_(compare_complex(ws.conditional_formatting.cf_rules['X1:X10'], [{'aboveAverage': '0', 'dxfId': '20', 'priority': '7', 'type': 'aboveAverage', 'equalAverage': '1'}]))
    ok_(compare_complex(ws.conditional_formatting.cf_rules['Y1:Y10'], [{'priority': '6', 'dxfId': '19', 'type': 'aboveAverage', 'stdDev': '1'}]))
    ok_(compare_complex(ws.conditional_formatting.cf_rules['Z1:Z10'], [{'aboveAverage': '0', 'dxfId': '18', 'type': 'aboveAverage', 'stdDev': '1', 'priority': '5'}]))
    ok_(compare_complex(ws.conditional_formatting.cf_rules['AA1:AA10'], [{'priority': '4', 'dxfId': '17', 'type': 'aboveAverage', 'stdDev': '2'}]))
    ok_(compare_complex(ws.conditional_formatting.cf_rules['AB1:AB10'], [{'priority': '3', 'dxfId': '16', 'type': 'duplicateValues'}]))
    ok_(compare_complex(ws.conditional_formatting.cf_rules['AC1:AC10'], [{'priority': '2', 'dxfId': '15', 'type': 'uniqueValues'}]))
    ok_(compare_complex(ws.conditional_formatting.cf_rules['AD1:AD10'], [{'priority': '1', 'dxfId': '14', 'type': 'expression', 'formula': ['AD1>3']}]))

def test_change_existing_styles():
    reference_file = os.path.join(DATADIR, 'reader', 'complex-styles.xlsx')
    wb = load_workbook(reference_file)
    ws = wb.get_active_sheet()

    ws.column_dimensions['A'].width = 20
    ws.cell('A2').style.font.name = 'Times New Roman'
    ws.cell('A2').style.font.size = 12
    ws.cell('A2').style.font.bold = True
    ws.cell('A2').style.font.italic = True
    ws.cell('A3').style.font.name = 'Times New Roman'
    ws.cell('A3').style.font.size = 14
    ws.cell('A3').style.font.bold = False
    ws.cell('A3').style.font.italic = True
    ws.cell('A4').style.font.name = 'Times New Roman'
    ws.cell('A4').style.font.size = 16
    ws.cell('A4').style.font.bold = True
    ws.cell('A4').style.font.italic = False
    ws.cell('A5').style.font.color.index = 'FF66FF66'
    ws.cell('A6').style.font.color.index = 'theme:1:'
    ws.cell('A7').style.fill.start_color.index = 'FF330066'
    ws.cell('A8').style.fill.start_color.index = 'theme:2:'
    ws.cell('A9').style.alignment.horizontal = 'center'
    ws.cell('A10').style.alignment.horizontal = 'left'
    ws.cell('A11').style.alignment.horizontal = 'right'
    ws.cell('A12').style.alignment.vertical = 'bottom'
    ws.cell('A13').style.alignment.vertical = 'top'
    ws.cell('A14').style.alignment.vertical = 'center'
    ws.cell('A15').style.number_format._format_code = '0.00%'
    ws.cell('A16').style.number_format._format_code = '0.00'
    ws.cell('A17').style.number_format._format_code = 'mm-dd-yy'
    ws.unmerge_cells('A18:B18')
    ws.cell('A19').style.borders.top.color.index = 'FF006600'
    ws.cell('A19').style.borders.bottom.color.index = 'FF006600'
    ws.cell('A19').style.borders.left.color.index = 'FF006600'
    ws.cell('A19').style.borders.right.color.index = 'FF006600'
    ws.cell('A21').style.borders.top.color.index = 'theme:7:'
    ws.cell('A21').style.borders.bottom.color.index = 'theme:7:'
    ws.cell('A21').style.borders.left.color.index = 'theme:7:'
    ws.cell('A21').style.borders.right.color.index = 'theme:7:'
    ws.cell('A23').style.fill.start_color.index = 'FFCCCCFF'
    ws.cell('A23').style.borders.top.color.index = 'theme:6:'
    ws.unmerge_cells('A23:B24')
    ws.cell('A25').style.alignment.wrap_text = False
    ws.cell('A26').style.alignment.shrink_to_fit = False

    saved_wb = save_virtual_workbook(wb)
    new_wb = load_workbook(BytesIO(saved_wb))
    ws = new_wb.get_active_sheet()

    eq_(ws.column_dimensions['A'].width, 20.0)
    eq_(ws.cell('A2').style.font.name, 'Times New Roman')
    eq_(ws.cell('A2').style.font.size, '12')
    eq_(ws.cell('A2').style.font.bold, True)
    eq_(ws.cell('A2').style.font.italic, True)
    eq_(ws.cell('A3').style.font.name, 'Times New Roman')
    eq_(ws.cell('A3').style.font.size, '14')
    eq_(ws.cell('A3').style.font.bold, False)
    eq_(ws.cell('A3').style.font.italic, True)
    eq_(ws.cell('A4').style.font.name, 'Times New Roman')
    eq_(ws.cell('A4').style.font.size, '16')
    eq_(ws.cell('A4').style.font.bold, True)
    eq_(ws.cell('A4').style.font.italic, False)
    eq_(ws.cell('A5').style.font.color.index, 'FF66FF66')
    eq_(ws.cell('A6').style.font.color.index, 'theme:1:')
    eq_(ws.cell('A7').style.fill.start_color.index, 'FF330066')
    eq_(ws.cell('A8').style.fill.start_color.index, 'theme:2:')
    eq_(ws.cell('A9').style.alignment.horizontal, 'center')
    eq_(ws.cell('A10').style.alignment.horizontal, 'left')
    eq_(ws.cell('A11').style.alignment.horizontal, 'right')
    eq_(ws.cell('A12').style.alignment.vertical, 'bottom')
    eq_(ws.cell('A13').style.alignment.vertical, 'top')
    eq_(ws.cell('A14').style.alignment.vertical, 'center')
    eq_(ws.cell('A15').style.number_format._format_code, '0.00%')
    eq_(ws.cell('A16').style.number_format._format_code, '0.00')
    eq_(ws.cell('A17').style.number_format._format_code, 'mm-dd-yy')
    eq_('A18:B18' in ws._merged_cells, False)
    eq_(ws.cell('B18').merged, False)
    eq_(ws.cell('A19').style.borders.top.color.index, 'FF006600')
    eq_(ws.cell('A19').style.borders.bottom.color.index, 'FF006600')
    eq_(ws.cell('A19').style.borders.left.color.index, 'FF006600')
    eq_(ws.cell('A19').style.borders.right.color.index, 'FF006600')
    eq_(ws.cell('A21').style.borders.top.color.index, 'theme:7:')
    eq_(ws.cell('A21').style.borders.bottom.color.index, 'theme:7:')
    eq_(ws.cell('A21').style.borders.left.color.index, 'theme:7:')
    eq_(ws.cell('A21').style.borders.right.color.index, 'theme:7:')
    eq_(ws.cell('A23').style.fill.start_color.index, 'FFCCCCFF')
    eq_(ws.cell('A23').style.borders.top.color.index, 'theme:6:')
    eq_('A23:B24' in ws._merged_cells, False)
    eq_(ws.cell('A24').merged, False)
    eq_(ws.cell('B23').merged, False)
    eq_(ws.cell('B24').merged, False)
    eq_(ws.cell('A25').style.alignment.wrap_text, False)
    eq_(ws.cell('A26').style.alignment.shrink_to_fit, False)

    # Verify that previously duplicate styles remain the same
    eq_(ws.column_dimensions['C'].width, 31.1640625)
    eq_(ws.cell('C2').style.font.name, 'Arial')
    eq_(ws.cell('C2').style.font.size, '10')
    eq_(ws.cell('C2').style.font.bold, False)
    eq_(ws.cell('C2').style.font.italic, False)
    eq_(ws.cell('C3').style.font.name, 'Arial')
    eq_(ws.cell('C3').style.font.size, '12')
    eq_(ws.cell('C3').style.font.bold, True)
    eq_(ws.cell('C3').style.font.italic, False)
    eq_(ws.cell('C4').style.font.name, 'Arial')
    eq_(ws.cell('C4').style.font.size, '14')
    eq_(ws.cell('C4').style.font.bold, False)
    eq_(ws.cell('C4').style.font.italic, True)
    eq_(ws.cell('C5').style.font.color.index, 'FF3300FF')
    eq_(ws.cell('C6').style.font.color.index, 'theme:9:')
    eq_(ws.cell('C7').style.fill.start_color.index, 'FFFFFF66')
    eq_(ws.cell('C8').style.fill.start_color.index, 'theme:8:')
    eq_(ws.cell('C9').style.alignment.horizontal, 'left')
    eq_(ws.cell('C10').style.alignment.horizontal, 'right')
    eq_(ws.cell('C11').style.alignment.horizontal, 'center')
    eq_(ws.cell('C12').style.alignment.vertical, 'top')
    eq_(ws.cell('C13').style.alignment.vertical, 'center')
    eq_(ws.cell('C14').style.alignment.vertical, 'bottom')
    eq_(ws.cell('C15').style.number_format._format_code, '0.00')
    eq_(ws.cell('C16').style.number_format._format_code, 'mm-dd-yy')
    eq_(ws.cell('C17').style.number_format._format_code, '0.00%')
    eq_('C18:D18' in ws._merged_cells, True)
    eq_(ws.cell('D18').merged, True)
    eq_(ws.cell('C19').style.borders.top.color.index, 'FF006600')
    eq_(ws.cell('C19').style.borders.bottom.color.index, 'FF006600')
    eq_(ws.cell('C19').style.borders.left.color.index, 'FF006600')
    eq_(ws.cell('C19').style.borders.right.color.index, 'FF006600')
    eq_(ws.cell('C21').style.borders.top.color.index, 'theme:7:')
    eq_(ws.cell('C21').style.borders.bottom.color.index, 'theme:7:')
    eq_(ws.cell('C21').style.borders.left.color.index, 'theme:7:')
    eq_(ws.cell('C21').style.borders.right.color.index, 'theme:7:')
    eq_(ws.cell('C23').style.fill.start_color.index, 'FFCCCCFF')
    eq_(ws.cell('C23').style.borders.top.color.index, 'theme:6:')
    eq_('C23:D24' in ws._merged_cells, True)
    eq_(ws.cell('C24').merged, True)
    eq_(ws.cell('D23').merged, True)
    eq_(ws.cell('D24').merged, True)
    eq_(ws.cell('C25').style.alignment.wrap_text, True)
    eq_(ws.cell('C26').style.alignment.shrink_to_fit, True)


def test_read_cell_style():
    reference_file = os.path.join(DATADIR, 'reader', 'empty-workbook-styles.xml')
    handle = open(reference_file, 'r')
    try:
        content = handle.read()
    finally:
        handle.close()
    style_properties = read_style_table(content)
    style_table = style_properties['table']
    eq_(2, len(style_table))