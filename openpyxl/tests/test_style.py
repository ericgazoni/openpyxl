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

# Python stdlib imports
import os.path
import datetime
from functools import partial

import pytest

# compatibility imports
from openpyxl.compat import BytesIO

# package imports
from openpyxl.reader.excel import load_workbook
from openpyxl.reader.style import read_style_table
from openpyxl.workbook import Workbook
from openpyxl.writer.excel import save_virtual_workbook
from openpyxl.writer.styles import StyleWriter
from openpyxl.styles import NumberFormat, Border, Color, Font, Fill, Borders, Protection
from openpyxl.formatting import ConditionalFormatting
from openpyxl.formatting.rules import FormulaRule
from openpyxl.xml.functions import Element, SubElement, tostring
from openpyxl.xml.constants import SHEET_MAIN_NS

# test imports
from openpyxl.tests.helper import DATADIR, get_xml, compare_xml


class TestCreateStyle(object):

    @classmethod
    def setup_class(cls):
        now = datetime.datetime.now()
        cls.workbook = Workbook(guess_types=True)
        cls.worksheet = cls.workbook.create_sheet()
        cls.worksheet.cell(coordinate='A1').value = '12.34%'
        cls.worksheet.cell(coordinate='B4').value = now
        cls.worksheet.cell(coordinate='B5').value = now
        cls.worksheet.cell(coordinate='C14').value = 'This is a test'
        cls.worksheet.cell(coordinate='D9').value = '31.31415'
        cls.worksheet.cell(coordinate='D9').style.number_format.format_code = NumberFormat.FORMAT_NUMBER_00
        cls.worksheet.cell(coordinate='D9').style.protection.locked = Protection.PROTECTION_UNPROTECTED
        cls.worksheet.cell(coordinate='E1').style.protection.hidden = Protection.PROTECTION_UNPROTECTED
        cls.writer = StyleWriter(cls.workbook)

    def test_create_style_table(self):
        assert len(self.writer.style_table) == 4

    @pytest.mark.xfail
    def test_write_style_table(self):
        reference_file = os.path.join(DATADIR, 'writer', 'expected', 'simple-styles.xml')

class TestStyleWriter(object):

    def setup(self):
        self.workbook = Workbook()
        self.worksheet = self.workbook.create_sheet()

    def test_no_style(self):
        w = StyleWriter(self.workbook)
        assert len(w.style_table) == 0

    def test_nb_style(self):
        for i in range(1, 6):
            self.worksheet.cell(row=1, column=i).style.font.size += i
        w = StyleWriter(self.workbook)
        assert len(w.style_table) == 5

        self.worksheet.cell('A10').style.borders.top = Border.BORDER_THIN
        w = StyleWriter(self.workbook)
        assert len(w.style_table) == 6

    def test_style_unicity(self):
        for i in range(1, 6):
            self.worksheet.cell(row=1, column=i).style.font.bold = True
        w = StyleWriter(self.workbook)
        assert len(w.style_table) == 1

    def test_fonts(self):
        self.worksheet.cell('A1').style.font.size = 12
        self.worksheet.cell('A1').style.font.bold = True
        w = StyleWriter(self.workbook)
        w._write_fonts()
        xml = get_xml(w._root)
        diff = compare_xml(xml, """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fonts count="2">
            <font>
              <sz val="11" />
              <color theme="1" />
              <name val="Calibri" />
              <family val="2" />
              <scheme val="minor" />
            </font>
            <font>
              <sz val="12" />
              <color rgb="FF000000" />
              <name val="Calibri" />
              <family val="2" />
              <b />
            </font>
          </fonts>
        </styleSheet>
        """)
        assert diff is None, diff

    def test_fonts_with_underline(self):
        self.worksheet.cell('A1').style.font.size = 12
        self.worksheet.cell('A1').style.font.bold = True
        self.worksheet.cell('A1').style.font.underline = Font.UNDERLINE_SINGLE
        w = StyleWriter(self.workbook)
        w._write_fonts()
        xml = get_xml(w._root)
        diff = compare_xml(xml, """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fonts count="2">
            <font>
              <sz val="11" />
              <color theme="1" />
              <name val="Calibri" />
              <family val="2" />
              <scheme val="minor" />
            </font>
            <font>
              <sz val="12" />
              <color rgb="FF000000" />
              <name val="Calibri" />
              <family val="2" />
              <b />
              <u />
            </font>
          </fonts>
        </styleSheet>
        """)
        assert diff is None, diff

    def test_fills(self):
        self.worksheet.cell('A1').style.fill.fill_type = 'solid'
        self.worksheet.cell('A1').style.fill.start_color.index = Color.DARKYELLOW
        w = StyleWriter(self.workbook)
        w._write_fills()
        xml = get_xml(w._root)
        diff = compare_xml(xml, """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fills count="3">
            <fill>
              <patternFill patternType="none" />
            </fill>
            <fill>
              <patternFill patternType="gray125" />
            </fill>
            <fill>
              <patternFill patternType="solid">
                <fgColor rgb="FF808000" />
              </patternFill>
            </fill>
          </fills>
        </styleSheet>
        """)
        assert diff is None, diff

    def test_borders(self):
        self.worksheet.cell('A1').style.borders.top.border_style = Border.BORDER_THIN
        self.worksheet.cell('A1').style.borders.top.color.index = Color.DARKYELLOW
        w = StyleWriter(self.workbook)
        w._write_borders()
        xml = get_xml(w._root)
        diff = compare_xml(xml, """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <borders count="2">
            <border>
              <left />
              <right />
              <top />
              <bottom />
              <diagonal />
            </border>
            <border>
              <left />
              <right />
              <top style="thin">
                <color rgb="FF808000" />
              </top>
              <bottom />
              <diagonal />
            </border>
          </borders>
        </styleSheet>
        """)
        assert diff is None, diff

    @pytest.mark.parametrize("value, expected",
                             [
                                 ('FFFFFF', {'rgb': 'FFFFFF'}),
                                 ('theme:7:', {'theme': '7'}),
                                 ('theme:7:8', {'theme':'7', 'tint':'8'})
                             ])
    def test_unpack_color(self, value, expected):
        w = StyleWriter(self.workbook)
        root = Element("root")
        w._unpack_color(root, value)
        assert root.find('color') is not None
        assert root.find('color').attrib == expected


    def test_write_cell_xfs_1(self):
        self.worksheet.cell('A1').style.font.size = 12
        w = StyleWriter(self.workbook)
        ft = w._write_fonts()
        nft = w._write_number_formats()
        w._write_cell_xfs(nft, ft, {}, {})
        xml = get_xml(w._root)
        assert 'applyFont="1"' in xml
        assert 'applyFill="1"' not in xml
        assert 'applyBorder="1"' not in xml
        assert 'applyAlignment="1"' not in xml

    def test_alignment(self):
        self.worksheet.cell('A1').style.alignment.horizontal = 'center'
        self.worksheet.cell('A1').style.alignment.vertical = 'center'
        w = StyleWriter(self.workbook)
        nft = w._write_number_formats()
        w._write_cell_xfs(nft, {}, {}, {})
        xml = get_xml(w._root)
        assert 'applyAlignment="1"' in xml
        assert 'horizontal="center"' in xml
        assert 'vertical="center"' in xml

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
        assert 'textRotation="90"' in xml
        assert 'textRotation="135"' in xml
        assert 'textRotation="124"' in xml

    def test_alignment_indent(self):
        self.worksheet.cell('A1').style.alignment.indent = 1
        self.worksheet.cell('A2').style.alignment.indent = 4
        self.worksheet.cell('A3').style.alignment.indent = 0
        self.worksheet.cell('A3').style.alignment.indent = -1
        w = StyleWriter(self.workbook)
        nft = w._write_number_formats()
        w._write_cell_xfs(nft, {}, {}, {})
        xml = get_xml(w._root)
        assert 'indent="1"' in xml
        assert 'indent="4"' in xml
        #Indents not greater than zero are ignored when writing
        assert 'indent="0"' not in xml
        assert 'indent="-1"' not in xml

    def test_rewrite_styles(self):
        """Test to verify Bugfix # 46"""
        self.worksheet['A1'].value = 'Value'
        self.worksheet['B2'].value = '14%'
        saved_wb = save_virtual_workbook(self.workbook)
        second_wb = load_workbook(BytesIO(saved_wb))
        assert isinstance(second_wb, Workbook)
        ws = second_wb.get_sheet_by_name('Sheet1')
        assert ws.cell('A1').value == 'Value'
        ws['A2'].value = 'Bar!'
        saved_wb = save_virtual_workbook(second_wb)
        third_wb = load_workbook(BytesIO(saved_wb))
        assert third_wb

    def test_write_dxf(self):
        redFill = Fill()
        redFill.start_color.index = 'FFEE1111'
        redFill.end_color.index = 'FFEE1111'
        redFill.fill_type = Fill.FILL_SOLID
        whiteFont = Font()
        whiteFont.color.index = "FFFFFFFF"
        whiteFont.bold = True
        whiteFont.italic = True
        whiteFont.underline = 'single'
        whiteFont.strikethrough = True
        blueBorder = Borders()
        blueBorder.left.border_style = 'medium'
        blueBorder.left.color = Color(Color.BLUE)
        blueBorder.right.border_style = 'medium'
        blueBorder.right.color = Color(Color.BLUE)
        blueBorder.top.border_style = 'medium'
        blueBorder.top.color = Color(Color.BLUE)
        blueBorder.bottom.border_style = 'medium'
        blueBorder.bottom.color = Color(Color.BLUE)
        cf = ConditionalFormatting()
        cf.add('A1:A2', FormulaRule(formula="[A1=1]", font=whiteFont, border=blueBorder, fill=redFill))
        cf.setDxfStyles(self.workbook)
        assert len(self.workbook.style_properties['dxf_list']) == 1
        assert 'font' in self.workbook.style_properties['dxf_list'][0]
        assert 'border' in self.workbook.style_properties['dxf_list'][0]
        assert 'fill' in self.workbook.style_properties['dxf_list'][0]

        w = StyleWriter(self.workbook)
        w._write_dxfs()
        xml = get_xml(w._root)

        diff = compare_xml(xml, """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <dxfs count="1">
            <dxf>
              <font>
                <color rgb="FFFFFFFF" />
                <b val="1" />
                <i val="1" />
                <u val="single" />
                <strike />
              </font>
              <fill>
                <patternFill patternType="solid">
                  <fgColor rgb="FFEE1111" />
                  <bgColor rgb="FFEE1111" />
                </patternFill>
              </fill>
              <border>
                <left style="medium">
                    <color rgb="FF0000FF"></color>
                </left>
                <right style="medium">
                    <color rgb="FF0000FF"></color>
                </right>
                <top style="medium">
                    <color rgb="FF0000FF"></color>
                </top>
                <bottom style="medium">
                    <color rgb="FF0000FF"></color>
                </bottom>
            </border>
            </dxf>
          </dxfs>
        </styleSheet>
        """)
        assert diff is None, diff

    def test_protection(self):
        self.worksheet.cell('A1').style.protection.locked = Protection.PROTECTION_UNPROTECTED
        self.worksheet.cell('A1').style.protection.hidden = Protection.PROTECTION_UNPROTECTED
        w = StyleWriter(self.workbook)
        nft = w._write_number_formats()
        w._write_cell_xfs(nft, {}, {}, {})
        xml = get_xml(w._root)
        assert 'protection' in xml
        assert 'locked="0"' in xml
        assert 'hidden="0"' in xml


def test_read_style():
    reference_file = os.path.join(DATADIR, 'reader', 'simple-styles.xml')

    handle = open(reference_file, 'r')
    try:
        content = handle.read()
    finally:
        handle.close()
    style_properties = read_style_table(content)
    style_table = style_properties['table']
    assert len(style_table) == 4
    assert NumberFormat._BUILTIN_FORMATS[9] == style_table[1].number_format
    assert 'yyyy-mm-dd' == style_table[2].number_format


def test_read_complex_style():
    reference_file = os.path.join(DATADIR, 'reader', 'complex-styles.xlsx')
    wb = load_workbook(reference_file)
    ws = wb.get_active_sheet()
    assert ws.column_dimensions['A'].width == 31.1640625

    style = partial(ws.get_style, read_only=True)
    assert style('I').fill.start_color.index == 'FF006600'
    assert style('I').font.color.index == 'FF3300FF'
    assert style('A2').font.name == 'Arial'
    assert style('A2').font.size == '10'
    assert not style('A2').font.bold
    assert not style('A2').font.italic
    assert style('A3').font.name == 'Arial'
    assert style('A3').font.size == '12'
    assert style('A3').font.bold
    assert not style('A3').font.italic
    assert style('A4').font.name == 'Arial'
    assert style('A4').font.size == '14'
    assert not style('A4').font.bold
    assert style('A4').font.italic
    assert style('A5').font.color.index == 'FF3300FF'
    assert style('A6').font.color.index == 'theme:9:'
    assert style('A7').fill.start_color.index == 'FFFFFF66'
    assert style('A8').fill.start_color.index == 'theme:8:'
    assert style('A9').alignment.horizontal == 'left'
    assert style('A10').alignment.horizontal == 'right'
    assert style('A11').alignment.horizontal == 'center'
    assert style('A12').alignment.vertical == 'top'
    assert style('A13').alignment.vertical == 'center'
    assert style('A14').alignment.vertical == 'bottom'
    assert style('A15').number_format == '0.00'
    assert style('A16').number_format == 'mm-dd-yy'
    assert style('A17').number_format == '0.00%'
    assert 'A18:B18' in ws._merged_cells
    assert ws.cell('B18').merged
    assert style('A19').borders.top.color.index == 'FF006600'
    assert style('A19').borders.bottom.color.index == 'FF006600'
    assert style('A19').borders.left.color.index == 'FF006600'
    assert style('A19').borders.right.color.index == 'FF006600'
    assert style('A21').borders.top.color.index == 'theme:7:'
    assert style('A21').borders.bottom.color.index == 'theme:7:'
    assert style('A21').borders.left.color.index == 'theme:7:'
    assert style('A21').borders.right.color.index == 'theme:7:'
    assert style('A23').fill.start_color.index == 'FFCCCCFF'
    assert style('A23').borders.top.color.index == 'theme:6:'
    assert 'A23:B24' in ws._merged_cells
    assert ws.cell('A24').merged
    assert ws.cell('B23').merged
    assert ws.cell('B24').merged
    assert style('A25').alignment.wrap_text
    assert style('A26').alignment.shrink_to_fit


def test_change_existing_styles():
    reference_file = os.path.join(DATADIR, 'reader', 'complex-styles.xlsx')
    wb = load_workbook(reference_file)
    ws = wb.get_active_sheet()

    ws.column_dimensions['A'].width = 20
    ws.get_style('I').fill.start_color.index = 'FF442200'
    ws.get_style('I').font.color.index = 'FF002244'
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
    ws.cell('A15').style.number_format.format_code = '0.00%'
    ws.cell('A16').style.number_format.format_code = '0.00'
    ws.cell('A17').style.number_format.format_code = 'mm-dd-yy'
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

    assert ws.column_dimensions['A'].width == 20.0

    style = partial(ws.get_style, read_only=True)

    assert ws.get_style('I').fill.start_color.index == 'FF442200'
    assert ws.get_style('I').font.color.index == 'FF002244'
    assert style('A2').font.name == 'Times New Roman'
    assert style('A2').font.size == '12'
    assert style('A2').font.bold
    assert style('A2').font.italic
    assert style('A3').font.name == 'Times New Roman'
    assert style('A3').font.size == '14'
    assert not style('A3').font.bold
    assert style('A3').font.italic
    assert style('A4').font.name == 'Times New Roman'
    assert style('A4').font.size == '16'
    assert style('A4').font.bold
    assert not style('A4').font.italic
    assert style('A5').font.color.index == 'FF66FF66'
    assert style('A6').font.color.index == 'theme:1:'
    assert style('A7').fill.start_color.index == 'FF330066'
    assert style('A8').fill.start_color.index == 'theme:2:'
    assert style('A9').alignment.horizontal == 'center'
    assert style('A10').alignment.horizontal == 'left'
    assert style('A11').alignment.horizontal == 'right'
    assert style('A12').alignment.vertical == 'bottom'
    assert style('A13').alignment.vertical == 'top'
    assert style('A14').alignment.vertical == 'center'
    assert style('A15').number_format == '0.00%'
    assert style('A16').number_format == '0.00'
    assert style('A17').number_format == 'mm-dd-yy'
    assert 'A18:B18' not in ws._merged_cells
    assert not ws.cell('B18').merged
    assert style('A19').borders.top.color.index == 'FF006600'
    assert style('A19').borders.bottom.color.index == 'FF006600'
    assert style('A19').borders.left.color.index == 'FF006600'
    assert style('A19').borders.right.color.index == 'FF006600'
    assert style('A21').borders.top.color.index == 'theme:7:'
    assert style('A21').borders.bottom.color.index == 'theme:7:'
    assert style('A21').borders.left.color.index == 'theme:7:'
    assert style('A21').borders.right.color.index == 'theme:7:'
    assert style('A23').fill.start_color.index == 'FFCCCCFF'
    assert style('A23').borders.top.color.index == 'theme:6:'
    assert 'A23:B24' not in ws._merged_cells
    assert not ws.cell('A24').merged
    assert not ws.cell('B23').merged
    assert not ws.cell('B24').merged
    assert not style('A25').alignment.wrap_text
    assert not style('A26').alignment.shrink_to_fit

    # Verify that previously duplicate styles remain the same
    assert ws.column_dimensions['C'].width == 31.1640625
    assert style('C2').font.name == 'Arial'
    assert style('C2').font.size == '10'
    assert not style('C2').font.bold
    assert not style('C2').font.italic
    assert style('C3').font.name == 'Arial'
    assert style('C3').font.size == '12'
    assert style('C3').font.bold
    assert not style('C3').font.italic
    assert style('C4').font.name == 'Arial'
    assert style('C4').font.size == '14'
    assert not style('C4').font.bold
    assert style('C4').font.italic
    assert style('C5').font.color.index == 'FF3300FF'
    assert style('C6').font.color.index == 'theme:9:'
    assert style('C7').fill.start_color.index == 'FFFFFF66'
    assert style('C8').fill.start_color.index == 'theme:8:'
    assert style('C9').alignment.horizontal == 'left'
    assert style('C10').alignment.horizontal == 'right'
    assert style('C11').alignment.horizontal == 'center'
    assert style('C12').alignment.vertical == 'top'
    assert style('C13').alignment.vertical == 'center'
    assert style('C14').alignment.vertical == 'bottom'
    assert style('C15').number_format == '0.00'
    assert style('C16').number_format == 'mm-dd-yy'
    assert style('C17').number_format == '0.00%'
    assert 'C18:D18' in ws._merged_cells
    assert ws.cell('D18').merged
    assert style('C19').borders.top.color.index == 'FF006600'
    assert style('C19').borders.bottom.color.index == 'FF006600'
    assert style('C19').borders.left.color.index == 'FF006600'
    assert style('C19').borders.right.color.index == 'FF006600'
    assert style('C21').borders.top.color.index == 'theme:7:'
    assert style('C21').borders.bottom.color.index == 'theme:7:'
    assert style('C21').borders.left.color.index == 'theme:7:'
    assert style('C21').borders.right.color.index == 'theme:7:'
    assert style('C23').fill.start_color.index == 'FFCCCCFF'
    assert style('C23').borders.top.color.index == 'theme:6:'
    assert 'C23:D24' in ws._merged_cells
    assert ws.cell('C24').merged
    assert ws.cell('D23').merged
    assert ws.cell('D24').merged
    assert style('C25').alignment.wrap_text
    assert style('C26').alignment.shrink_to_fit


def test_read_cell_style():
    reference_file = os.path.join(DATADIR, 'reader', 'empty-workbook-styles.xml')
    handle = open(reference_file, 'r')
    try:
        content = handle.read()
    finally:
        handle.close()
    style_properties = read_style_table(content)
    style_table = style_properties['table']
    assert len(style_table) == 2
