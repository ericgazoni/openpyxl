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
from openpyxl.shared.compat import BytesIO, StringIO, iterkeys

# package imports
from openpyxl.reader.excel import load_workbook
from openpyxl.reader.style import read_style_table
from openpyxl.shared.ooxml import ARC_STYLE
from openpyxl.workbook import Workbook
from openpyxl.writer.worksheet import write_worksheet
from openpyxl.writer.excel import save_virtual_workbook
from openpyxl.writer.styles import StyleWriter
from openpyxl.style import NumberFormat, Border, Color, Fill, Font, HashableObject, Borders

# test imports
from zipfile import ZIP_DEFLATED, ZipFile
from openpyxl.tests.helper import DATADIR, get_xml, compare_xml


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
        assert len(self.writer.style_table) == 3

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
        expected = """<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11" /><color theme="1" /><name val="Calibri" /><family val="2" /><scheme val="minor" /></font><font><sz val="12" /><color rgb="FF000000" /><name val="Calibri" /><family val="2" /><b /></font></fonts></styleSheet>"""
        xml = get_xml(w._root)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_fonts_with_underline(self):
        self.worksheet.cell('A1').style.font.size = 12
        self.worksheet.cell('A1').style.font.bold = True
        self.worksheet.cell('A1').style.font.underline = Font.UNDERLINE_SINGLE
        w = StyleWriter(self.workbook)
        w._write_fonts()
        expected = """<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11" /><color theme="1" /><name val="Calibri" /><family val="2" /><scheme val="minor" /></font><font><sz val="12" /><color rgb="FF000000" /><name val="Calibri" /><family val="2" /><b /><u /></font></fonts></styleSheet>"""
        xml = get_xml(w._root)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_fills(self):
        self.worksheet.cell('A1').style.fill.fill_type = 'solid'
        self.worksheet.cell('A1').style.fill.start_color.index = Color.DARKYELLOW
        w = StyleWriter(self.workbook)
        w._write_fills()
        expected = """<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fills count="3"><fill><patternFill patternType="none" /></fill><fill><patternFill patternType="gray125" /></fill><fill><patternFill patternType="solid"><fgColor rgb="FF808000" /></patternFill></fill></fills></styleSheet>"""
        xml = get_xml(w._root)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_borders(self):
        self.worksheet.cell('A1').style.borders.top.border_style = Border.BORDER_THIN
        self.worksheet.cell('A1').style.borders.top.color.index = Color.DARKYELLOW
        w = StyleWriter(self.workbook)
        w._write_borders()
        expected = """<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><borders count="2"><border><left /><right /><top /><bottom /><diagonal /></border><border><left /><right /><top style="thin"><color rgb="FF808000" /></top><bottom /><diagonal /></border></borders></styleSheet>"""
        xml = get_xml(w._root)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

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

    def test_conditional_formatting_add2ColorScale(self):
        self.worksheet.conditional_formatting.add2ColorScale('A1:A10', 'min', None, 'FFAA0000', 'max', None, 'FF00AA00')
        xml = write_worksheet(self.worksheet, None, None)
        expected = '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xml:space="preserve" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><outlinePr summaryRight="1" summaryBelow="1"></outlinePr></sheetPr><dimension ref="A1:A1"></dimension><sheetViews><sheetView workbookViewId="0"><selection sqref="A1" activeCell="A1"></selection></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"></sheetFormatPr><sheetData></sheetData><conditionalFormatting sqref="A1:A10"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"></cfvo><cfvo type="max"></cfvo><color rgb="FFAA0000"></color><color rgb="FF00AA00"></color></colorScale></cfRule></conditionalFormatting></worksheet>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_conditional_formatting_add3ColorScale(self):
        self.worksheet.conditional_formatting.add3ColorScale('B1:B10', 'percentile', 10, 'FFAA0000', 'percentile', 50,
                                                             'FF0000AA', 'percentile', 90, 'FF00AA00')
        xml = write_worksheet(self.worksheet, None, None)
        expected = '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xml:space="preserve" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><outlinePr summaryRight="1" summaryBelow="1"></outlinePr></sheetPr><dimension ref="A1:A1"></dimension><sheetViews><sheetView workbookViewId="0"><selection sqref="A1" activeCell="A1"></selection></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"></sheetFormatPr><sheetData></sheetData><conditionalFormatting sqref="B1:B10"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="percentile" val="10"></cfvo><cfvo type="percentile" val="50"></cfvo><cfvo type="percentile" val="90"></cfvo><color rgb="FFAA0000"></color><color rgb="FF0000AA"></color><color rgb="FF00AA00"></color></colorScale></cfRule></conditionalFormatting></worksheet>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_conditional_formatting_addCellIs_greaterThan(self):
        redFill = Fill()
        redFill.start_color.index = 'FFEE1111'
        redFill.end_color.index = 'FFEE1111'
        redFill.fill_type = Fill.FILL_SOLID
        self.worksheet.conditional_formatting.addCellIs('U10:U18', 'greaterThan', ['U$7'], True, self.workbook,
                                                        None, None, redFill)
        self.worksheet.conditional_formatting.addCellIs('V10:V18', '>', ['V$7'], True, self.workbook,
                                                        None, None, redFill)
        xml = write_worksheet(self.worksheet, None, None)
        expected = '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xml:space="preserve" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><outlinePr summaryRight="1" summaryBelow="1"></outlinePr></sheetPr><dimension ref="A1:A1"></dimension><sheetViews><sheetView workbookViewId="0"><selection sqref="A1" activeCell="A1"></selection></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"></sheetFormatPr><sheetData></sheetData><conditionalFormatting sqref="U10:U18"><cfRule priority="1" dxfId="0" type="cellIs" stopIfTrue="1" operator="greaterThan"><formula>U$7</formula></cfRule></conditionalFormatting><conditionalFormatting sqref="V10:V18"><cfRule priority="2" dxfId="1" type="cellIs" stopIfTrue="1" operator="greaterThan"><formula>V$7</formula></cfRule></conditionalFormatting></worksheet>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_conditional_formatting_addCellIs_greaterThanOrEqual(self):
        redFill = Fill()
        redFill.start_color.index = 'FFEE1111'
        redFill.end_color.index = 'FFEE1111'
        redFill.fill_type = Fill.FILL_SOLID
        self.worksheet.conditional_formatting.addCellIs('U10:U18', 'greaterThanOrEqual', ['U$7'], True, self.workbook,
                                                        None, None, redFill)
        self.worksheet.conditional_formatting.addCellIs('V10:V18', '>=', ['V$7'], True, self.workbook,
                                                        None, None, redFill)
        xml = write_worksheet(self.worksheet, None, None)
        expected = '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xml:space="preserve" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><outlinePr summaryRight="1" summaryBelow="1"></outlinePr></sheetPr><dimension ref="A1:A1"></dimension><sheetViews><sheetView workbookViewId="0"><selection sqref="A1" activeCell="A1"></selection></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"></sheetFormatPr><sheetData></sheetData><conditionalFormatting sqref="U10:U18"><cfRule priority="1" dxfId="0" type="cellIs" stopIfTrue="1" operator="greaterThanOrEqual"><formula>U$7</formula></cfRule></conditionalFormatting><conditionalFormatting sqref="V10:V18"><cfRule priority="2" dxfId="1" type="cellIs" stopIfTrue="1" operator="greaterThanOrEqual"><formula>V$7</formula></cfRule></conditionalFormatting></worksheet>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_conditional_formatting_addCellIs_lessThan(self):
        redFill = Fill()
        redFill.start_color.index = 'FFEE1111'
        redFill.end_color.index = 'FFEE1111'
        redFill.fill_type = Fill.FILL_SOLID
        self.worksheet.conditional_formatting.addCellIs('U10:U18', 'lessThan', ['U$7'], True, self.workbook,
                                                        None, None, redFill)
        self.worksheet.conditional_formatting.addCellIs('V10:V18', '<', ['V$7'], True, self.workbook,
                                                        None, None, redFill)
        xml = write_worksheet(self.worksheet, None, None)
        expected = '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xml:space="preserve" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><outlinePr summaryRight="1" summaryBelow="1"></outlinePr></sheetPr><dimension ref="A1:A1"></dimension><sheetViews><sheetView workbookViewId="0"><selection sqref="A1" activeCell="A1"></selection></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"></sheetFormatPr><sheetData></sheetData><conditionalFormatting sqref="U10:U18"><cfRule priority="1" dxfId="0" type="cellIs" stopIfTrue="1" operator="lessThan"><formula>U$7</formula></cfRule></conditionalFormatting><conditionalFormatting sqref="V10:V18"><cfRule priority="2" dxfId="1" type="cellIs" stopIfTrue="1" operator="lessThan"><formula>V$7</formula></cfRule></conditionalFormatting></worksheet>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_conditional_formatting_addCellIs_lessThanOrEqual(self):
        redFill = Fill()
        redFill.start_color.index = 'FFEE1111'
        redFill.end_color.index = 'FFEE1111'
        redFill.fill_type = Fill.FILL_SOLID
        self.worksheet.conditional_formatting.addCellIs('U10:U18', 'lessThanOrEqual', ['U$7'], True, self.workbook,
                                                        None, None, redFill)
        self.worksheet.conditional_formatting.addCellIs('V10:V18', '<=', ['V$7'], True, self.workbook,
                                                        None, None, redFill)
        xml = write_worksheet(self.worksheet, None, None)
        expected = '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xml:space="preserve" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><outlinePr summaryRight="1" summaryBelow="1"></outlinePr></sheetPr><dimension ref="A1:A1"></dimension><sheetViews><sheetView workbookViewId="0"><selection sqref="A1" activeCell="A1"></selection></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"></sheetFormatPr><sheetData></sheetData><conditionalFormatting sqref="U10:U18"><cfRule priority="1" dxfId="0" type="cellIs" stopIfTrue="1" operator="lessThanOrEqual"><formula>U$7</formula></cfRule></conditionalFormatting><conditionalFormatting sqref="V10:V18"><cfRule priority="2" dxfId="1" type="cellIs" stopIfTrue="1" operator="lessThanOrEqual"><formula>V$7</formula></cfRule></conditionalFormatting></worksheet>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_conditional_formatting_addCellIs_equal(self):
        redFill = Fill()
        redFill.start_color.index = 'FFEE1111'
        redFill.end_color.index = 'FFEE1111'
        redFill.fill_type = Fill.FILL_SOLID
        self.worksheet.conditional_formatting.addCellIs('U10:U18', 'equal', ['U$7'], True, self.workbook,
                                                        None, None, redFill)
        self.worksheet.conditional_formatting.addCellIs('V10:V18', '=', ['V$7'], True, self.workbook,
                                                        None, None, redFill)
        self.worksheet.conditional_formatting.addCellIs('W10:W18', '==', ['W$7'], True, self.workbook,
                                                        None, None, redFill)
        xml = write_worksheet(self.worksheet, None, None)
        expected = '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xml:space="preserve" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><outlinePr summaryRight="1" summaryBelow="1"></outlinePr></sheetPr><dimension ref="A1:A1"></dimension><sheetViews><sheetView workbookViewId="0"><selection sqref="A1" activeCell="A1"></selection></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"></sheetFormatPr><sheetData></sheetData><conditionalFormatting sqref="U10:U18"><cfRule priority="1" dxfId="0" type="cellIs" stopIfTrue="1" operator="equal"><formula>U$7</formula></cfRule></conditionalFormatting><conditionalFormatting sqref="V10:V18"><cfRule priority="2" dxfId="1" type="cellIs" stopIfTrue="1" operator="equal"><formula>V$7</formula></cfRule></conditionalFormatting><conditionalFormatting sqref="W10:W18"><cfRule priority="3" dxfId="2" type="cellIs" stopIfTrue="1" operator="equal"><formula>W$7</formula></cfRule></conditionalFormatting></worksheet>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_conditional_formatting_addCellIs_notEqual(self):
        redFill = Fill()
        redFill.start_color.index = 'FFEE1111'
        redFill.end_color.index = 'FFEE1111'
        redFill.fill_type = Fill.FILL_SOLID
        self.worksheet.conditional_formatting.addCellIs('U10:U18', 'notEqual', ['U$7'], True, self.workbook,
                                                        None, None, redFill)
        self.worksheet.conditional_formatting.addCellIs('V10:V18', '!=', ['V$7'], True, self.workbook,
                                                        None, None, redFill)
        xml = write_worksheet(self.worksheet, None, None)
        expected = '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xml:space="preserve" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><outlinePr summaryRight="1" summaryBelow="1"></outlinePr></sheetPr><dimension ref="A1:A1"></dimension><sheetViews><sheetView workbookViewId="0"><selection sqref="A1" activeCell="A1"></selection></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"></sheetFormatPr><sheetData></sheetData><conditionalFormatting sqref="U10:U18"><cfRule priority="1" dxfId="0" type="cellIs" stopIfTrue="1" operator="notEqual"><formula>U$7</formula></cfRule></conditionalFormatting><conditionalFormatting sqref="V10:V18"><cfRule priority="2" dxfId="1" type="cellIs" stopIfTrue="1" operator="notEqual"><formula>V$7</formula></cfRule></conditionalFormatting></worksheet>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_conditional_formatting_addCellIs_between(self):
        redFill = Fill()
        redFill.start_color.index = 'FFEE1111'
        redFill.end_color.index = 'FFEE1111'
        redFill.fill_type = Fill.FILL_SOLID
        self.worksheet.conditional_formatting.addCellIs('U10:U18', 'between', ['U$7', 'U$8'], True, self.workbook,
                                                        None, None, redFill)
        xml = write_worksheet(self.worksheet, None, None)
        expected = '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xml:space="preserve" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><outlinePr summaryRight="1" summaryBelow="1"></outlinePr></sheetPr><dimension ref="A1:A1"></dimension><sheetViews><sheetView workbookViewId="0"><selection sqref="A1" activeCell="A1"></selection></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"></sheetFormatPr><sheetData></sheetData><conditionalFormatting sqref="U10:U18"><cfRule priority="1" dxfId="0" type="cellIs" stopIfTrue="1" operator="between"><formula>U$7</formula><formula>U$8</formula></cfRule></conditionalFormatting></worksheet>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_conditional_formatting_addCellIs_notBetween(self):
        redFill = Fill()
        redFill.start_color.index = 'FFEE1111'
        redFill.end_color.index = 'FFEE1111'
        redFill.fill_type = Fill.FILL_SOLID
        self.worksheet.conditional_formatting.addCellIs('U10:U18', 'notBetween', ['U$7', 'U$8'], True, self.workbook,
                                                        None, None, redFill)
        xml = write_worksheet(self.worksheet, None, None)
        expected = '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xml:space="preserve" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><outlinePr summaryRight="1" summaryBelow="1"></outlinePr></sheetPr><dimension ref="A1:A1"></dimension><sheetViews><sheetView workbookViewId="0"><selection sqref="A1" activeCell="A1"></selection></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"></sheetFormatPr><sheetData></sheetData><conditionalFormatting sqref="U10:U18"><cfRule priority="1" dxfId="0" type="cellIs" stopIfTrue="1" operator="notBetween"><formula>U$7</formula><formula>U$8</formula></cfRule></conditionalFormatting></worksheet>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_conditional_formatting_addCustomRule(self):
        dxfId = self.worksheet.conditional_formatting.addDxfStyle(self.workbook, None, None, None)
        self.worksheet.conditional_formatting.addCustomRule('C1:C10',  {'type': 'expression', 'dxfId': dxfId, 'formula': ['ISBLANK(C1)'], 'stopIfTrue': '1'})
        xml = write_worksheet(self.worksheet, None, None)
        assert dxfId == 0
        expected = '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xml:space="preserve" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><outlinePr summaryRight="1" summaryBelow="1"></outlinePr></sheetPr><dimension ref="A1:A1"></dimension><sheetViews><sheetView workbookViewId="0"><selection sqref="A1" activeCell="A1"></selection></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"></sheetFormatPr><sheetData></sheetData><conditionalFormatting sqref="C1:C10"><cfRule dxfId="0" type="expression" stopIfTrue="1" priority="1"><formula>ISBLANK(C1)</formula></cfRule></conditionalFormatting></worksheet>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_conditional_formatting_addDxfStyle(self):
        fill = Fill()
        fill.start_color.index = 'FFEE1111'
        fill.end_color.index = 'FFEE1111'
        fill.fill_type = Fill.FILL_SOLID
        font = Font()
        font.name = 'Arial'
        font.size = 12
        font.bold = True
        font.underline = Font.UNDERLINE_SINGLE
        borders = Borders()
        borders.top.border_style = Border.BORDER_THIN
        borders.top.color.index = Color.DARKYELLOW
        borders.bottom.border_style = Border.BORDER_THIN
        borders.bottom.color.index = Color.BLACK
        dxfId = self.worksheet.conditional_formatting.addDxfStyle(self.workbook, font, borders, fill)
        assert dxfId == 0
        dxfId = self.worksheet.conditional_formatting.addDxfStyle(self.workbook, None, None, fill)
        assert dxfId == 1
        assert len(self.workbook.style_properties['dxf_list']) == 2
        #assert self.workbook.style_properties['dxf_list'][0]) == {'font': ['Arial':12:True:False:False:False:'single':False:'FF000000'], 'border': ['none':'FF000000':'none':'FF000000':'thin':'FF808000':'thin':'FF000000':'none':'FF000000':0:'none':'FF000000':'none':'FF000000':'none':'FF000000':'none':'FF000000':'none':'FF000000'], 'fill': ['solid':0:'FFEE1111':'FFEE1111']}
        #assert self.workbook.style_properties['dxf_list'][1] == {'fill': ['solid':0:'FFEE1111':'FFEE1111']}

    def test_conditional_formatting_setRules(self):
        rules = {'A1:A4': [{'type': 'colorScale', 'priority': '13',
                            'colorScale': {'cfvo': [{'type': 'min'}, {'type': 'max'}],
                                           'color': [Color('FFFF7128'), Color('FFFFEF9C')]}}]}
        self.worksheet.conditional_formatting.setRules(rules)
        xml = write_worksheet(self.worksheet, None, None)
        expected = '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xml:space="preserve" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><outlinePr summaryRight="1" summaryBelow="1"></outlinePr></sheetPr><dimension ref="A1:A1"></dimension><sheetViews><sheetView workbookViewId="0"><selection sqref="A1" activeCell="A1"></selection></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"></sheetFormatPr><sheetData></sheetData><conditionalFormatting sqref="A1:A4"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"></cfvo><cfvo type="max"></cfvo><color rgb="FFFF7128"></color><color rgb="FFFFEF9C"></color></colorScale></cfRule></conditionalFormatting></worksheet>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

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
    assert nFormat.builtin_format_code(2) == nFormat._format_code


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
    assert NumberFormat._BUILTIN_FORMATS[9] == style_table[1].number_format.format_code
    assert 'yyyy-mm-dd' == style_table[2].number_format.format_code


def test_read_complex_style():
    reference_file = os.path.join(DATADIR, 'reader', 'complex-styles.xlsx')
    wb = load_workbook(reference_file)
    ws = wb.get_active_sheet()
    assert ws.column_dimensions['A'].width == 31.1640625
    assert ws.cell('A2').style.font.name == 'Arial'
    assert ws.cell('A2').style.font.size == '10'
    assert not ws.cell('A2').style.font.bold
    assert not ws.cell('A2').style.font.italic
    assert ws.cell('A3').style.font.name == 'Arial'
    assert ws.cell('A3').style.font.size == '12'
    assert ws.cell('A3').style.font.bold
    assert not ws.cell('A3').style.font.italic
    assert ws.cell('A4').style.font.name == 'Arial'
    assert ws.cell('A4').style.font.size == '14'
    assert not ws.cell('A4').style.font.bold
    assert ws.cell('A4').style.font.italic
    assert ws.cell('A5').style.font.color.index == 'FF3300FF'
    assert ws.cell('A6').style.font.color.index == 'theme:9:'
    assert ws.cell('A7').style.fill.start_color.index == 'FFFFFF66'
    assert ws.cell('A8').style.fill.start_color.index == 'theme:8:'
    assert ws.cell('A9').style.alignment.horizontal == 'left'
    assert ws.cell('A10').style.alignment.horizontal == 'right'
    assert ws.cell('A11').style.alignment.horizontal == 'center'
    assert ws.cell('A12').style.alignment.vertical == 'top'
    assert ws.cell('A13').style.alignment.vertical == 'center'
    assert ws.cell('A14').style.alignment.vertical == 'bottom'
    assert ws.cell('A15').style.number_format._format_code == '0.00'
    assert ws.cell('A16').style.number_format._format_code == 'mm-dd-yy'
    assert ws.cell('A17').style.number_format._format_code == '0.00%'
    assert 'A18:B18' in ws._merged_cells
    assert ws.cell('B18').merged
    assert ws.cell('A19').style.borders.top.color.index == 'FF006600'
    assert ws.cell('A19').style.borders.bottom.color.index == 'FF006600'
    assert ws.cell('A19').style.borders.left.color.index == 'FF006600'
    assert ws.cell('A19').style.borders.right.color.index == 'FF006600'
    assert ws.cell('A21').style.borders.top.color.index == 'theme:7:'
    assert ws.cell('A21').style.borders.bottom.color.index == 'theme:7:'
    assert ws.cell('A21').style.borders.left.color.index == 'theme:7:'
    assert ws.cell('A21').style.borders.right.color.index == 'theme:7:'
    assert ws.cell('A23').style.fill.start_color.index == 'FFCCCCFF'
    assert ws.cell('A23').style.borders.top.color.index == 'theme:6:'
    assert 'A23:B24' in ws._merged_cells
    assert ws.cell('A24').merged
    assert ws.cell('B23').merged
    assert ws.cell('B24').merged
    assert ws.cell('A25').style.alignment.wrap_text
    assert ws.cell('A26').style.alignment.shrink_to_fit


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
            for k in iterkeys(a):
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
    assert compare_complex(ws.conditional_formatting.cf_rules['A1:A1048576'], [{'priority': '27', 'type': 'colorScale', 'colorScale': {'color': [Color('FFFF7128'), 'FFFFEF9C'], 'cfvo': [{'type': 'min'}, {'type': 'max'}]}}])
    assert compare_complex(ws.conditional_formatting.cf_rules['B1:B10'], [{'priority': '26', 'type': 'colorScale', 'colorScale': {'color': ['theme:6:', 'theme:4:'], 'cfvo': [{'type': 'num', 'val': '3'}, {'type': 'num', 'val': '7'}]}}])
    assert compare_complex(ws.conditional_formatting.cf_rules['C1:C10'], [{'priority': '25', 'type': 'colorScale', 'colorScale': {'color': ['FFFF7128', 'FFFFEF9C'], 'cfvo': [{'type': 'percent', 'val': '10'}, {'type': 'percent', 'val': '90'}]}}])
    assert compare_complex(ws.conditional_formatting.cf_rules['D1:D10'], [{'priority': '24', 'type': 'colorScale', 'colorScale': {'color': ['theme:6:', 'theme:5:'], 'cfvo': [{'type': 'formula', 'val': '2'}, {'type': 'formula', 'val': '4'}]}}])
    assert compare_complex(ws.conditional_formatting.cf_rules['E1:E10'], [{'priority': '23', 'type': 'colorScale', 'colorScale': {'color': ['FFFF7128', 'FFFFEF9C'], 'cfvo': [{'type': 'percentile', 'val': '10'}, {'type': 'percentile', 'val': '90'}]}}])
    assert compare_complex(ws.conditional_formatting.cf_rules['F1:F10'], [{'priority': '22', 'type': 'colorScale', 'colorScale': {'color': ['FFFF7128', 'FFFFEB84', 'FF63BE7B'], 'cfvo': [{'type': 'min'}, {'type': 'percentile', 'val': '50'}, {'type': 'max'}]}}])
    assert compare_complex(ws.conditional_formatting.cf_rules['G1:G10'], [{'priority': '21', 'type': 'colorScale', 'colorScale': {'color': ['theme:4:', 'FFFFEB84', 'theme:5:'], 'cfvo': [{'type': 'num', 'val': '0'}, {'type': 'percentile', 'val': '50'}, {'type': 'num', 'val': '10'}]}}])
    assert compare_complex(ws.conditional_formatting.cf_rules['H1:H10'], [{'priority': '20', 'type': 'colorScale', 'colorScale': {'color': ['FFFF7128', 'FFFFEB84', 'FF63BE7B'], 'cfvo': [{'type': 'percent', 'val': '0'}, {'type': 'percent', 'val': '50'}, {'type': 'percent', 'val': '100'}]}}])
    assert compare_complex(ws.conditional_formatting.cf_rules['I1:I10'], [{'priority': '19', 'type': 'colorScale', 'colorScale': {'color': ['FF0000FF', 'FFFF6600', 'FF008000'], 'cfvo': [{'type': 'formula', 'val': '2'}, {'type': 'formula', 'val': '7'}, {'type': 'formula', 'val': '9'}]}}])
    assert compare_complex(ws.conditional_formatting.cf_rules['J1:J10'], [{'priority': '18', 'type': 'colorScale', 'colorScale': {'color': ['FFFF7128', 'FFFFEB84', 'FF63BE7B'], 'cfvo': [{'type': 'percentile', 'val': '10'}, {'type': 'percentile', 'val': '50'}, {'type': 'percentile', 'val': '90'}]}}])
    assert compare_complex(ws.conditional_formatting.cf_rules['K1:K10'], [])  # K - M are dataBar conditional formatting, which are not
    assert compare_complex(ws.conditional_formatting.cf_rules['L1:L10'], [])  # handled at the moment, and should not load, but also
    assert compare_complex(ws.conditional_formatting.cf_rules['M1:M10'], [])  # should not interfere with the loading / saving of the file.
    assert compare_complex(ws.conditional_formatting.cf_rules['N1:N10'], [{'priority': '17', 'iconSet': {'cfvo': [{'type': 'percent', 'val': '0'}, {'type': 'percent', 'val': '33'}, {'type': 'percent', 'val': '67'}]}, 'type': 'iconSet'}])
    assert compare_complex(ws.conditional_formatting.cf_rules['O1:O10'], [{'priority': '16', 'iconSet': {'cfvo': [{'type': 'percent', 'val': '0'}, {'type': 'num', 'val': '2'}, {'type': 'num', 'val': '4'}, {'type': 'num', 'val': '6'}], 'showValue': '0', 'iconSet': '4ArrowsGray', 'reverse': '1'}, 'type': 'iconSet'}])
    assert compare_complex(ws.conditional_formatting.cf_rules['P1:P10'], [{'priority': '15', 'iconSet': {'cfvo': [{'type': 'percent', 'val': '0'}, {'type': 'percentile', 'val': '20'}, {'type': 'percentile', 'val': '40'}, {'type': 'percentile', 'val': '60'}, {'type': 'percentile', 'val': '80'}], 'iconSet': '5Rating'}, 'type': 'iconSet'}])
    assert compare_complex(ws.conditional_formatting.cf_rules['Q1:Q10'], [{'text': '3', 'priority': '14', 'dxfId': '27', 'operator': 'containsText', 'formula': ['NOT(ISERROR(SEARCH("3",Q1)))'], 'type': 'containsText'}])
    assert compare_complex(ws.conditional_formatting.cf_rules['R1:R10'], [{'operator': 'between', 'dxfId': '26', 'type': 'cellIs', 'formula': ['2', '7'], 'priority': '13'}])
    assert compare_complex(ws.conditional_formatting.cf_rules['S1:S10'], [{'priority': '12', 'dxfId': '25', 'percent': '1', 'type': 'top10', 'rank': '10'}])
    assert compare_complex(ws.conditional_formatting.cf_rules['T1:T10'], [{'priority': '11', 'dxfId': '24', 'type': 'top10', 'rank': '4', 'bottom': '1'}])
    assert compare_complex(ws.conditional_formatting.cf_rules['U1:U10'], [{'priority': '10', 'dxfId': '23', 'type': 'aboveAverage'}])
    assert compare_complex(ws.conditional_formatting.cf_rules['V1:V10'], [{'aboveAverage': '0', 'dxfId': '22', 'type': 'aboveAverage', 'priority': '9'}])
    assert compare_complex(ws.conditional_formatting.cf_rules['W1:W10'], [{'priority': '8', 'dxfId': '21', 'type': 'aboveAverage', 'equalAverage': '1'}])
    assert compare_complex(ws.conditional_formatting.cf_rules['X1:X10'], [{'aboveAverage': '0', 'dxfId': '20', 'priority': '7', 'type': 'aboveAverage', 'equalAverage': '1'}])
    assert compare_complex(ws.conditional_formatting.cf_rules['Y1:Y10'], [{'priority': '6', 'dxfId': '19', 'type': 'aboveAverage', 'stdDev': '1'}])
    assert compare_complex(ws.conditional_formatting.cf_rules['Z1:Z10'], [{'aboveAverage': '0', 'dxfId': '18', 'type': 'aboveAverage', 'stdDev': '1', 'priority': '5'}])
    assert compare_complex(ws.conditional_formatting.cf_rules['AA1:AA10'], [{'priority': '4', 'dxfId': '17', 'type': 'aboveAverage', 'stdDev': '2'}])
    assert compare_complex(ws.conditional_formatting.cf_rules['AB1:AB10'], [{'priority': '3', 'dxfId': '16', 'type': 'duplicateValues'}])
    assert compare_complex(ws.conditional_formatting.cf_rules['AC1:AC10'], [{'priority': '2', 'dxfId': '15', 'type': 'uniqueValues'}])
    assert compare_complex(ws.conditional_formatting.cf_rules['AD1:AD10'], [{'priority': '1', 'dxfId': '14', 'type': 'expression', 'formula': ['AD1>3']}])


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

    assert ws.column_dimensions['A'].width == 20.0
    assert ws.cell('A2').style.font.name == 'Times New Roman'
    assert ws.cell('A2').style.font.size == '12'
    assert ws.cell('A2').style.font.bold
    assert ws.cell('A2').style.font.italic
    assert ws.cell('A3').style.font.name == 'Times New Roman'
    assert ws.cell('A3').style.font.size == '14'
    assert not ws.cell('A3').style.font.bold
    assert ws.cell('A3').style.font.italic
    assert ws.cell('A4').style.font.name == 'Times New Roman'
    assert ws.cell('A4').style.font.size == '16'
    assert ws.cell('A4').style.font.bold
    assert not ws.cell('A4').style.font.italic
    assert ws.cell('A5').style.font.color.index == 'FF66FF66'
    assert ws.cell('A6').style.font.color.index == 'theme:1:'
    assert ws.cell('A7').style.fill.start_color.index == 'FF330066'
    assert ws.cell('A8').style.fill.start_color.index == 'theme:2:'
    assert ws.cell('A9').style.alignment.horizontal == 'center'
    assert ws.cell('A10').style.alignment.horizontal == 'left'
    assert ws.cell('A11').style.alignment.horizontal == 'right'
    assert ws.cell('A12').style.alignment.vertical == 'bottom'
    assert ws.cell('A13').style.alignment.vertical == 'top'
    assert ws.cell('A14').style.alignment.vertical == 'center'
    assert ws.cell('A15').style.number_format._format_code == '0.00%'
    assert ws.cell('A16').style.number_format._format_code == '0.00'
    assert ws.cell('A17').style.number_format._format_code == 'mm-dd-yy'
    assert 'A18:B18' not in ws._merged_cells
    assert not ws.cell('B18').merged
    assert ws.cell('A19').style.borders.top.color.index == 'FF006600'
    assert ws.cell('A19').style.borders.bottom.color.index == 'FF006600'
    assert ws.cell('A19').style.borders.left.color.index == 'FF006600'
    assert ws.cell('A19').style.borders.right.color.index == 'FF006600'
    assert ws.cell('A21').style.borders.top.color.index == 'theme:7:'
    assert ws.cell('A21').style.borders.bottom.color.index == 'theme:7:'
    assert ws.cell('A21').style.borders.left.color.index == 'theme:7:'
    assert ws.cell('A21').style.borders.right.color.index == 'theme:7:'
    assert ws.cell('A23').style.fill.start_color.index == 'FFCCCCFF'
    assert ws.cell('A23').style.borders.top.color.index == 'theme:6:'
    assert 'A23:B24' not in ws._merged_cells
    assert not ws.cell('A24').merged
    assert not ws.cell('B23').merged
    assert not ws.cell('B24').merged
    assert not ws.cell('A25').style.alignment.wrap_text
    assert not ws.cell('A26').style.alignment.shrink_to_fit

    # Verify that previously duplicate styles remain the same
    assert ws.column_dimensions['C'].width == 31.1640625
    assert ws.cell('C2').style.font.name == 'Arial'
    assert ws.cell('C2').style.font.size == '10'
    assert not ws.cell('C2').style.font.bold
    assert not ws.cell('C2').style.font.italic
    assert ws.cell('C3').style.font.name == 'Arial'
    assert ws.cell('C3').style.font.size == '12'
    assert ws.cell('C3').style.font.bold
    assert not ws.cell('C3').style.font.italic
    assert ws.cell('C4').style.font.name == 'Arial'
    assert ws.cell('C4').style.font.size == '14'
    assert not ws.cell('C4').style.font.bold
    assert ws.cell('C4').style.font.italic
    assert ws.cell('C5').style.font.color.index == 'FF3300FF'
    assert ws.cell('C6').style.font.color.index == 'theme:9:'
    assert ws.cell('C7').style.fill.start_color.index == 'FFFFFF66'
    assert ws.cell('C8').style.fill.start_color.index == 'theme:8:'
    assert ws.cell('C9').style.alignment.horizontal == 'left'
    assert ws.cell('C10').style.alignment.horizontal == 'right'
    assert ws.cell('C11').style.alignment.horizontal == 'center'
    assert ws.cell('C12').style.alignment.vertical == 'top'
    assert ws.cell('C13').style.alignment.vertical == 'center'
    assert ws.cell('C14').style.alignment.vertical == 'bottom'
    assert ws.cell('C15').style.number_format._format_code == '0.00'
    assert ws.cell('C16').style.number_format._format_code == 'mm-dd-yy'
    assert ws.cell('C17').style.number_format._format_code == '0.00%'
    assert 'C18:D18' in ws._merged_cells
    assert ws.cell('D18').merged
    assert ws.cell('C19').style.borders.top.color.index == 'FF006600'
    assert ws.cell('C19').style.borders.bottom.color.index == 'FF006600'
    assert ws.cell('C19').style.borders.left.color.index == 'FF006600'
    assert ws.cell('C19').style.borders.right.color.index == 'FF006600'
    assert ws.cell('C21').style.borders.top.color.index == 'theme:7:'
    assert ws.cell('C21').style.borders.bottom.color.index == 'theme:7:'
    assert ws.cell('C21').style.borders.left.color.index == 'theme:7:'
    assert ws.cell('C21').style.borders.right.color.index == 'theme:7:'
    assert ws.cell('C23').style.fill.start_color.index == 'FFCCCCFF'
    assert ws.cell('C23').style.borders.top.color.index == 'theme:6:'
    assert 'C23:D24' in ws._merged_cells
    assert ws.cell('C24').merged
    assert ws.cell('D23').merged
    assert ws.cell('D24').merged
    assert ws.cell('C25').style.alignment.wrap_text
    assert ws.cell('C26').style.alignment.shrink_to_fit


def test_parse_dxfs():
    reference_file = os.path.join(DATADIR, 'reader', 'conditional-formatting.xlsx')
    wb = load_workbook(reference_file)
    archive = ZipFile(reference_file, 'r', ZIP_DEFLATED)
    read_xml = archive.read(ARC_STYLE)

    # Verify length
    assert '<dxfs count="164">' in str(read_xml)
    assert len(wb.style_properties['dxf_list']) == 164

    # Verify first dxf style
    reference_file = os.path.join(DATADIR, 'writer', 'expected', 'dxf_style.xml')
    with open(reference_file) as expected:
        diff = compare_xml(read_xml, expected.read())
        assert diff is None, diff
    from openpyxl.style import Fill
    fill = Fill()
    fill.end_color = "FFFFC7CE"
    expected_style = {'font': {'color': 'FF9C0006', 'bold': False, 'italic': False}, 'border': [], 'fill':fill}
    #assert wb.style_properties['dxf_list'][0] == expected_style


    # Verify that the dxf styles stay the same when they're written and read back in.
    w = StyleWriter(wb)
    w._write_dxfs()
    write_xml = get_xml(w._root)
    read_style_prop = read_style_table(write_xml)
    assert len(read_style_prop['dxf_list']) == len(wb.style_properties['dxf_list'])
    for i, dxf in enumerate(read_style_prop['dxf_list']):
        assert repr(wb.style_properties['dxf_list'][i] == dxf)


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