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

# 3rd party imports
#from nose.tools import eq_, raises, assert_raises
import pytest
from .helper import compare_xml


# package imports
from openpyxl.workbook import Workbook
from openpyxl.worksheet import Worksheet, Relationship, flatten
from openpyxl.cell import Cell, coordinate_from_string
from openpyxl.comments import Comment
from openpyxl.exceptions import (
    CellCoordinatesException,
    SheetTitleException,
    InsufficientCoordinatesException,
    NamedRangeException
    )
from openpyxl.writer.worksheet import write_worksheet

class TestWorksheet(object):

    @classmethod
    def setup_class(cls):
        cls.wb = Workbook()

    def test_new_worksheet(self):
        ws = Worksheet(self.wb)
        assert self.wb == ws._parent

    def test_new_sheet_name(self):
        self.wb.worksheets = []
        ws = Worksheet(self.wb, title='')
        assert repr(ws) == '<Worksheet "Sheet1">'

    def test_get_cell(self):
        ws = Worksheet(self.wb)
        cell = ws.cell('A1')
        assert cell.coordinate == 'A1'

    def test_set_bad_title(self):
        with pytest.raises(SheetTitleException):
            Worksheet(self.wb, 'X' * 50)

    def test_increment_title(self):
        ws1 = self.wb.create_sheet(title="Test")
        assert ws1.title == "Test"
        ws2 = self.wb.create_sheet(title="Test")
        assert ws2.title == "Test1"

    @pytest.mark.parametrize("value", ["[", "]", "*", ":", "?", "/", "\\"])
    def test_set_bad_title_character(self, value):
        with pytest.raises(SheetTitleException):
            Worksheet(self.wb, value)


    def test_unique_sheet_title(self):
        ws = self.wb.create_sheet(title="AGE")
        assert ws.unique_sheet_name("GE") == "GE"


    def test_worksheet_dimension(self):
        ws = Worksheet(self.wb)
        assert 'A1:A1' == ws.calculate_dimension()
        ws.cell('B12').value = 'AAA'
        assert 'A1:B12' == ws.calculate_dimension()

    def test_worksheet_range(self):
        ws = Worksheet(self.wb)
        xlrange = ws.range('A1:C4')
        assert isinstance(xlrange, tuple)
        assert 4 == len(xlrange)
        assert 3 == len(xlrange[0])

    def test_worksheet_named_range(self):
        ws = Worksheet(self.wb)
        self.wb.create_named_range('test_range', ws, 'C5')
        xlrange = ws.range('test_range')
        assert isinstance(xlrange, Cell)
        assert 5 == xlrange.row

    def test_bad_named_range(self):
        ws = Worksheet(self.wb)
        with pytest.raises(NamedRangeException):
            ws.range('bad_range')

    def test_named_range_wrong_sheet(self):
        ws1 = Worksheet(self.wb)
        ws2 = Worksheet(self.wb)
        self.wb.create_named_range('wrong_sheet_range', ws1, 'C5')
        with pytest.raises(NamedRangeException):
            ws2.range('wrong_sheet_range')

    def test_range_offset(self):
        ws = Worksheet(self.wb)
        xlrange = ws.range('A1:C4', 1, 3)
        assert isinstance(xlrange, tuple)
        assert 4 == len(xlrange)
        assert 3 == len(xlrange[0])
        assert 'D2' == xlrange[0][0].coordinate

    def test_cell_alternate_coordinates(self):
        ws = Worksheet(self.wb)
        cell = ws.cell(row=8, column=4)
        assert 'E9' == cell.coordinate

    def test_cell_insufficient_coordinates(self):
        ws = Worksheet(self.wb)
        with pytest.raises(InsufficientCoordinatesException):
            ws.cell(row=8)

    def test_cell_range_name(self):
        ws = Worksheet(self.wb)
        self.wb.create_named_range('test_range_single', ws, 'B12')
        with pytest.raises(CellCoordinatesException):
            ws.cell('test_range_single')
        c_range_name = ws.range('test_range_single')
        c_range_coord = ws.range('B12')
        c_cell = ws.cell('B12')
        assert c_range_coord == c_range_name
        assert c_range_coord == c_cell

    def test_garbage_collect(self):
        ws = Worksheet(self.wb)
        ws.cell('A1').value = ''
        ws.cell('B2').value = '0'
        ws.cell('C4').value = 0
        ws.cell('D1').comment = Comment('Comment', 'Comment')
        ws.garbage_collect()
        assert set(ws.get_cell_collection()), set([ws.cell('B2'), ws.cell('C4') == ws.cell('D1')])

    def test_hyperlink_relationships(self):
        ws = Worksheet(self.wb)
        assert len(ws.relationships) == 0

        ws.cell('A1').hyperlink = "http://test.com"
        assert len(ws.relationships) == 1
        assert "rId1" == ws.cell('A1').hyperlink_rel_id
        assert "rId1" == ws.relationships[0].id
        assert "http://test.com" == ws.relationships[0].target
        assert "External" == ws.relationships[0].target_mode

        ws.cell('A2').hyperlink = "http://test2.com"
        assert len(ws.relationships) == 2
        assert "rId2" == ws.cell('A2').hyperlink_rel_id
        assert "rId2" == ws.relationships[1].id
        assert "http://test2.com" == ws.relationships[1].target
        assert "External" == ws.relationships[1].target_mode

    def test_bad_relationship_type(self):
        with pytest.raises(ValueError):
            Relationship('bad_type')

    def test_append_list(self):
        ws = Worksheet(self.wb)

        ws.append(['This is A1', 'This is B1'])

        assert 'This is A1' == ws.cell('A1').value
        assert 'This is B1' == ws.cell('B1').value

    def test_append_dict_letter(self):
        ws = Worksheet(self.wb)

        ws.append({'A' : 'This is A1', 'C' : 'This is C1'})

        assert 'This is A1' == ws.cell('A1').value
        assert 'This is C1' == ws.cell('C1').value

    def test_append_dict_index(self):
        ws = Worksheet(self.wb)

        ws.append({0 : 'This is A1', 2 : 'This is C1'})

        assert 'This is A1' == ws.cell('A1').value
        assert 'This is C1' == ws.cell('C1').value

    def test_bad_append(self):
        ws = Worksheet(self.wb)
        with pytest.raises(TypeError):
            ws.append("test")

    def test_append_2d_list(self):

        ws = Worksheet(self.wb)

        ws.append(['This is A1', 'This is B1'])
        ws.append(['This is A2', 'This is B2'])

        vals = ws.range('A1:B2')
        expected = (
            ('This is A1', 'This is B1'),
            ('This is A2', 'This is B2'),
        )
        for e, v in zip(expected, flatten(vals)):
            assert e == tuple(v)

    def test_rows(self):

        ws = Worksheet(self.wb)

        ws.cell('A1').value = 'first'
        ws.cell('C9').value = 'last'

        rows = ws.rows

        assert len(rows) == 9

        assert rows[0][0].value == 'first'
        assert rows[-1][-1].value == 'last'

    def test_cols(self):

        ws = Worksheet(self.wb)

        ws.cell('A1').value = 'first'
        ws.cell('C9').value = 'last'

        cols = ws.columns

        assert len(cols) == 3

        assert cols[0][0].value == 'first'
        assert cols[-1][-1].value == 'last'

    def test_auto_filter(self):
        ws = Worksheet(self.wb)
        ws.auto_filter.ref = ws.range('a1:f1')
        assert ws.auto_filter.ref == 'A1:F1'

        ws.auto_filter.ref = ''
        assert ws.auto_filter.ref is None

        ws.auto_filter.ref = 'c1:g9'
        assert ws.auto_filter.ref == 'C1:G9'

    def test_getitem(self):
        ws = Worksheet(self.wb)
        c = ws['A1']
        assert isinstance(c, Cell)
        assert c.coordinate == "A1"
        assert ws['A1'].value is None

    def test_setitem(self):
        ws = Worksheet(self.wb)
        ws['A12'] = 5
        assert ws['A12'].value == 5

    def test_getslice(self):
        ws = Worksheet(self.wb)
        cell_range = ws['A1':'B2']
        assert isinstance(cell_range, tuple)
        assert (cell_range) == ((ws['A1'], ws['B1']), (ws['A2'], ws['B2']))


    def test_freeze(self):
        ws = Worksheet(self.wb)
        ws.freeze_panes = ws.cell('b2')
        assert ws.freeze_panes == 'B2'

        ws.freeze_panes = ''
        assert ws.freeze_panes is None

        ws.freeze_panes = 'c5'
        assert ws.freeze_panes == 'C5'

        ws.freeze_panes = ws.cell('A1')
        assert ws.freeze_panes is None


class TestWorkSheetWriter(object):

    @classmethod
    def setup_class(cls):
        cls.wb = Workbook()

    def test_write_empty(self):
        ws = Worksheet(self.wb)
        xml = write_worksheet(ws, None, None)
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheetPr>
            <outlinePr summaryRight="1" summaryBelow="1"/>
          </sheetPr>
          <dimension ref="A1:A1"/>
          <sheetViews>
            <sheetView workbookViewId="0">
              <selection sqref="A1" activeCell="A1"/>
            </sheetView>
          </sheetViews>
          <sheetFormatPr defaultRowHeight="15"/>
          <sheetData/>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_page_margins(self):
        ws = Worksheet(self.wb)
        ws.page_margins.left = 2.0
        ws.page_margins.right = 2.0
        ws.page_margins.top = 2.0
        ws.page_margins.bottom = 2.0
        ws.page_margins.header = 1.5
        ws.page_margins.footer = 1.5
        xml = write_worksheet(ws, None, None)
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheetPr>
            <outlinePr summaryRight="1" summaryBelow="1"/>
          </sheetPr>
          <dimension ref="A1:A1"/>
          <sheetViews>
            <sheetView workbookViewId="0">
              <selection sqref="A1" activeCell="A1"/>
            </sheetView>
          </sheetViews>
          <sheetFormatPr defaultRowHeight="15"/>
          <sheetData/>
          <pageMargins left="2.00" right="2.00" top="2.00" bottom="2.00" header="1.50" footer="1.50"/>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_merge(self):
        ws = Worksheet(self.wb)
        string_table = {'':'', 'Cell A1':'Cell A1', 'Cell B1':'Cell B1'}
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheetPr>
            <outlinePr summaryRight="1" summaryBelow="1"/>
          </sheetPr>
          <dimension ref="A1:B1"/>
          <sheetViews>
            <sheetView workbookViewId="0">
              <selection sqref="A1" activeCell="A1"/>
            </sheetView>
          </sheetViews>
          <sheetFormatPr defaultRowHeight="15"/>
          <cols>
            <col min="1" max="1" width="9.10"/>
            <col min="2" max="2" width="9.10"/>
          </cols>
          <sheetData>
            <row r="1" spans="1:2">
              <c r="A1" t="s">
                <v>Cell A1</v>
              </c>
              <c r="B1" t="s">
                <v>Cell B1</v>
              </c>
            </row>
          </sheetData>
        </worksheet>
        """

        ws.cell('A1').value = 'Cell A1'
        ws.cell('B1').value = 'Cell B1'
        xml = write_worksheet(ws, string_table, None)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

        ws.merge_cells('A1:B1')
        xml = write_worksheet(ws, string_table, None)
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheetPr>
            <outlinePr summaryRight="1" summaryBelow="1"/>
          </sheetPr>
          <dimension ref="A1:B1"/>
          <sheetViews>
            <sheetView workbookViewId="0">
              <selection sqref="A1" activeCell="A1"/>
            </sheetView>
          </sheetViews>
          <sheetFormatPr defaultRowHeight="15"/>
          <cols>
            <col min="1" max="1" width="9.10"/>
            <col min="2" max="2" width="9.10"/>
          </cols>
          <sheetData>
            <row r="1" spans="1:2">
              <c r="A1" t="s">
                <v>Cell A1</v>
              </c>
              <c r="B1" t="s"/>
            </row>
          </sheetData>
          <mergeCells count="1">
            <mergeCell ref="A1:B1"/>
          </mergeCells>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

        ws.unmerge_cells('A1:B1')
        xml = write_worksheet(ws, string_table, None)
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheetPr>
            <outlinePr summaryRight="1" summaryBelow="1"/>
          </sheetPr>
          <dimension ref="A1:B1"/>
          <sheetViews>
            <sheetView workbookViewId="0">
              <selection sqref="A1" activeCell="A1"/>
            </sheetView>
          </sheetViews>
          <sheetFormatPr defaultRowHeight="15"/>
          <cols>
            <col min="1" max="1" width="9.10"/>
            <col min="2" max="2" width="9.10"/>
          </cols>
          <sheetData>
            <row r="1" spans="1:2">
              <c r="A1" t="s">
                <v>Cell A1</v>
              </c>
              <c r="B1" t="s"/>
            </row>
          </sheetData>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_printer_settings(self):
        ws = Worksheet(self.wb)
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_TABLOID
        ws.page_setup.fitToPage = True
        ws.page_setup.fitToHeight = 0
        ws.page_setup.fitToWidth = 1
        ws.page_setup.horizontalCentered = True
        ws.page_setup.verticalCentered = True
        xml = write_worksheet(ws, None, None)
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheetPr>
            <outlinePr summaryRight="1" summaryBelow="1"/>
            <pageSetUpPr fitToPage="1"/>
          </sheetPr>
          <dimension ref="A1:A1"/>
          <sheetViews>
            <sheetView workbookViewId="0">
              <selection sqref="A1" activeCell="A1"/>
            </sheetView>
          </sheetViews>
          <sheetFormatPr defaultRowHeight="15"/>
          <sheetData/>
          <printOptions horizontalCentered="1" verticalCentered="1"/>
          <pageSetup orientation="landscape" paperSize="3" fitToHeight="0" fitToWidth="1"/>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_header_footer(self):
        ws = Worksheet(self.wb)
        ws.header_footer.left_header.text = "Left Header Text"
        ws.header_footer.center_header.text = "Center Header Text"
        ws.header_footer.center_header.font_name = "Arial,Regular"
        ws.header_footer.center_header.font_size = 6
        ws.header_footer.center_header.font_color = "445566"
        ws.header_footer.right_header.text = "Right Header Text"
        ws.header_footer.right_header.font_name = "Arial,Bold"
        ws.header_footer.right_header.font_size = 8
        ws.header_footer.right_header.font_color = "112233"
        ws.header_footer.left_footer.text = "Left Footer Text\nAnd &[Date] and &[Time]"
        ws.header_footer.left_footer.font_name = "Times New Roman,Regular"
        ws.header_footer.left_footer.font_size = 10
        ws.header_footer.left_footer.font_color = "445566"
        ws.header_footer.center_footer.text = "Center Footer Text &[Path]&[File] on &[Tab]"
        ws.header_footer.center_footer.font_name = "Times New Roman,Bold"
        ws.header_footer.center_footer.font_size = 12
        ws.header_footer.center_footer.font_color = "778899"
        ws.header_footer.right_footer.text = "Right Footer Text &[Page] of &[Pages]"
        ws.header_footer.right_footer.font_name = "Times New Roman,Italic"
        ws.header_footer.right_footer.font_size = 14
        ws.header_footer.right_footer.font_color = "AABBCC"
        xml_string = write_worksheet(ws, None, None)
        diff = compare_xml(xml_string, """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheetPr>
            <outlinePr summaryRight="1" summaryBelow="1"/>
          </sheetPr>
          <dimension ref="A1:A1"/>
          <sheetViews>
            <sheetView workbookViewId="0">
              <selection sqref="A1" activeCell="A1"/>
            </sheetView>
          </sheetViews>
          <sheetFormatPr defaultRowHeight="15"/>
          <sheetData/>
          <headerFooter>
            <oddHeader>&amp;L&amp;"Calibri,Regular"&amp;K000000Left Header Text&amp;C&amp;"Arial,Regular"&amp;6&amp;K445566Center Header Text&amp;R&amp;"Arial,Bold"&amp;8&amp;K112233Right Header Text</oddHeader>
            <oddFooter>&amp;L&amp;"Times New Roman,Regular"&amp;10&amp;K445566Left Footer Text_x000D_And &amp;D and &amp;T&amp;C&amp;"Times New Roman,Bold"&amp;12&amp;K778899Center Footer Text &amp;Z&amp;F on &amp;A&amp;R&amp;"Times New Roman,Italic"&amp;14&amp;KAABBCCRight Footer Text &amp;P of &amp;N</oddFooter>
          </headerFooter>
        </worksheet>
        """)
        assert diff is None, diff

        ws = Worksheet(self.wb)
        xml_string = write_worksheet(ws, None, None)
        diff = compare_xml(xml_string, """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheetPr>
            <outlinePr summaryRight="1" summaryBelow="1"/>
          </sheetPr>
          <dimension ref="A1:A1"/>
          <sheetViews>
            <sheetView workbookViewId="0">
              <selection sqref="A1" activeCell="A1"/>
            </sheetView>
          </sheetViews>
          <sheetFormatPr defaultRowHeight="15"/>
          <sheetData/>
        </worksheet>
        """)
        assert diff is None, diff


class TestPositioning(object):
    def test_point(self):
        wb = Workbook()
        ws = wb.get_active_sheet()
        assert ws.point_pos(top=40, left=150), ('C' == 3)

    @pytest.mark.parametrize("value", ('A1', 'D52', 'X11'))
    def test_roundtrip(self, value):
        wb = Workbook()
        ws = wb.get_active_sheet()
        assert ws.point_pos(*ws.cell(value).anchor) == coordinate_from_string(value)


@pytest.fixture
def PageSetup():
    from openpyxl.worksheet import PageSetup
    return PageSetup


@pytest.mark.xfail
def test_page_setup(PageSetup):
    p = PageSetup()
    assert p.setup == {}
    p.scale = 1
    assert p.setup['scale'] == 1


def test_page_options(PageSetup):
    p = PageSetup()
    assert p.options == {}
    p.horizontalCentered = True
    p.verticalCentered = True
    assert p.options == {'verticalCentered': '1', 'horizontalCentered': '1'}
