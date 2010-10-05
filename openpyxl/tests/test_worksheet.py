# file openpyxl/tests/test_worksheet.py

# 3rd party imports
from nose.tools import eq_, raises, assert_raises

# package imports
from openpyxl.workbook import Workbook
from openpyxl.worksheet import Worksheet, Relationship
from openpyxl.cell import Cell
from openpyxl.shared.exc import CellCoordinatesException, \
        SheetTitleException, InsufficientCoordinatesException


class TestWorksheet():

    @classmethod
    def setup_class(cls):
        cls.wb = Workbook()

    def test_new_worksheet(self):
        ws = Worksheet(self.wb)
        eq_(self.wb, ws._parent)

    def test_get_cell(self):
        ws = Worksheet(self.wb)
        cell = ws.cell('A1')
        eq_(cell.get_coordinate(), 'A1')

    @raises(SheetTitleException)
    def test_set_bad_title(self):
        Worksheet(self.wb, 'X' * 50)

    def test_set_bad_title_character(self):
        assert_raises(SheetTitleException, Worksheet, self.wb, '[')
        assert_raises(SheetTitleException, Worksheet, self.wb, ']')
        assert_raises(SheetTitleException, Worksheet, self.wb, '*')
        assert_raises(SheetTitleException, Worksheet, self.wb, ':')
        assert_raises(SheetTitleException, Worksheet, self.wb, '?')
        assert_raises(SheetTitleException, Worksheet, self.wb, '/')
        assert_raises(SheetTitleException, Worksheet, self.wb, '\\')

    def test_worksheet_dimension(self):
        ws = Worksheet(self.wb)
        eq_('A1:A1', ws.calculate_dimension())
        ws.cell('B12').value = 'AAA'
        eq_('A1:B12', ws.calculate_dimension())

    def test_worksheet_range(self):
        ws = Worksheet(self.wb)
        xlrange = ws.range('A1:C4')
        assert isinstance(xlrange, tuple)
        eq_(4, len(xlrange))
        eq_(3, len(xlrange[0]))

    def test_worksheet_named_range(self):
        ws = Worksheet(self.wb)
        self.wb.create_named_range('test_range', ws, 'C5')
        xlrange = ws.range("test_range")
        assert isinstance(xlrange, Cell)
        eq_(5, xlrange.row)

    def test_cell_offset(self):
        ws = Worksheet(self.wb)
        eq_('C17', ws.cell('B15').offset(2, 1).get_coordinate())

    def test_range_offset(self):
        ws = Worksheet(self.wb)
        xlrange = ws.range('A1:C4', 1, 3)
        assert isinstance(xlrange, tuple)
        eq_(4, len(xlrange))
        eq_(3, len(xlrange[0]))
        eq_('D2', xlrange[0][0].get_coordinate())

    def test_cell_alternate_coordinates(self):
        ws = Worksheet(self.wb)
        cell = ws.cell(row=8, column=4)
        eq_('D8', cell.get_coordinate())

    @raises(InsufficientCoordinatesException)
    def test_cell_insufficient_coordinates(self):
        ws = Worksheet(self.wb)
        cell = ws.cell(row=8)

    def test_cell_range_name(self):
        ws = Worksheet(self.wb)
        self.wb.create_named_range('test_range_single', ws, 'B12')
        assert_raises(CellCoordinatesException, ws.cell, 'test_range_single')
        c_range_name = ws.range('test_range_single')
        c_range_coord = ws.range('B12')
        c_cell = ws.cell('B12')
        eq_(c_range_coord, c_range_name)
        eq_(c_range_coord, c_cell)

    def test_garbage_collect(self):
        ws = Worksheet(self.wb)
        ws.cell('A1').value = ''
        ws.cell('B2').value = '0'
        ws.cell('C4').value = 0
        ws.garbage_collect()
        eq_(ws.get_cell_collection(), [ws.cell('B2'), ws.cell('C4')])

    def test_hyperlink_relationships(self):
        ws = Worksheet(self.wb)
        eq_(len(ws.relationships), 0)

        ws.cell('A1').hyperlink = "http://test.com"
        eq_(len(ws.relationships), 1)
        eq_("rId1", ws.cell('A1').hyperlink_rel_id)
        eq_("rId1", ws.relationships[0].id)
        eq_("http://test.com", ws.relationships[0].target)
        eq_("External", ws.relationships[0].target_mode)

        ws.cell('A2').hyperlink = "http://test2.com"
        eq_(len(ws.relationships), 2)
        eq_("rId2", ws.cell('A2').hyperlink_rel_id)
        eq_("rId2", ws.relationships[1].id)
        eq_("http://test2.com", ws.relationships[1].target)
        eq_("External", ws.relationships[1].target_mode)

    @raises(ValueError)
    def test_bad_relationship_type(self):
        rel = Relationship('bad_type')
