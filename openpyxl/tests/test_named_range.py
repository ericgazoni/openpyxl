# file openpyxl/tests/test_named_range.py

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

# 3rd-party imports
from nose.tools import eq_, assert_raises

# package imports
from openpyxl.tests.helper import DATADIR, TMPDIR, clean_tmpdir, make_tmpdir
from openpyxl.namedrange import split_named_range, NamedRange
from openpyxl.reader.workbook import read_named_ranges
from openpyxl.shared.exc import NamedRangeException
from openpyxl.reader.excel import load_workbook


def test_split():
    eq_([('My Sheet', '$D$8'), ], split_named_range("'My Sheet'!$D$8"))


def test_split_no_quotes():
    eq_([('HYPOTHESES', '$B$3:$L$3'), ], split_named_range('HYPOTHESES!$B$3:$L$3'))


def test_bad_range_name():
    assert_raises(NamedRangeException, split_named_range, 'HYPOTHESES$B$3')


def test_read_named_ranges():

    class DummyWs(object):
        title = 'My Sheeet'

        def __str__(self):
            return self.title

    class DummyWB(object):

        def get_sheet_by_name(self, name):
            return DummyWs()

    handle = open(os.path.join(DATADIR, 'reader', 'workbook.xml'))
    try:
        content = handle.read()
        named_ranges = read_named_ranges(content, DummyWB())
        eq_(["My Sheeet!$D$8"], [str(range) for range in named_ranges])
    finally:
        handle.close()

def test_oddly_shaped_named_ranges():

    ranges_counts = ((4, 'TEST_RANGE'),
                     (3, 'TRAP_1'),
                     (13, 'TRAP_2'))

    def check_ranges(ws, count, range_name):

        eq_(count, len(ws.range(range_name)))

    wb = load_workbook(os.path.join(DATADIR, 'genuine', 'merge_range.xlsx'),
                       use_iterators = False)

    ws = wb.worksheets[0]

    for count, range_name in ranges_counts:

        yield check_ranges, ws, count, range_name


def test_merged_cells_named_range():

    wb = load_workbook(os.path.join(DATADIR, 'genuine', 'merge_range.xlsx'),
                       use_iterators = False)

    ws = wb.worksheets[0]

    cell = ws.range('TRAP_3')

    eq_('B15', cell.get_coordinate())

    eq_(10, cell.value)



class TestNameRefersToValue(object):
    def setUp(self):
        self.wb = load_workbook(os.path.join(DATADIR, 'genuine', 'NameWithValueBug.xlsx'))
        self.ws = self.wb.get_sheet_by_name("Sheet1")
        make_tmpdir()

    def tearDown(self):
        clean_tmpdir()

    def test_has_ranges(self):
        ranges = self.wb.get_named_ranges()
        eq_(['MyRef', 'MySheetRef', 'MySheetRef', 'MySheetValue', 'MySheetValue', 'MyValue'], [range.name for range in ranges])

    def test_workbook_has_normal_range(self):
        normal_range = self.wb.get_named_range("MyRef")
        eq_("MyRef", normal_range.name)

    def test_workbook_has_value_range(self):
        value_range = self.wb.get_named_range("MyValue")
        eq_("MyValue", value_range.name)
        eq_("9.99", value_range.value)

    def test_worksheet_range(self):
        range = self.ws.range("MyRef")

    def test_worksheet_range_error_on_value_range(self):
        assert_raises(NamedRangeException, self.ws.range, "MyValue")

    def range_as_string(self, range, include_value=False):
        def scope_as_string(range):
            if range.scope:
                return range.scope.title
            else:
                return "Workbook"
        retval = "%s: %s" % (range.name, scope_as_string(range))
        if include_value:
            if isinstance(range, NamedRange):
                retval += "=[range]"
            else:
                retval += "=" + range.value
        return retval

    def test_handles_scope(self):
        ranges = self.wb.get_named_ranges()
        eq_(['MyRef: Workbook', 'MySheetRef: Sheet1', 'MySheetRef: Sheet2', 'MySheetValue: Sheet1', 'MySheetValue: Sheet2', 'MyValue: Workbook'], 
            [self.range_as_string(range) for range in ranges])

    def test_can_be_saved(self):
        FNAME = os.path.join(TMPDIR, "foo.xlsx")
        self.wb.save(FNAME)

        wbcopy = load_workbook(FNAME)
        eq_(['MyRef: Workbook=[range]', 'MySheetRef: Sheet1=[range]', 'MySheetRef: Sheet2=[range]', 'MySheetValue: Sheet1=3.33', 'MySheetValue: Sheet2=14.4', 'MyValue: Workbook=9.99'], 
            [self.range_as_string(range, include_value=True) for range in wbcopy.get_named_ranges()])
