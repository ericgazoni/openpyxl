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

import pytest

# package imports
from openpyxl.tests.helper import DATADIR
from openpyxl.namedrange import split_named_range, NamedRange
from openpyxl.reader.workbook import read_named_ranges
from openpyxl.exceptions import NamedRangeException
from openpyxl.reader.excel import load_workbook
from openpyxl.workbook import Workbook


def test_split():
    assert [('My Sheet', '$D$8'), ] == split_named_range("'My Sheet'!$D$8")


def test_split_no_quotes():
    assert [('HYPOTHESES', '$B$3:$L$3'), ] == split_named_range('HYPOTHESES!$B$3:$L$3')


def test_bad_range_name():
    with pytest.raises(NamedRangeException):
        split_named_range('HYPOTHESES$B$3')

def test_range_name_worksheet_special_chars():
        class DummyWs(object):
            title = 'My Sheeet with a , and \''

            def __str__(self):
                return self.title
        ws = DummyWs()

        class DummyWB(object):

            def get_sheet_by_name(self, name):
                if name == ws.title:
                    return ws

        handle = open(os.path.join(DATADIR, 'reader', 'workbook_namedrange.xml'))
        try:
            content = handle.read()
            named_ranges = read_named_ranges(content, DummyWB())
            assert 1 == len(named_ranges)
            assert isinstance(named_ranges[0], NamedRange)
            assert [(ws, '$U$16:$U$24'), (ws, '$V$28:$V$36')] == named_ranges[0].destinations
        finally:
            handle.close()


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
        assert ["My Sheeet!$D$8"] == [str(range) for range in named_ranges]
    finally:
        handle.close()


ranges_counts = (
    (4, 'TEST_RANGE'),
    (3, 'TRAP_1'),
    (13, 'TRAP_2')
)
@pytest.mark.parametrize("count, range_name", ranges_counts)
def test_oddly_shaped_named_ranges(count, range_name):

    wb = load_workbook(os.path.join(DATADIR, 'genuine', 'merge_range.xlsx'),
                       use_iterators = False)
    ws = wb.worksheets[0]
    assert len(ws.range(range_name)) == count


def test_merged_cells_named_range():

    wb = load_workbook(os.path.join(DATADIR, 'genuine', 'merge_range.xlsx'),
                       use_iterators = False)
    ws = wb.worksheets[0]
    cell = ws.range('TRAP_3')
    assert 'B15' == cell.get_coordinate()
    assert 10 == cell.value


def test_print_titles():
    wb = Workbook()
    ws1 = wb.create_sheet()
    ws2 = wb.create_sheet()
    ws1.add_print_title(2)
    ws2.add_print_title(3, rows_or_cols='cols')

    def mystr(nr):
        return ','.join(['%s!%s' % (sheet.title, name) for sheet, name in nr.destinations])

    actual_named_ranges = set([(nr.name, nr.scope, mystr(nr)) for nr in wb.get_named_ranges()])
    expected_named_ranges = set([('_xlnm.Print_Titles', ws1, 'Sheet1!$1:$2'),
                                 ('_xlnm.Print_Titles', ws2, 'Sheet2!$A:$C')])
    assert(actual_named_ranges == expected_named_ranges)


class TestNameRefersToValue(object):
    def setup(self):
        self.wb = load_workbook(os.path.join(DATADIR, 'genuine', 'NameWithValueBug.xlsx'))
        self.ws = self.wb.get_sheet_by_name("Sheet1")

    def test_has_ranges(self):
        ranges = self.wb.get_named_ranges()
        assert ['MyRef', 'MySheetRef', 'MySheetRef', 'MySheetValue', 'MySheetValue',
                'MyValue'] == [range.name for range in ranges]

    def test_workbook_has_normal_range(self):
        normal_range = self.wb.get_named_range("MyRef")
        assert normal_range.name == "MyRef"
        assert normal_range.destinations == [(self.ws, '$A$1')]
        assert normal_range.scope is None

    def test_workbook_has_value_range(self):
        value_range = self.wb.get_named_range("MyValue")
        assert "MyValue" == value_range.name
        assert "9.99" == value_range.value

    def test_worksheet_range(self):
        range = self.ws.range("MyRef")
        assert range.get_coordinate() == "A1"

    def test_worksheet_range_error_on_value_range(self):
        with pytest.raises(NamedRangeException):
            self.ws.range("MyValue")

    # TODO document this helper function
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
        assert self.wb.get_sheet_names() == ["Sheet1", "Sheet2"]
        assert set(['MyRef: Workbook', 'MySheetRef: Sheet1', 'MySheetRef: Sheet2', 'MySheetValue: Sheet1',
                'MySheetValue: Sheet2', 'MyValue: Workbook']) == set([self.range_as_string(range) for range in ranges])

    def test_can_be_saved(self, tmpdir):
        tmpdir.chdir()
        FNAME = "foo.xlsx"
        self.wb.save(FNAME)

        wbcopy = load_workbook(FNAME)
        ranges = wbcopy.get_named_ranges()
        assert wbcopy.get_sheet_names() == ['Sheet1', 'Sheet2']
        assert ['MyRef: Workbook', 'MySheetRef: Sheet1', 'MySheetRef: Sheet2',
                'MySheetValue: Sheet1', 'MySheetValue: Sheet2', 'MyValue: Workbook'] \
               == [self.range_as_string(range) for range in ranges]

        assert [ 'MyRef: Workbook=[range]', 'MySheetRef: Sheet1=[range]',
                 'MySheetRef: Sheet2=[range]', 'MySheetValue: Sheet1=3.33',
                 'MySheetValue: Sheet2=14.4', 'MyValue: Workbook=9.99' ] \
               == [self.range_as_string(range, include_value=True) for range in
                   wbcopy.get_named_ranges()]
