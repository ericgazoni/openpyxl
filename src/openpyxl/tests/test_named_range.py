# file openpyxl/tests/test_named_range.py

# Python stdlib imports
from __future__ import with_statement
import os.path

# 3rd-party imports
from nose.tools import eq_, assert_raises

# package imports
from openpyxl.tests.helper import DATADIR
from openpyxl.namedrange import split_named_range
from openpyxl.reader.workbook import read_named_ranges
from openpyxl.shared.exc import NamedRangeException


def test_split():
    eq_(('My Sheet', 'D', 8), split_named_range("'My Sheet'!$D$8"))


def test_split_no_quotes():
    eq_(('HYPOTHESES', 'B', 3), split_named_range('HYPOTHESES!$B$3:$L$3'))


def test_bad_range_name():
    assert_raises(NamedRangeException, split_named_range, 'HYPOTHESES$B$3')


def test_read_named_ranges():

    class DummyWs(object):
        title = 'My Sheeet'

    class DummyWB(object):

        def get_sheet_by_name(self, name):
            return DummyWs()

    with open(os.path.join(DATADIR, 'reader', 'workbook.xml')) as handle:
        content = handle.read()
    named_ranges = read_named_ranges(content, DummyWB())
    eq_(["My Sheeet!D8"], [str(range) for range in named_ranges])
