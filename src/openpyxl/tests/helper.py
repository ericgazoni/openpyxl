# file openpyxl/tests/helper.py

# Python stdlib imports
from __future__ import with_statement
import os
import os.path
import shutil
import unittest
import difflib
from StringIO import StringIO
from pprint import pprint
from tempfile import gettempdir

# package imports
from openpyxl.shared.xmltools import fromstring, ElementTree
from openpyxl.shared.xmltools import pretty_indent

# constants
DATADIR = os.path.abspath(os.path.join(os.path.dirname(__file__), 'test_data'))
TMPDIR = os.path.join(gettempdir(), 'openpyxl_test_temp')


def clean_tmpdir():
    if os.path.isdir(TMPDIR):
        shutil.rmtree(TMPDIR, ignore_errors=True)
    os.makedirs(TMPDIR)

def assert_equals_file_content(reference_file, fixture, filetype='xml'):
    if os.path.isfile(fixture):
        with open(fixture) as fixture_file:
            fixture_content = fixture_file.read()
    else:
        fixture_content = fixture

    with open(reference_file) as expected_file:
        expected_content = expected_file.read()

    if filetype == 'xml':
        print fixture_content
        fixture_content = fromstring(fixture_content)
        pretty_indent(fixture_content)
        temp = StringIO()
        ElementTree(fixture_content).write(temp)
        fixture_content = temp.getvalue()

        expected_content = fromstring(expected_content)
        pretty_indent(expected_content)
        temp = StringIO()
        ElementTree(expected_content).write(temp)
        expected_content = temp.getvalue()

    fixture_lines = fixture_content.split('\n')
    expected_lines = expected_content.split('\n')
    differences = list(difflib.unified_diff(expected_lines, fixture_lines))
    if differences:
        temp = StringIO()
        pprint(differences, stream=temp)
        assert False, 'Differences found : %s' % temp.getvalue()


class BaseTestCase(unittest.TestCase):

    def assertEqualsFileContent(self, reference_file, fixture, filetype='xml'):
        if os.path.isfile(fixture):
            with open(fixture) as fixture_file:
                fixture_content = fixture_file.read()
        else:
            fixture_content = fixture

        with open(reference_file) as expected_file:
            expected_content = expected_file.read()

        if filetype == 'xml':
            print fixture_content
            fixture_content = fromstring(fixture_content)
            pretty_indent(fixture_content)
            temp = StringIO()
            ElementTree(fixture_content).write(temp)
            fixture_content = temp.getvalue()

            expected_content = fromstring(expected_content)
            pretty_indent(expected_content)
            temp = StringIO()
            ElementTree(expected_content).write(temp)
            expected_content = temp.getvalue()

        fixture_lines = fixture_content.split('\n')
        expected_lines = expected_content.split('\n')
        differences = list(difflib.unified_diff(expected_lines, fixture_lines))
        if differences:
            temp = StringIO()
            pprint(differences, stream=temp)
            self.fail('Differences found : %s' % temp.getvalue())

    def tearDown(self):
        self.clean_tmpdir()

    def clean_tmpdir(self):
        clean_tmpdir()
