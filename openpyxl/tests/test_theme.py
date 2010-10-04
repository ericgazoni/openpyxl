# file openpyxl/tests/test_theme.py

# Python stdlib imports
import os.path

# package imports
from openpyxl.tests.helper import DATADIR, assert_equals_file_content
from openpyxl.writer.theme import write_theme


def test_write_theme():
    content = write_theme()
    assert_equals_file_content(
            os.path.join(DATADIR, 'writer', 'expected', 'theme1.xml'), content)
