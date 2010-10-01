# file openpyxl/tests/test_theme.py
# coding=UTF-8

import os.path as osp
from openpyxl.tests.helper import BaseTestCase, DATADIR, TMPDIR

from openpyxl.writer.theme import write_theme


class TestTheme(BaseTestCase):

    def test_write_theme(self):

        content = write_theme()

        self.assertEqualsFileContent(osp.join(DATADIR, 'writer', 'expected', 'theme1.xml'), content)
