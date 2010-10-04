# file openpyxl/tests/test_password_hash.py

# 3rd party imports
from nose.tools import eq_

# package imports
from openpyxl.shared.password_hasher import hash_password
from openpyxl.worksheet import SheetProtection


def test_hasher():
    eq_('CBEB', hash_password('test'))


def test_sheet_protection():
    protection = SheetProtection()
    protection.password = 'test'
    eq_('CBEB', protection.password)
