from __future__ import absolute_import
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


from .password_hasher import hash_password

class SheetProtection(object):
    """Information about protection of various aspects of a sheet."""

    def __init__(self):
        self.sheet = False
        self.objects = False
        self.scenarios = False
        self.format_cells = False
        self.format_columns = False
        self.format_rows = False
        self.insert_columns = False
        self.insert_rows = False
        self.insert_hyperlinks = False
        self.delete_columns = False
        self.delete_rows = False
        self.select_locked_cells = False
        self.sort = False
        self.auto_filter = False
        self.pivot_tables = False
        self.select_unlocked_cells = False
        self._password = ''
        self.enabled = False

    def set_password(self, value='', already_hashed=False):
        """Set a password on this sheet."""
        if not already_hashed:
            value = hash_password(value)
        self._password = value
        self.enabled = True

    def _set_raw_password(self, value):
        """Set a password directly, forcing a hash step."""
        self.set_password(value, already_hashed=False)

    def _get_raw_password(self):
        """Return the password value, regardless of hash."""
        return self._password

    def enable(self):
        self.enabled = True

    def disable(self):
        self.enabled = False

    password = property(_get_raw_password, _set_raw_password,
            'get/set the password (if already hashed, '
            'use set_password() instead)')

