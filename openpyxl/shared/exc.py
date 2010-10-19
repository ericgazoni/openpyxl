# file openpyxl/shared/exc.py

# Copyright (c) 2010 openpyxl
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
# @author: Eric Gazoni

"""Definitions for openpyxl shared exception classes."""


class CellCoordinatesException(Exception):
    """Error for converting between numeric and A1-style cell references."""
    pass


class ColumnStringIndexException(Exception):
    """Error for bad column names in A1-style cell references."""
    pass


class DataTypeException(Exception):
    """Error for any data type inconsistencies."""
    pass


class NamedRangeException(Exception):
    """Error for badly formatted named ranges."""
    pass


class SheetTitleException(Exception):
    """Error for bad sheet names."""
    pass


class InsufficientCoordinatesException(Exception):
    """Error for partially specified cell coordinates."""
    pass
