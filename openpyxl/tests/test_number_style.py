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


def test_format_comparisions(NumberFormat):
    format1 = NumberFormat()
    format2 = NumberFormat()
    format3 = NumberFormat()
    format1.format_code = 'm/d/yyyy'
    format2.format_code = 'm/d/yyyy'
    format3.format_code = 'mm/dd/yyyy'
    assert format1 == format2
    assert format1 == 'm/d/yyyy' and format1 != 'mm/dd/yyyy'
    assert format3 != 'm/d/yyyy' and format3 == 'mm/dd/yyyy'
    assert format1 != format3


def test_builtin_format(NumberFormat):
    fmt = NumberFormat()
    fmt.format_code = '0.00'
    assert fmt.builtin_format_code(2) == fmt.format_code
