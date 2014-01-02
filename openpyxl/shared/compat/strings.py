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
from __future__ import absolute_import
import sys

VER = sys.version_info

if VER[0] == 3:
    basestring = str
    unicode = str
    from io import BufferedReader
    file = BufferedReader
    from io import BufferedRandom
    tempfile = BufferedRandom
else:
    basestring = basestring
    unicode = unicode
    file = file
    tempfile = file

if VER[0] == 3:
    from io import BytesIO, StringIO
else:
    from StringIO import StringIO
    BytesIO = StringIO

def safe_string(value):
    from numbers import Number
    """Safely and consistently format numeric values"""
    if isinstance(value, Number):
        value = "%.15g" % value
    elif not isinstance(value, basestring):
        value = str(value)
    return value
