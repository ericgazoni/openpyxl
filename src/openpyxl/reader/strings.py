# coding=UTF-8
'''
Copyright (c) 2010 openpyxl

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.

@license: http://www.opensource.org/licenses/mit-license.php
@author: Eric Gazoni
'''

from openpyxl.shared.xmltools import fromstring, QName
from openpyxl.shared.ooxml import NAMESPACES

def read_string_table(xml_source):

    table = {}

    xmlns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

    root = fromstring(text = xml_source)

    si_nodes = root.findall(QName(xmlns, 'si').text)

    for i, si in enumerate(si_nodes):

        table[i] = get_string(xmlns, si)

    return table

def get_string(xmlns, si):

    rich_nodes = si.findall(QName(xmlns, 'r').text)

    if rich_nodes:

        res = ''

        for r in rich_nodes:

            cur = get_text(xmlns, r)

            res += cur

        return res

    else:

        return get_text(xmlns, si)


def get_text(xmlns, r):

    t = r.find(QName(xmlns, 't').text)

    cur = t.text

    if t.get(QName(NAMESPACES['xml'], 'space').text) != 'preserve':

        cur = cur.strip()

    return cur
