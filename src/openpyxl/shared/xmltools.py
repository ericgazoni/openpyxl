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
from __future__ import with_statement
from openpyxl import __name__ as prefix
from os import close, remove
from tempfile import mkstemp
import atexit

try:
    from xml.etree.ElementTree import ElementTree, Element, SubElement, QName, fromstring #pylint: disable-msg=W0611
except ImportError:
    from cElementTree import ElementTree, Element, SubElement, QName, fromstring #pylint: disable-msg=F0401

XML_TEMP_FILES = []

@atexit.register
def cleanup_tempfiles():

    for handle, filename in XML_TEMP_FILES:
        try:
            close(handle)
            remove(filename)
        except:
            pass

def get_tempfile():

    fd, filename = mkstemp(prefix = prefix, text = True)
    XML_TEMP_FILES.append((fd, filename))

    return filename

def get_document_content(xml_node):

    pretty_indent(xml_node)

    filename = get_tempfile()

    with open(filename, 'w') as fl:

        ElementTree(xml_node).write(file = fl, encoding = 'UTF-8')

    return filename

def pretty_indent(elem, level = 0):
    i = "\n" + level * "  "
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "  "
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
        for elem in elem:
            pretty_indent(elem, level + 1)
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i
