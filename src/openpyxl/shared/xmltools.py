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
from openpyxl import __name__ as prefix
from os import close, remove
from tempfile import mkstemp
from xml.sax.saxutils import XMLGenerator
from xml.sax.xmlreader import AttributesNSImpl
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

    fl = open(filename, 'w')
    ElementTree(xml_node).write(file = fl, encoding = 'UTF-8')
    fl.close()

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

#===============================================================================
# Shortcut functions taken from 
# http://lethain.com/entry/2009/jan/22/handling-very-large-csv-and-xml-files-in-python/
#===============================================================================

def start_tag(doc, name, attr = {}, body = None, namespace = None):
    attr_vals = {}
    attr_keys = {}
    for key, val in attr.iteritems():
        key_tuple = (namespace, key)
        attr_vals[key_tuple] = val
        attr_keys[key_tuple] = key

    attr2 = AttributesNSImpl(attr_vals, attr_keys)
    doc.startElementNS((namespace, name), name, attr2)
    if body:
        doc.characters(body)

def end_tag(doc, name, namespace = None):
    doc.endElementNS((namespace, name), name)

def tag(doc, name, attr = {}, body = None, namespace = None):
    start_tag(doc, name, attr, body, namespace)
    end_tag(doc, name, namespace)
