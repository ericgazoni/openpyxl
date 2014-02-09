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

"""Shared xml tools.

Shortcut functions taken from:
    http://lethain.com/entry/2009/jan/22/handling-very-large-csv-and-xml-files-in-python/

"""

# Python stdlib imports
from xml.sax.saxutils import XMLGenerator
from xml.sax.xmlreader import AttributesNSImpl


# compatibility
from openpyxl.compat import OrderedDict

# package imports
from openpyxl import LXML

if LXML is True:
    from lxml.etree import (
    Element,
    ElementTree,
    SubElement,
    fromstring,
    tostring,
    register_namespace,
    iterparse
    )
    from xml.etree.cElementTree import iterparse
else:
    try:
        from xml.etree.cElementTree import (
        ElementTree,
        Element,
        SubElement,
        fromstring,
        tostring,
        iterparse
        )
    except ImportError:
        from xml.etree.ElementTreee import (
        ElementTree,
        Element,
        SubElement,
        fromstring,
        tostring,
        iterparse
        )
    from .namespace import register_namespace

from openpyxl.xml.constants import (
    CHART_NS,
    DRAWING_NS,
    SHEET_DRAWING_NS,
    CHART_DRAWING_NS,
    SHEET_MAIN_NS,
    REL_NS,
    VTYPES_NS,
    COREPROPS_NS,
    DCTERMS_NS,
    DCTERMS_PREFIX
)

if LXML is False:
    _iterparse = iterparse
    def safe_iterparse(source, *args, **kw):
        return _iterparse(source)

    iterparse = safe_iterparse


register_namespace(DCTERMS_PREFIX, DCTERMS_NS)
register_namespace('dcmitype', 'http://purl.org/dc/dcmitype/')
register_namespace('cp', COREPROPS_NS)
register_namespace('c', CHART_NS)
register_namespace('a', DRAWING_NS)
register_namespace('s', SHEET_MAIN_NS)
register_namespace('r', REL_NS)
register_namespace('vt', VTYPES_NS)
register_namespace('xdr', SHEET_DRAWING_NS)
register_namespace('cdr', CHART_DRAWING_NS)

def get_document_content(xml_node):
    """Print nicely formatted xml to a string."""
    pretty_indent(xml_node)
    return tostring(xml_node, encoding='utf-8')


def pretty_indent(elem, level=0):
    """Format xml with nice indents and line breaks."""
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


def start_tag(doc, name, attr=None, body=None, namespace=None):
    """Wrapper to start an xml tag."""
    if attr is None:
        attr = {}
        dct_type = dict
    elif isinstance(attr, OrderedDict):
        dct_type = OrderedDict
    else:
        dct_type = dict

    attr_vals = dct_type()
    attr_keys = dct_type()
    for key, val in attr.items():
        key_tuple = (namespace, key)
        attr_vals[key_tuple] = val
        attr_keys[key_tuple] = key
    attr2 = AttributesNSImpl(attr_vals, attr_keys)
    doc.startElementNS((namespace, name), name, attr2)
    if body:
        doc.characters(body)


def end_tag(doc, name, namespace=None):
    """Wrapper to close an xml tag."""
    doc.endElementNS((namespace, name), name)


def tag(doc, name, attr=None, body=None, namespace=None):
    """Wrapper to print xml tags and comments."""
    if attr is None:
        attr = {}
    start_tag(doc, name, attr, body, namespace)
    end_tag(doc, name, namespace)


def safe_iterator(node, tag=None):
    """Return an iterator that is compatible with Python 2.6"""
    if hasattr(node, "iter"):
        return node.iter(tag)
    else:
        return node.getiterator(tag)
