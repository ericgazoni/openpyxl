# file openpyxl/shared/xmltools.py

"""Shared xml tools.

Shortcut functions taken from:
    http://lethain.com/entry/2009/jan/22/handling-very-large-csv-and-xml-files-in-python/

"""

# Python stdlib imports
from __future__ import with_statement
from xml.sax.xmlreader import AttributesNSImpl
from xml.sax.saxutils import XMLGenerator
try:
    from xml.etree.ElementTree import ElementTree, Element, SubElement, \
            QName, fromstring, tostring
except ImportError:
    from cElementTree import ElementTree, Element, SubElement, \
            QName, fromstring, tostring

# package imports
from .. import __name__ as prefix


def get_document_content(xml_node):
    """Print nicely formatted xml to a string."""
    pretty_indent(xml_node)
    return tostring(xml_node, 'utf-8')


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


def end_tag(doc, name, namespace=None):
    """Wrapper to close an xml tag."""
    doc.endElementNS((namespace, name), name)


def tag(doc, name, attr=None, body=None, namespace=None):
    """Wrapper to print xml tags and comments."""
    if attr is None:
        attr = {}
    start_tag(doc, name, attr, body, namespace)
    end_tag(doc, name, namespace)
