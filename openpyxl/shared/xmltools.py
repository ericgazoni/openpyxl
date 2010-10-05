# file openpyxl/shared/xmltools.py

"""Shared xml tools.

Shortcut functions taken from:
    http://lethain.com/entry/2009/jan/22/handling-very-large-csv-and-xml-files-in-python/

"""

# Python stdlib imports
from __future__ import with_statement
from os import close, remove
from tempfile import mkstemp
from xml.sax.xmlreader import AttributesNSImpl
from xml.sax.saxutils import XMLGenerator
import atexit
try:
    from xml.etree.ElementTree import ElementTree, Element, SubElement, \
            QName, fromstring
except ImportError:
    from cElementTree import ElementTree, Element, SubElement, \
            QName, fromstring

# package imports
from openpyxl import __name__ as prefix

# constants
XML_TEMP_FILES = []


@atexit.register
def cleanup_tempfiles():
    """Delete any temporary files when the program finishes."""
    for handle, filename in XML_TEMP_FILES:
        try:
            close(handle)
            remove(filename)
        except:
            pass


def get_tempfile():
    """Create a temporary file."""
    fd, filename = mkstemp(prefix=prefix, text=True)
    XML_TEMP_FILES.append((fd, filename))
    return filename


def get_document_content(xml_node):
    """Print nicely formatted xml to a temp file."""
    pretty_indent(xml_node)
    filename = get_tempfile()
    with open(filename, 'w') as handle:
        ElementTree(xml_node).write(handle, encoding='UTF-8')
    return filename


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
