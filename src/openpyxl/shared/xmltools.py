# file openpyxl/shared/xmltools.py
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
