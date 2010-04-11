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

from xml.etree.cElementTree import ElementTree, Element, SubElement

from openpyxl.shared.xmltools import get_document_content
from openpyxl.shared.ooxml import NAMESPACES, ARC_CORE
from openpyxl.shared.date_time import datetime_to_W3CDTF

def write_properties(properties):

    root = Element('cp:coreProperties', {'xmlns:cp': NAMESPACES['cp'],
                                         'xmlns:dc': NAMESPACES['dc'],
                                         'xmlns:dcterms': NAMESPACES['dcterms'],
                                         'xmlns:dcmitype': NAMESPACES['dcmitype'],
                                         'xmlns:xsi': NAMESPACES['xsi']})

    SubElement(root, 'dc:creator').text = properties.creator
    SubElement(root, 'cp:lastModifiedBy').text = properties.last_modified_by

    SubElement(root, 'dcterms:created', {'xsi:type': 'dcterms:W3CDTF'}).text = datetime_to_W3CDTF(properties.created)
    SubElement(root, 'dcterms:modified', {'xsi:type': 'dcterms:W3CDTF'}).text = datetime_to_W3CDTF(properties.modified)

    return get_document_content(root)


def write_content_types(workbook):

    root = Element('Types', {'xmlns' : "http://schemas.openxmlformats.org/package/2006/content-types"})

    SubElement(root, 'Override', {'PartName' : '/xl/theme/theme1.xml',
                                  'ContentType' : 'application/vnd.openxmlformats-officedocument.theme+xml'})
    SubElement(root, 'Override', {'PartName' : '/xl/styles.xml',
                                  'ContentType' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'})
    SubElement(root, 'Default', {'Extension' : 'rels',
                                  'ContentType' : 'application/vnd.openxmlformats-package.relationships+xml'})
    SubElement(root, 'Default', {'Extension' : 'xml',
                                  'ContentType' : 'application/xml'})
    SubElement(root, 'Override', {'PartName' : '/xl/workbook.xml',
                                  'ContentType' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'})
    SubElement(root, 'Override', {'PartName' : '/docProps/app.xml',
                                  'ContentType' : 'application/vnd.openxmlformats-officedocument.extended-properties+xml'})
    SubElement(root, 'Override', {'PartName' : '/docProps/core.xml',
                                  'ContentType' : 'application/vnd.openxmlformats-package.core-properties+xml'})

    for sheet_id in xrange(len(workbook.worksheets)):
        SubElement(root, 'Override', {'PartName' : '/xl/worksheets/sheet%d.xml' % (sheet_id + 1),
                                      'ContentType' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'})

    return get_document_content(root)
