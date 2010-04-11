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
from openpyxl.shared.ooxml import NAMESPACES, ARC_CORE, ARC_WORKBOOK, ARC_APP, ARC_THEME, ARC_STYLE, ARC_SHARED_STRINGS
from openpyxl.shared.date_time import datetime_to_W3CDTF

def write_properties_core(properties):

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

    SubElement(root, 'Override', {'PartName' : ARC_THEME,
                                  'ContentType' : 'application/vnd.openxmlformats-officedocument.theme+xml'})
    SubElement(root, 'Override', {'PartName' : ARC_STYLE,
                                  'ContentType' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'})
    SubElement(root, 'Default', {'Extension' : 'rels',
                                  'ContentType' : 'application/vnd.openxmlformats-package.relationships+xml'})
    SubElement(root, 'Default', {'Extension' : 'xml',
                                  'ContentType' : 'application/xml'})
    SubElement(root, 'Override', {'PartName' : ARC_WORKBOOK,
                                  'ContentType' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'})
    SubElement(root, 'Override', {'PartName' : ARC_APP,
                                  'ContentType' : 'application/vnd.openxmlformats-officedocument.extended-properties+xml'})
    SubElement(root, 'Override', {'PartName' : ARC_CORE,
                                  'ContentType' : 'application/vnd.openxmlformats-package.core-properties+xml'})
    SubElement(root, 'Override', {'PartName' : ARC_SHARED_STRINGS,
                                  'ContentType' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'})

    for sheet_id in xrange(len(workbook.worksheets)):
        SubElement(root, 'Override', {'PartName' : '/xl/worksheets/sheet%d.xml' % (sheet_id + 1),
                                      'ContentType' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'})

    return get_document_content(root)

def write_properties_app(workbook):

    worksheets_count = len(workbook.worksheets)


    root = Element('Properties', {'xmlns' : 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
                                  'xmlns:vt' : 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'})

    SubElement(root, 'Application').text = 'Microsoft Excel'
    SubElement(root, 'DocSecurity').text = '0'
    SubElement(root, 'ScaleCrop').text = 'false'
    SubElement(root, 'Company')

    SubElement(root, 'LinksUpToDate').text = 'false'
    SubElement(root, 'SharedDoc').text = 'false'
    SubElement(root, 'HyperlinksChanged').text = 'false'
    SubElement(root, 'AppVersion').text = '12.0000'

    # heading pairs part
    heading_pairs = SubElement(root, 'HeadingPairs')
    vector = SubElement(heading_pairs, 'vt:vector', {'size' : '2',
                                                     'baseType' : 'variant'})
    variant = SubElement(vector, 'vt:variant')
    SubElement(variant, 'vt:lpstr').text = 'Worksheets'

    variant = SubElement(vector, 'vt:variant')
    SubElement(variant, 'vt:i4').text = '%d' % worksheets_count

    # title of parts
    title_of_parts = SubElement(root, 'TitlesOfParts')
    vector = SubElement(title_of_parts, 'vt:vector', {'size' : '%d' % worksheets_count,
                                                     'baseType' : 'lpstr'})

    for ws in workbook.worksheets:
        SubElement(vector, 'vt:lpstr').text = '%s' % ws.title

    return get_document_content(root)

def write_root_rels(workbook):

    root = Element('Relationships', {'xmlns' : "http://schemas.openxmlformats.org/package/2006/relationships"})

    SubElement(root, 'Relationship', {'Id' : 'rId1',
                                      'Type' : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
                                      'Target' : ARC_WORKBOOK})
    SubElement(root, 'Relationship', {'Id' : 'rId2',
                                      'Type' : 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
                                      'Target' : ARC_CORE})
    SubElement(root, 'Relationship', {'Id' : 'rId3',
                                      'Type' : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
                                      'Target' : ARC_APP})

    return get_document_content(root)
