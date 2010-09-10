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
from openpyxl.shared.ooxml import NAMESPACES, ARC_CORE, ARC_APP
from openpyxl.workbook import DocumentProperties
from openpyxl.shared.date_time import W3CDTF_to_datetime
from openpyxl.namedrange import NamedRange, split_named_range

def read_properties_core(xml_source):

    properties = DocumentProperties()

    root = fromstring(text = xml_source)

    creator_node = root.find(QName(NAMESPACES['dc'], 'creator').text)

    if creator_node is not None:
        properties.creator = creator_node.text
    else:
        properties.creator = ''

    last_modified_by_node = root.find(QName(NAMESPACES['cp'], 'lastModifiedBy').text)

    if last_modified_by_node is not None:
        properties.last_modified_by = last_modified_by_node.text
    else:
        properties.last_modified_by = ''

    properties.created = W3CDTF_to_datetime(root.find(QName(NAMESPACES['dcterms'], 'created').text).text)
    properties.modified = W3CDTF_to_datetime(root.find(QName(NAMESPACES['dcterms'], 'modified').text).text)

    return properties

def read_sheets_titles(xml_source):

    root = fromstring(text = xml_source)

    titles_root = root.find(QName('http://schemas.openxmlformats.org/officeDocument/2006/extended-properties', 'TitlesOfParts').text)

    vector = titles_root.find(QName(NAMESPACES['vt'], 'vector').text)

    parts, names = get_number_of_parts(xml_source)
    # we can't assume 'Worksheets' to be written in english, 
    # but it's always the first item of the parts list (see bug #22)
    size = parts[names[0]]

    children = [c.text for c in vector.getchildren()]

    return children[:size]

def read_named_ranges(xml_source, workbook):

    named_ranges = []

    root = fromstring(text = xml_source)

    names_root = root.find(QName('http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'definedNames').text)

    BUGGY_NAMED_RANGES = ['NA()', '#REF!']
    DISCARDED_RANGES = ['Excel_BuiltIn', ]

    if names_root:

        for name_node in names_root.getchildren():

            range_name = name_node.get('name')

            discard = False

            for bug in BUGGY_NAMED_RANGES:
                if bug in name_node.text:
                    discard = True

            for bug in DISCARDED_RANGES:
                if bug in range_name:
                    discard = True

            if not discard:

                worksheet_name, column, row = split_named_range(range_string = name_node.text)
                worksheet = workbook.get_sheet_by_name(worksheet_name)
                range = '%s%s' % (column, row)

                named_range = NamedRange(name = range_name,
                                         worksheet = worksheet,
                                         range = range)

                named_ranges.append(named_range)

    return named_ranges

def get_number_of_parts(xml_source):

    parts_size = {}
    parts_names = []

    root = fromstring(text = xml_source)

    heading_pairs = root.find(QName('http://schemas.openxmlformats.org/officeDocument/2006/extended-properties', 'HeadingPairs').text)

    vector = heading_pairs.find(QName(NAMESPACES['vt'], 'vector').text)

    children = vector.getchildren()

    for child_id in range(0, len(children), 2):

        part_name = children[child_id].find(QName(NAMESPACES['vt'], 'lpstr').text).text
        if not part_name in parts_names:
            parts_names.append(part_name)

        part_size = int(children[child_id + 1].find(QName(NAMESPACES['vt'], 'i4').text).text)
        parts_size[part_name] = part_size

    return parts_size, parts_names
