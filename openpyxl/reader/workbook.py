# file openpyxl/reader/workbook.py

"""Read in global settings to be maintained by the workbook object."""

# package imports
from ..shared.xmltools import fromstring, QName
from ..shared.ooxml import NAMESPACES
from ..workbook import DocumentProperties
from ..shared.date_time import W3CDTF_to_datetime
from ..namedrange import NamedRange, split_named_range

# constants
BUGGY_NAMED_RANGES = ['NA()', '#REF!']
DISCARDED_RANGES = ['Excel_BuiltIn', ]


def read_properties_core(xml_source):
    """Read assorted file properties."""
    properties = DocumentProperties()
    root = fromstring(xml_source)
    creator_node = root.find(QName(NAMESPACES['dc'], 'creator').text)
    if creator_node is not None:
        properties.creator = creator_node.text
    else:
        properties.creator = ''
    last_modified_by_node = root.find(
            QName(NAMESPACES['cp'], 'lastModifiedBy').text)
    if last_modified_by_node is not None:
        properties.last_modified_by = last_modified_by_node.text
    else:
        properties.last_modified_by = ''
    properties.created = W3CDTF_to_datetime(
            root.find(QName(NAMESPACES['dcterms'], 'created').text).text)
    properties.modified = W3CDTF_to_datetime(
            root.find(QName(NAMESPACES['dcterms'], 'modified').text).text)
    return properties


def get_number_of_parts(xml_source):
    """Get a list of contents of the workbook."""
    parts_size = {}
    parts_names = []
    root = fromstring(xml_source)
    heading_pairs = root.find(QName('http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
            'HeadingPairs').text)
    vector = heading_pairs.find(QName(NAMESPACES['vt'], 'vector').text)
    children = vector.getchildren()
    for child_id in range(0, len(children), 2):
        part_name = children[child_id].find(QName(NAMESPACES['vt'],
                'lpstr').text).text
        if not part_name in parts_names:
            parts_names.append(part_name)
        part_size = int(children[child_id + 1].find(QName(
                NAMESPACES['vt'], 'i4').text).text)
        parts_size[part_name] = part_size
    return parts_size, parts_names


def read_sheets_titles(xml_source):
    """Read titles for all sheets."""
    root = fromstring(xml_source)
    titles_root = root.find(QName('http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
            'TitlesOfParts').text)
    vector = titles_root.find(QName(NAMESPACES['vt'], 'vector').text)
    parts, names = get_number_of_parts(xml_source)

    # we can't assume 'Worksheets' to be written in english,
    # but it's always the first item of the parts list (see bug #22)
    size = parts[names[0]]
    children = [c.text for c in vector.getchildren()]
    return children[:size]


def read_named_ranges(xml_source, workbook):
    """Read named ranges, excluding poorly defined ranges."""
    named_ranges = []
    root = fromstring(xml_source)
    names_root = root.find(QName('http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            'definedNames').text)
    if names_root:
        bad_ranges = BUGGY_NAMED_RANGES
        bad_ranges.extend(DISCARDED_RANGES)
        for name_node in names_root.getchildren():
            range_name = name_node.get('name')
            for bad_range in bad_ranges:
                if bad_range in name_node.text:
                    # this is a bad range, skip it
                    break
            else:
                # this is a good named range, read it
                worksheet_name, column, row = split_named_range(name_node.text)
                worksheet = workbook.get_sheet_by_name(worksheet_name)
                xlrange = '%s%s' % (column, row)
                named_range = NamedRange(range_name, worksheet, xlrange)
                named_ranges.append(named_range)
    return named_ranges
