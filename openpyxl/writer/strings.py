# file openpyxl/writer/strings.py

"""Write the shared string table."""

# Python stdlib imports
from StringIO import StringIO

# package imports
from ..shared.xmltools import start_tag, end_tag, tag, XMLGenerator


def create_string_table(workbook):
    """Compile the string table for a workbook."""
    strings = set()
    for sheet in workbook.worksheets:
        for cell in sheet.get_cell_collection():
            if cell.data_type == cell.TYPE_STRING:
                strings.add(cell._value)
    return dict((key, i) for i, key in enumerate(strings))


def write_string_table(string_table):
    """Write the string table xml."""
    temp_buffer = StringIO()
    doc = XMLGenerator(temp_buffer, 'utf-8')
    start_tag(doc, 'sst', {'xmlns':
            'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            'uniqueCount': '%d' % len(string_table)})
    strings_to_write = sorted(string_table.iteritems(),
            key=lambda pair: pair[1])
    for key in [pair[0] for pair in strings_to_write]:
        start_tag(doc, 'si')
        if key.strip() != key:
            attr = {'xml:space': 'preserve'}
        else:
            attr = {}
        tag(doc, 't', attr, key)
        end_tag(doc, 'si')
    end_tag(doc, 'sst')
    string_table_xml = temp_buffer.getvalue()
    temp_buffer.close()
    return string_table_xml
