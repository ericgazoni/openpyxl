# file openpyxl/writer/strings.py
from openpyxl.shared.xmltools import ElementTree, Element, SubElement, \
    get_document_content, get_tempfile, start_tag, end_tag, tag, XMLGenerator


def create_string_table(workbook):

    strings_list = []

    for sheet in workbook.worksheets:

        for cell in sheet.get_cell_collection():

            if cell.data_type == cell.TYPE_STRING:

                strings_list.append(cell._value)

    return dict((key, i) for i, key in enumerate(set(strings_list)))

def write_string_table(string_table):

    filename = get_tempfile()

    xml_file = open(filename, 'w')

    doc = XMLGenerator(out = xml_file, encoding = 'utf-8')

    start_tag(doc, 'sst', {'xmlns' : 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                           'uniqueCount' : '%d' % len(string_table)})

    strings_to_write = sorted(string_table.iteritems(), key = lambda pair:pair[1])

    for key in [key for (key, rank) in strings_to_write]:

        start_tag(doc, 'si')

        if key.strip() != key:
            attr = {'xml:space' : 'preserve'}
        else:
            attr = {}

        tag(doc, 't', attr = attr, body = key)

        end_tag(doc, 'si')

    end_tag(doc, 'sst')

    xml_file.close()

    return filename
