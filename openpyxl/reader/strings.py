# file openpyxl/reader/strings.py

"""Read the shared strings table"""

# package imports
from ..shared.xmltools import fromstring, QName
from ..shared.ooxml import NAMESPACES


def read_string_table(xml_source):
    """Read in all shared strings in the table"""
    table = {}
    xmlns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    root = fromstring(text=xml_source)
    string_index_nodes = root.findall(QName(xmlns, 'si').text)
    for index, string_index_node in enumerate(string_index_nodes):
        table[index] = get_string(xmlns, string_index_node)
    return table


def get_string(xmlns, string_index_node):
    """Read the contents of a specific string index"""
    rich_nodes = string_index_node.findall(QName(xmlns, 'r').text)
    if rich_nodes:
        reconstructed_text = []
        for rich_node in rich_nodes:
            partial_text = get_text(xmlns, rich_node)
            reconstructed_text.append(partial_text)
        return ''.join(reconstructed_text)
    else:
        return get_text(xmlns, string_index_node)


def get_text(xmlns, rich_node):
    """Read rich text, discarding formatting if not disallowed"""
    text_node = rich_node.find(QName(xmlns, 't').text)
    partial_text = text_node.text
    if text_node.get(QName(NAMESPACES['xml'], 'space').text) != 'preserve':
        partial_text = partial_text.strip()
    return partial_text
