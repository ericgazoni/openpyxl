# file openpyxl/reader/style.py

"""Read shared style definitions"""

# package imports
from ..shared.xmltools import fromstring, QName
from ..style import Style, NumberFormat


def read_style_table(xml_source):
    """Read styles from the shared style table"""
    table = {}
    xmlns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    root = fromstring(xml_source)
    custom_num_formats = parse_custom_num_formats(root, xmlns)
    builtin_formats = NumberFormat._BUILTIN_FORMATS
    cell_xfs = root.find(QName(xmlns, 'cellXfs').text)
    cell_xfs_nodes = cell_xfs.findall(QName(xmlns, 'xf').text)
    for index, cell_xfs_node in enumerate(cell_xfs_nodes):
        new_style = Style()
        number_format_id = int(cell_xfs_node.get('numFmtId'))
        if number_format_id < 164:
            new_style.number_format.format_code = \
                    builtin_formats[number_format_id]
        else:
            new_style.number_format.format_code = \
                    custom_num_formats[number_format_id]
        table[index] = new_style
    return table


def parse_custom_num_formats(root, xmlns):
    """Read in custom numeric formatting rules from the shared style table"""
    custom_formats = {}
    num_fmts = root.find(QName(xmlns, 'numFmts').text)
    if num_fmts is not None:
        num_fmt_nodes = num_fmts.findall(QName(xmlns, 'numFmt').text)
        for num_fmt_node in num_fmt_nodes:
            custom_formats[int(num_fmt_node.get('numFmtId'))] = \
                    num_fmt_node.get('formatCode')
    return custom_formats
