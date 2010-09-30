# file openpyxl/reader/excel.py

"""Read an xlsx file into Python"""

# Python stdlib imports
from zipfile import ZipFile, ZIP_DEFLATED

# package imports
from openpyxl.shared.ooxml import ARC_SHARED_STRINGS, ARC_CORE, ARC_APP, \
        ARC_WORKBOOK, PACKAGE_WORKSHEETS, ARC_STYLE
from openpyxl.workbook import Workbook
from openpyxl.reader.strings import read_string_table
from openpyxl.reader.style import read_style_table
from openpyxl.reader.workbook import read_sheets_titles, read_named_ranges, \
        read_properties_core
from openpyxl.reader.worksheet import read_worksheet


def load_workbook(filename):
    """Open the given filename and return the workbook

    :param filename: the path to open
    :type filename: string

    :rtype: :class:`openpyxl.workbook.Workbook`

    """
    archive = ZipFile(filename, 'r', ZIP_DEFLATED)
    wb = Workbook()
    try:
        # get workbook-level information
        wb.properties = read_properties_core(archive.read(ARC_CORE))
        try:
            string_table = read_string_table(archive.read(ARC_SHARED_STRINGS))
        except KeyError:
            string_table = {}
        style_table = read_style_table(archive.read(ARC_STYLE))
        wb._named_ranges = read_named_ranges(archive.read(ARC_WORKBOOK), wb)

        # get worksheets
        wb.worksheets = []  # remove preset worksheet
        sheet_names = read_sheets_titles(archive.read(ARC_APP))
        for i, sheet_name in enumerate(sheet_names):
            worksheet_path = '%s/%s' % \
                    (PACKAGE_WORKSHEETS, 'sheet%d.xml' % (i + 1))
            new_ws = read_worksheet(archive.read(worksheet_path),
                    wb, sheet_name, string_table, style_table)
            wb.add_sheet(new_ws, index=i)
    finally:
        archive.close()
    return wb
