# file openpyxl/writer/excel.py

"""Write a .xlsx file."""

# Python stdlib imports
from zipfile import ZipFile, ZIP_DEFLATED

# package imports
from openpyxl.shared.ooxml import ARC_SHARED_STRINGS, ARC_CONTENT_TYPES, \
        ARC_ROOT_RELS, ARC_WORKBOOK_RELS, ARC_APP, ARC_CORE, ARC_THEME, \
        ARC_STYLE, ARC_WORKBOOK, PACKAGE_WORKSHEETS
from openpyxl.writer.strings import create_string_table, write_string_table
from openpyxl.writer.workbook import write_content_types, write_root_rels, \
        write_workbook_rels, write_properties_app, write_properties_core, \
        write_workbook
from openpyxl.writer.theme import write_theme
from openpyxl.writer.styles import create_style_table, write_style_table
from openpyxl.writer.worksheet import write_worksheet, write_worksheet_rels


class ExcelWriter(object):
    """Write a workbook object to an Excel file."""

    def __init__(self, workbook):
        self.workbook = workbook

    def write_data(self, archive):
        """Write the various xml files into the zip archive."""
        # cleanup all worksheets
        for ws in self.workbook.worksheets:
            ws.garbage_collect()
        shared_string_table = create_string_table(self.workbook)
        shared_style_table = create_style_table(self.workbook)
        archive.write(write_string_table(shared_string_table),
                ARC_SHARED_STRINGS)
        archive.write(write_content_types(self.workbook), ARC_CONTENT_TYPES)
        archive.write(write_root_rels(self.workbook), ARC_ROOT_RELS)
        archive.write(write_workbook_rels(self.workbook), ARC_WORKBOOK_RELS)
        archive.write(write_properties_app(self.workbook), ARC_APP)
        archive.write(write_properties_core(self.workbook.properties), ARC_CORE)
        archive.write(write_theme(), ARC_THEME)
        archive.write(write_style_table(shared_style_table), ARC_STYLE)
        archive.write(write_workbook(self.workbook), ARC_WORKBOOK)
        style_id_by_hash = dict([(style.__crc__(), style_id) for
                style, style_id in shared_style_table.iteritems()])
        for i, sheet in enumerate(self.workbook.worksheets):
            archive.write(write_worksheet(sheet, shared_string_table,
                    style_id_by_hash),
                    PACKAGE_WORKSHEETS + '/sheet%d.xml' % (i + 1))
            if sheet.relationships:
                archive.write(write_worksheet_rels(sheet),
                PACKAGE_WORKSHEETS + '/_rels/sheet%d.xml.rels' % (i + 1))

    def save(self, filename):
        """Write data into the archive."""
        try:
            archive = ZipFile(filename, 'w', ZIP_DEFLATED, False)
            self.write_data(archive)
        finally:
            archive.close()


def save_workbook(workbook, filename):
    """Save the given workbook on the filesystem under the name filename.

    :param workbook: the workbook to save
    :type workbook: :class:`openpyxl.workbook.Workbook`

    :param filename: the path to which save the workbook
    :type filename: string

    :rtype: bool

    """
    writer = ExcelWriter(workbook)
    writer.save(filename)
    return True
