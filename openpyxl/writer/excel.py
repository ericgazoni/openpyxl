# file openpyxl/writer/excel.py

from openpyxl.shared.zip import ZipArchive
from openpyxl.shared.xmltools import cleanup_tempfiles

from openpyxl.shared.ooxml import ARC_SHARED_STRINGS, ARC_CONTENT_TYPES, ARC_ROOT_RELS, ARC_WORKBOOK_RELS, ARC_APP, ARC_CORE, ARC_THEME, ARC_STYLE, ARC_WORKBOOK, PACKAGE_WORKSHEETS

from openpyxl.writer.strings import create_string_table, write_string_table
from openpyxl.writer.workbook import write_content_types, write_root_rels, write_workbook_rels, write_properties_app, write_properties_core, write_workbook
from openpyxl.writer.theme import write_theme
from openpyxl.writer.styles import create_style_table, write_style_table
from openpyxl.writer.worksheet import write_worksheet, write_worksheet_rels

def save_workbook(workbook, filename):
    """Save the given workbook on the filesystem under the name fielename

    :param workbook: the workbook to save
    :type workbook: :class:`openpyxl.workbook.Workbook`

    :param filename: the path to which save the workbook
    :type filename: string

    :rtype: bool
    """

    ew = ExcelWriter(workbook)
    ew.save(filename)

    return True

class ExcelWriter(object):

    def __init__(self, workbook):

        self.workbook = workbook

    def save(self, filename):

        archive = ZipArchive(filename = filename, mode = 'w')

        # cleanup all worksheets
        for ws in self.workbook.worksheets:
            ws.garbage_collect()

        shared_string_table = create_string_table(workbook = self.workbook)
        shared_style_table = create_style_table(workbook = self.workbook)

        # write shared strings
        archive.add_from_file(arc_name = ARC_SHARED_STRINGS,
                                content = write_string_table(string_table = shared_string_table))

        # write content types
        archive.add_from_file(arc_name = ARC_CONTENT_TYPES,
                                content = write_content_types(workbook = self.workbook))

        # write relationships
        archive.add_from_file(arc_name = ARC_ROOT_RELS,
                                content = write_root_rels(workbook = self.workbook))

        archive.add_from_file(arc_name = ARC_WORKBOOK_RELS,
                                content = write_workbook_rels(workbook = self.workbook))

        # write properties
        archive.add_from_file(arc_name = ARC_APP,
                                content = write_properties_app(workbook = self.workbook))

        archive.add_from_file(arc_name = ARC_CORE,
                                content = write_properties_core(properties = self.workbook.properties))

        # write theme
        archive.add_from_file(arc_name = ARC_THEME, content = write_theme())

        # write style
        archive.add_from_file(arc_name = ARC_STYLE, content = write_style_table(style_table = shared_style_table))

        # write workbook
        archive.add_from_file(arc_name = ARC_WORKBOOK, content = write_workbook(workbook = self.workbook))

        # write sheets
        style_id_by_hash = dict([(style.__crc__(), id) for style, id in shared_style_table.iteritems()])

        for i, sheet in enumerate(self.workbook.worksheets):
            archive.add_from_file(arc_name = PACKAGE_WORKSHEETS + '/sheet%d.xml' % (i + 1),
                                    content = write_worksheet(worksheet = sheet,
                                                              string_table = shared_string_table,
                                                              style_table = style_id_by_hash))
            if sheet.relationships:
                archive.add_from_file(arc_name = PACKAGE_WORKSHEETS + '/_rels/sheet%d.xml.rels' % (i + 1),
                                        content = write_worksheet_rels(worksheet = sheet))

