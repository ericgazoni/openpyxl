# file openpyxl/reader/worksheet.py

"""Reader for a single worksheet."""

# Python stdlib imports
from xml.sax import parseString
from xml.sax.handler import ContentHandler

# package imports
from ..cell import Cell
from ..worksheet import Worksheet


class WorksheetReader(ContentHandler):
    """An xml.sax handler for reading xlsx worksheets."""

    def __init__(self, ws, string_table, style_table):
        ContentHandler.__init__(self)
        self.ws = ws
        self.string_table = string_table
        self.style_table = style_table
        self.read_value = False
        self.data_type = None
        self.style_id = None
        self.coordinate = None

    def startElement(self, name, attrs):
        """Start reading information for a cell defined in xml."""
        if name == 'c':
            self.coordinate = attrs.get('r')
            self.data_type = attrs.get('t', 'n')
            self.style_id = attrs.get('s')
            self.read_value = True

    def characters(self, value):
        """Unpack the string and style tables."""
        if self.read_value and value is not None:
            if self.data_type == Cell.TYPE_STRING:
                value = self.string_table.get(int(value))
            self.ws.cell(self.coordinate).value = value
            if self.style_id is not None:
                self.ws._styles[self.coordinate] = \
                        self.style_table.get(int(self.style_id))

    def endElement(self, name):
        """Stop reading information for a cell defined in xml."""
        if name == 'c':
            self.read_value = False


def read_worksheet(xml_source, parent, preset_title,
        string_table, style_table):
    """Read an xml worksheet"""
    ws = Worksheet(parent, preset_title)
    handler = WorksheetReader(ws, string_table, style_table)
    parseString(xml_source, handler)
    return ws
