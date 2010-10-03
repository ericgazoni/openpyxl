# file openpyxl/worksheet.py

"""Worksheet is the 2nd-level container in Excel."""

# Python stdlib imports
import re

# package imports
from openpyxl.cell import Cell, coordinate_from_string, \
    column_index_from_string, get_column_letter
from openpyxl.shared.exc import SheetTitleException, \
    InsufficientCoordinatesException, CellCoordinatesException, \
    NamedRangeException
from openpyxl.shared.password_hasher import hash_password
from openpyxl.style import Style


class Relationship(object):
    """Represents many kinds of relationships."""
    # TODO: Use this object for workbook relationships as well as
    # worksheet relationships
    TYPES = {
        'hyperlink': 'http://schemas.openxmlformats.org/'
                'officeDocument/2006/relationships/hyperlink',
        #'worksheet': 'http://schemas.openxmlformats.org/'
        #        'officeDocument/2006/relationships/worksheet',
        #'sharedStrings': 'http://schemas.openxmlformats.org/'
        #        'officeDocument/2006/relationships/sharedStrings',
        #'styles': 'http://schemas.openxmlformats.org/'
        #        'officeDocument/2006/relationships/styles',
        #'theme': 'http://schemas.openxmlformats.org/'
        #        'officeDocument/2006/relationships/theme',
    }

    def __init__(self, rel_type):
        if rel_type not in self.TYPES:
            raise Exception("Invalid relationship type %s" % rel_type)
        self.type = self.TYPES[rel_type]
        self.target = ""
        self.target_mode = ""
        self.id = ""


class PageSetup(object):
    """Information about page layout for this sheet"""
    pass


class HeaderFooter(object):
    """Information about the header/footer for this sheet."""
    pass


class SheetView(object):
    """Information about the visible portions of this sheet."""
    pass


class RowDimension(object):
    """Information about the display properties of a row."""
    __slots__ = ('row_index',
                 'height',
                 'visible',
                 'outline_level',
                 'collapsed',
                 'style_index',)

    def __init__(self, index=0):
        self.row_index = index
        self.height = -1
        self.visible = True
        self.outline_level = 0
        self.collapsed = False
        self.style_index = None


class ColumnDimension(object):
    """Information about the display properties of a column."""
    __slots__ = ('column_index',
                 'width',
                 'auto_size',
                 'visible',
                 'outline_level',
                 'collapsed',
                 'style_index',)

    def __init__(self, index='A'):
        self.column_index = index
        self.width = -1
        self.auto_size = False
        self.visible = True
        self.outline_level = 0
        self.collapsed = False
        self.style_index = 0


class PageMargins(object):
    """Information about page margins for view/print layouts."""

    def __init__(self):
        self.left = self.right = 0.7
        self.top = self.bottom = 0.75
        self.header = self.footer = 0.3


class SheetProtection(object):
    """Information about protection of various aspects of a sheet."""

    def __init__(self):
        self.sheet = False
        self.objects = False
        self.scenarios = False
        self.format_cells = False
        self.format_columns = False
        self.format_rows = False
        self.insert_columns = False
        self.insert_rows = False
        self.insert_hyperlinks = False
        self.delete_columns = False
        self.delete_rows = False
        self.select_locked_cells = False
        self.sort = False
        self.auto_filter = False
        self.pivot_tables = False
        self.select_unlocked_cells = False
        self._password = ''

    def set_password(self, value='', already_hashed=False):
        """Set a password on this sheet."""
        if not already_hashed:
            value = hash_password(value)
        self._password = value

    def _set_raw_password(self, value):
        """Set a password directly, forcing a hash step."""
        self.set_password(value, already_hashed=False)

    def _get_raw_password(self):
        """Return the password value, regardless of hash."""
        return self._password

    password = property(_get_raw_password, _set_raw_password,
            'get/set the password (if already hashed, '
            'use set_password() instead)')


class Worksheet(object):
    """Represents a worksheet.

    Do not create worksheets yourself,
    use :func:`openpyxl.workbook.Workbook.create_sheet` instead

    """
    BREAK_NONE = 0
    BREAK_ROW = 1
    BREAK_COLUMN = 2

    SHEETSTATE_VISIBLE = 'visible'
    SHEETSTATE_HIDDEN = 'hidden'
    SHEETSTATE_VERYHIDDEN = 'veryHidden'

    def __repr__(self):
        return u'<Worksheet %s>' % self.title

    def __init__(self, parent_workbook, title='Sheet'):
        self._parent = parent_workbook
        self._title = ''
        if not title:
            self.title = 'Sheet%d' % (1 + len(self._parent.worksheets))
        else:
            self.title = title
        self.row_dimensions = {}
        self.column_dimensions = {}
        self._cells = {}
        self._styles = {}
        self.relationships = []
        self.selected_cell = 'A1'
        self.active_cell = 'A1'
        self.sheet_state = self.SHEETSTATE_VISIBLE
        self.page_setup = PageSetup()
        self.page_margins = PageMargins()
        self.header_footer = HeaderFooter()
        self.sheet_view = SheetView()
        self.protection = SheetProtection()
        self.show_gridlines = True
        self.print_gridlines = False
        self.show_summary_below = True
        self.show_summary_right = True
        self.default_row_dimension = RowDimension()
        self.default_column_dimension = ColumnDimension()

    def garbage_collect(self):
        """Delete cells that are not storing a value."""
        delete_list = [coordinate for coordinate, cell in
                self._cells.iteritems() if cell.value == '']
        for coordinate in delete_list:
            del self._cells[coordinate]

    def get_cell_collection(self):
        """Return a list of values in cells in this worksheet."""
        return self._cells.values()

    def _set_title(self, value):
        """Set a sheet title, ensuring it is valid."""
        bad_title_char_re = re.compile(r'[\\*?:/\[\]]')
        if bad_title_char_re.search(value):
            msg = 'Invalid character found in sheet title'
            raise SheetTitleException(msg)

        # check if sheet_name already exists
        # do this *before* length check
        if self._parent.get_sheet_by_name(value):
            # use name, but append with lowest possible integer
            i = 1
            while self._parent.get_sheet_by_name('%s%d' % (value, i)):
                i += 1
            value = '%s%d' % (value, i)
        if len(value) > 31:
            msg = 'Maximum 31 characters allowed in sheet title'
            raise SheetTitleException(msg)
        self._title = value

    def _get_title(self):
        """Return the title for this sheet."""
        return self._title

    title = property(_get_title, _set_title,
            'Get or set the title of the worksheet.  '
            'Limited to 31 characters, no special characters.')

    def cell(self, coordinate=None, row=None, column=None):
        """Returns a cell object based on the given coordinates.

        Usage: cell(coodinate='A15') *or* cell(row=15, column=0)

        If `coordinates` are not given, then row *and* column must be given.

        Cells are kept in a dictionary which is empty at the worksheet
        creation.  Calling `cell` creates the cell in memory when they
        are first accessed, to reduce memory usage.

        :param coordinate: coordinates of the cell (e.g. 'B12')
        :type coordinate: string

        :param row: row index of the cell (e.g. 4)
        :type row: int

        :param column: column index of the cell (e.g. 3)
        :type column: int

        :raise: InsufficientCoordinatesException when coordinate or
        (row and column) are not given

        :rtype: :class:`openpyxl.cell.Cell`

        """
        if not coordinate:
            if not (row and column):
                msg = "You have to provide a value either for " \
                        "'coordinate' or for 'row' *and* 'column'"
                raise InsufficientCoordinatesException(msg)
            else:
                coordinate = '%s%s' % (get_column_letter(column), row)

        if not coordinate in self._cells:
            column, row = coordinate_from_string(coordinate)
            new_cell = Cell(self, column, row)
            self._cells[coordinate] = new_cell
            if column not in self.column_dimensions:
                self.column_dimensions[column] = ColumnDimension(column)
            if row not in self.row_dimensions:
                self.row_dimensions[row] = RowDimension(row)
        return self._cells[coordinate]

    def get_highest_row(self):
        """Returns the maximum row index containing data"""
        max_row = 1
        for row_dim in self.row_dimensions.values():
            if row_dim.row_index > max_row:
                max_row = row_dim.row_index
        return max_row

    def get_highest_column(self):
        """Get the largest value for column currently stored."""
        max_col = 1
        for col_dim in self.column_dimensions.values():
            col_index = column_index_from_string(col_dim.column_index)
            if col_index > max_col:
                max_col = col_index
        return max_col

    def calculate_dimension(self):
        """Return the minimum bounding range for all cells containing data."""
        return 'A1:%s%d' % (get_column_letter(self.get_highest_column()),
                            self.get_highest_row())

    def range(self, range_string, row=0, column=0):
        """Returns a 2D array of cells, with optional row and column offsets.

        :param range_string: cell range string or `named range` name
        :type range_string: string

        :param row: number of rows to offset
        :type row: int

        :param column: number of columns to offset
        :type column: int

        :rtype: tuples of tuples of :class:`openpyxl.cell.Cell`

        """
        if ':' in range_string:
            # R1C1 range
            result = []
            min_range, max_range = range_string.split(':')
            min_col, min_row = coordinate_from_string(min_range)
            max_col, max_row = coordinate_from_string(max_range)
            if column:
                min_col = get_column_letter(
                        column_index_from_string(min_col) + column)
                max_col = get_column_letter(
                        column_index_from_string(max_col) + column)
            min_col = column_index_from_string(min_col)
            max_col = column_index_from_string(max_col)
            cache_cols = {}
            for col in xrange(min_col, max_col + 1):
                cache_cols[col] = get_column_letter(col)
            rows = xrange(min_row + row, max_row + row + 1)
            cols = xrange(min_col, max_col + 1)
            for row in rows:
                new_row = []
                for col in cols:
                    new_row.append(self.cell('%s%s' % (cache_cols[col], row)))
                result.append(tuple(new_row))
            return tuple(result)
        else:
            try:
                return self.cell(coordinate=range_string, row=row,
                        column=column)
            except CellCoordinatesException:
                pass
            # named range
            named_range = self._parent.get_named_range(range_string)
            if named_range is None:
                msg = '%s is not a valid range name' % range_string
                raise NamedRangeException(msg)
            if named_range.worksheet is not self:
                msg = 'Range %s is not defined on worksheet %s' % \
                        (range_string, self.title)
                raise NamedRangeException(msg)
            return self.cell(named_range.range)

    def get_style(self, coordinate):
        """Return the style object for the specified cell."""
        if not coordinate in self._styles:
            self._styles[coordinate] = Style()
        return self._styles[coordinate]

    def create_relationship(self, rel_type):
        """Add a relationship for this sheet."""
        rel = Relationship(rel_type)
        self.relationships.append(rel)
        rel_id = self.relationships.index(rel)
        rel.id = 'rId' + str(rel_id + 1)
        return self.relationships[rel_id]
