# file openpyxl/style.py

# Copyright (c) 2010-2011 openpyxl
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#
# @license: http://www.opensource.org/licenses/mit-license.php
# @author: see AUTHORS file

"""Style and formatting option tracking."""

# Python stdlib imports
import re
from collections import namedtuple
from openpyxl.shared.compat import any


class Color(namedtuple('Color', ('index',))):
    BLACK = 'FF000000'
    WHITE = 'FFFFFFFF'
    RED = 'FFFF0000'
    DARKRED = 'FF800000'
    BLUE = 'FF0000FF'
    DARKBLUE = 'FF000080'
    GREEN = 'FF00FF00'
    DARKGREEN = 'FF008000'
    YELLOW = 'FFFFFF00'
    DARKYELLOW = 'FF808000'


class Font(namedtuple('Font', ('name',
                               'size',
                               'bold',
                               'italic',
                               'superscript',
                               'subscript',
                               'underline',
                               'strikethrough',
                               'color'))):
    """Font options used in styles."""
    UNDERLINE_NONE = 'none'
    UNDERLINE_DOUBLE = 'double'
    UNDERLINE_DOUBLE_ACCOUNTING = 'doubleAccounting'
    UNDERLINE_SINGLE = 'single'
    UNDERLINE_SINGLE_ACCOUNTING = 'singleAccounting'

    def __new__(cls,
                name='Calibri',
                size=11,
                bold=False,
                italic=False,
                superscript=False,
                subscript=False,
                underline=UNDERLINE_NONE,
                strikethrough=False,
                color=Color(Color.BLACK)
                ):
        return super(Font, cls).__new__(cls, name, size,
                                        bold, italic, superscript,
                                        subscript, underline,
                                        strikethrough, color)


class Fill(namedtuple('Fill', ('fill_type',
                               'rotation',
                               'start_color',
                               'end_color'))):
    """Area fill patterns for use in styles."""
    FILL_NONE = 'none'
    FILL_SOLID = 'solid'
    FILL_GRADIENT_LINEAR = 'linear'
    FILL_GRADIENT_PATH = 'path'
    FILL_PATTERN_DARKDOWN = 'darkDown'
    FILL_PATTERN_DARKGRAY = 'darkGray'
    FILL_PATTERN_DARKGRID = 'darkGrid'
    FILL_PATTERN_DARKHORIZONTAL = 'darkHorizontal'
    FILL_PATTERN_DARKTRELLIS = 'darkTrellis'
    FILL_PATTERN_DARKUP = 'darkUp'
    FILL_PATTERN_DARKVERTICAL = 'darkVertical'
    FILL_PATTERN_GRAY0625 = 'gray0625'
    FILL_PATTERN_GRAY125 = 'gray125'
    FILL_PATTERN_LIGHTDOWN = 'lightDown'
    FILL_PATTERN_LIGHTGRAY = 'lightGray'
    FILL_PATTERN_LIGHTGRID = 'lightGrid'
    FILL_PATTERN_LIGHTHORIZONTAL = 'lightHorizontal'
    FILL_PATTERN_LIGHTTRELLIS = 'lightTrellis'
    FILL_PATTERN_LIGHTUP = 'lightUp'
    FILL_PATTERN_LIGHTVERTICAL = 'lightVertical'
    FILL_PATTERN_MEDIUMGRAY = 'mediumGray'

    def __new__(cls,
                fill_type=FILL_NONE,
                rotation=0,
                start_color=Color(Color.WHITE),
                end_color=Color(Color.BLACK)):
        return super(Fill, cls).__new__(cls, fill_type, rotation,
                                        start_color, end_color)


class Border(namedtuple('Border', ('border_style', 'color'))):
    """Border options for use in styles."""
    BORDER_NONE = 'none'
    BORDER_DASHDOT = 'dashDot'
    BORDER_DASHDOTDOT = 'dashDotDot'
    BORDER_DASHED = 'dashed'
    BORDER_DOTTED = 'dotted'
    BORDER_DOUBLE = 'double'
    BORDER_HAIR = 'hair'
    BORDER_MEDIUM = 'medium'
    BORDER_MEDIUMDASHDOT = 'mediumDashDot'
    BORDER_MEDIUMDASHDOTDOT = 'mediumDashDotDot'
    BORDER_MEDIUMDASHED = 'mediumDashed'
    BORDER_SLANTDASHDOT = 'slantDashDot'
    BORDER_THICK = 'thick'
    BORDER_THIN = 'thin'

    def __new__(cls, border_style=BORDER_NONE,
                color=Color(Color.BLACK)):
        return super(Border, cls).__new__(cls, border_style, color)


class Borders(namedtuple('Borders', ('left',
                                     'right',
                                     'top',
                                     'bottom',
                                     'diagonal',
                                     'diagonal_direction',
                                     'all_borders',
                                     'outline',
                                     'inside',
                                     'vertical',
                                     'horizontal'))):
    """Border positioning for use in styles."""
    DIAGONAL_NONE = 0
    DIAGONAL_UP = 1
    DIAGONAL_DOWN = 2
    DIAGONAL_BOTH = 3

    def __new__(cls, left=Border(), right=Border(),
                     top=Border(), bottom=Border(),
                     diagonal=Border(), diagonal_direction=DIAGONAL_NONE,
                     all_borders=Border(), outline=Border(),
                     inside=Border(), vertical=Border(), horizontal=Border()):
        return super(Borders, cls).__new__(cls, left, right, top, bottom,
                                           diagonal, diagonal_direction,
                                           all_borders, outline, inside,
                                           vertical, horizontal)


class Alignment(namedtuple('Alignment', ('horizontal',
                                         'vertical',
                                         'text_rotation',
                                         'wrap_text',
                                         'shrink_to_fit',
                                         'indent'))):
    """Alignment options for use in styles."""
    HORIZONTAL_GENERAL = 'general'
    HORIZONTAL_LEFT = 'left'
    HORIZONTAL_RIGHT = 'right'
    HORIZONTAL_CENTER = 'center'
    HORIZONTAL_CENTER_CONTINUOUS = 'centerContinuous'
    HORIZONTAL_JUSTIFY = 'justify'
    VERTICAL_BOTTOM = 'bottom'
    VERTICAL_TOP = 'top'
    VERTICAL_CENTER = 'center'
    VERTICAL_JUSTIFY = 'justify'

    def __new__(cls, horizontal=HORIZONTAL_GENERAL,
                vertical=VERTICAL_BOTTOM,
                text_rotation=0,
                wrap_text=False,
                shrink_to_fit=False,
                indent=0):
        return super(Alignment, cls).__new__(cls, horizontal, vertical,
                                             text_rotation, wrap_text,
                                             shrink_to_fit, indent)


class NumberFormat(namedtuple('NumberFormat', ('format_code',))):
    """Numer formatting for use in styles."""
    FORMAT_GENERAL = 'General'
    FORMAT_TEXT = '@'
    FORMAT_NUMBER = '0'
    FORMAT_NUMBER_00 = '0.00'
    FORMAT_NUMBER_COMMA_SEPARATED1 = '#,##0.00'
    FORMAT_NUMBER_COMMA_SEPARATED2 = '#,##0.00_-'
    FORMAT_PERCENTAGE = '0%'
    FORMAT_PERCENTAGE_00 = '0.00%'
    FORMAT_DATE_YYYYMMDD2 = 'yyyy-mm-dd'
    FORMAT_DATE_YYYYMMDD = 'yy-mm-dd'
    FORMAT_DATE_DDMMYYYY = 'dd/mm/yy'
    FORMAT_DATE_DMYSLASH = 'd/m/y'
    FORMAT_DATE_DMYMINUS = 'd-m-y'
    FORMAT_DATE_DMMINUS = 'd-m'
    FORMAT_DATE_MYMINUS = 'm-y'
    FORMAT_DATE_XLSX14 = 'mm-dd-yy'
    FORMAT_DATE_XLSX15 = 'd-mmm-yy'
    FORMAT_DATE_XLSX16 = 'd-mmm'
    FORMAT_DATE_XLSX17 = 'mmm-yy'
    FORMAT_DATE_XLSX22 = 'm/d/yy h:mm'
    FORMAT_DATE_DATETIME = 'd/m/y h:mm'
    FORMAT_DATE_TIME1 = 'h:mm AM/PM'
    FORMAT_DATE_TIME2 = 'h:mm:ss AM/PM'
    FORMAT_DATE_TIME3 = 'h:mm'
    FORMAT_DATE_TIME4 = 'h:mm:ss'
    FORMAT_DATE_TIME5 = 'mm:ss'
    FORMAT_DATE_TIME6 = 'h:mm:ss'
    FORMAT_DATE_TIME7 = 'i:s.S'
    FORMAT_DATE_TIME8 = 'h:mm:ss@'
    FORMAT_DATE_TIMEDELTA = '[hh]:mm:ss'
    FORMAT_DATE_YYYYMMDDSLASH = 'yy/mm/dd@'
    FORMAT_CURRENCY_USD_SIMPLE = '"$"#,##0.00_-'
    FORMAT_CURRENCY_USD = '$#,##0_-'
    FORMAT_CURRENCY_EUR_SIMPLE = '[$EUR ]#,##0.00_-'
    _BUILTIN_FORMATS = {
        0: 'General',
        1: '0',
        2: '0.00',
        3: '#,##0',
        4: '#,##0.00',

        9: '0%',
        10: '0.00%',
        11: '0.00E+00',
        12: '# ?/?',
        13: '# ??/??',
        14: 'mm-dd-yy',
        15: 'd-mmm-yy',
        16: 'd-mmm',
        17: 'mmm-yy',
        18: 'h:mm AM/PM',
        19: 'h:mm:ss AM/PM',
        20: 'h:mm',
        21: 'h:mm:ss',
        22: 'm/d/yy h:mm',

        37: '#,##0 (#,##0)',
        38: '#,##0 [Red](#,##0)',
        39: '#,##0.00(#,##0.00)',
        40: '#,##0.00[Red](#,##0.00)',

        41: '_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)',
        42: '_("$"* #,##0_);_("$"* \(#,##0\);_("$"* "-"_);_(@_)',
        43: '_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)',

        44: '_("$"* #,##0.00_)_("$"* \(#,##0.00\)_("$"* "-"??_)_(@_)',
        45: 'mm:ss',
        46: '[h]:mm:ss',
        47: 'mmss.0',
        48: '##0.0E+0',
        49: '@', }
    _BUILTIN_FORMATS_REVERSE = dict(
            [(value, key) for key, value in _BUILTIN_FORMATS.items()])

    DATE_INDICATORS = 'dmyhs'
    BAD_DATE_RE = re.compile(r'(\[|").*[dmhys].*(\]|")')

    def __new__(cls, format_code=FORMAT_GENERAL):
        return super(NumberFormat, cls).__new__(cls, format_code)

    @property
    def _format_index(self):
        return self.builtin_format_id(self.format_code)

    def builtin_format_code(self, index):
        """Return one of the standard format codes by index."""
        return self._BUILTIN_FORMATS[index]

    def is_builtin(self, format=None):
        """Check if a format code is a standard format code."""
        if format is None:
            format = self.format_code
        return format in self._BUILTIN_FORMATS.values()

    def builtin_format_id(self, format):
        """Return the id of a standard style."""
        return self._BUILTIN_FORMATS_REVERSE.get(format, None)

    def is_date_format(self, format=None):
        """Check if the number format is actually representing a date."""
        if format is None:
            format = self.format_code

        if any([x in format for x in self.DATE_INDICATORS]):
            if self.BAD_DATE_RE.search(format) is None:
                return True

        return False


class Protection(namedtuple('Protection', ('locked', 'hidden'))):
    """Protection options for use in styles."""
    PROTECTION_INHERIT = 'inherit'
    PROTECTION_PROTECTED = 'protected'
    PROTECTION_UNPROTECTED = 'unprotected'

    def __new__(cls, locked=PROTECTION_INHERIT, hidden=PROTECTION_INHERIT):
        return super(Protection, cls).__new__(cls, locked, hidden)


class Style(namedtuple('Style', ('font', 'fill', 'borders', 'alignment',
                                 'number_format', 'protection'))):
    """Style object containing all formatting details."""

    def __new__(cls, font=Font(),
                fill=Fill(),
                borders=Borders(),
                alignment=Alignment(),
                number_format=NumberFormat(),
                protection=Protection()):
        return super(Style, cls).__new__(cls, font, fill, borders, alignment,
                                         number_format, protection)

    def copy(self, **kwargs):
        """
        returns a `openpyxl.style.Style` which inherits from the current style,
        but takes alterations as kwargs.

            >>> s = Style(font=Font(size=31))
            >>> s2 = s.copy(font=Font(bold=True))
            >>> print s2
            Style(font=Font(size=31, bold=True, ...), ...)
        """
        return self._replace(**kwargs)

DEFAULTS = Style()
