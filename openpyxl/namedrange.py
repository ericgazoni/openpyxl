# file openpyxl/namedrange.py

"""Track named groups of cells in a worksheet"""

# Python stdlib imports
import re

# package imports
from .shared.exc import NamedRangeException

# constants
NAMED_RANGE_RE = re.compile("'?([^']*)'?!\$([A-Za-z]+)\$([0-9]+)")


class NamedRange(object):
    """A named group of cells"""
    __slots__ = ('name', 'worksheet', 'range', 'local_only')

    def __init__(self, name, worksheet, range):
        self.name = name
        self.worksheet = worksheet
        self.range = range
        self.local_only = False

    def __str__(self):
        return u'%s!%s' % (self.worksheet.title, self.range)


def split_named_range(range_string):
    """Separate a named range into its component parts"""
    match = NAMED_RANGE_RE.match(range_string)
    if not match:
        raise NamedRangeException('Invalid named range string: "%s"')
    else:
        sheet_name, column, row = match.groups()
        return sheet_name, column, int(row)
