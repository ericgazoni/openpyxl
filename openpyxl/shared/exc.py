# file openpyxl/shared/exc.py

"""Definitions for openpyxl shared exception classes."""


class CellCoordinatesException(Exception):
    """Error for converting between numeric and A1-style cell references."""
    pass


class ColumnStringIndexException(Exception):
    """Error for bad column names in A1-style cell references."""
    pass


class DataTypeException(Exception):
    """Error for any data type inconsistencies."""
    pass


class NamedRangeException(Exception):
    """Error for badly formatted named ranges."""
    pass


class SheetTitleException(Exception):
    """Error for bad sheet names."""
    pass


class InsufficientCoordinatesException(Exception):
    """Error for partially specified cell coordinates."""
    pass
