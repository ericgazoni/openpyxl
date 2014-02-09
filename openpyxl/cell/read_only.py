from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl


from openpyxl.compat import unicode
from openpyxl.date_time import from_excel
from openpyxl.styles import is_date_format
from .cell import Cell


class ReadOnlyCell(object):

    __slots__ = ('row', 'column', '_value', 'data_type', '_style_id')


    def __init__(self, row, column, value, data_type, style_id=None):
        self.row = row
        self.column = column
        self.data_type = data_type
        self.style_id = style_id
        self.value = value

    def __eq__(self, other):
        for a in self.__slots__:
            if getattr(self, a) != getattr(other, a):
                return False
        return True

    @classmethod
    def set_string_table(cls, string_table):
        cls.string_table = string_table

    @classmethod
    def set_style_table(cls, style_table):
        cls.style_table = style_table

    @classmethod
    def set_base_date(cls, base_date):
        cls.base_date = base_date

    @property
    def coordinate(self):
        return "{1}{0}".format(self.row, self.column)

    @property
    def is_date(self):
        return self.data_type == Cell.TYPE_NUMERIC and is_date_format(self.number_format)

    @property
    def number_format(self):
        if self.style_id is None:
            return
        style = self.style_table[self._style_id]
        return style.number_format.format_code

    @property
    def style_id(self):
        return self._style_id

    @style_id.setter
    def style_id(self, value):
        if hasattr(self, '_style_id'):
            raise AttributeError("Cell is read only")
        if value is not None:
            value = int(value)
        self._style_id = value

    @property
    def internal_value(self):
        return self._value

    @property
    def value(self):
        if self._value is None:
            return
        if self.data_type == Cell.TYPE_BOOL:
            return self._value == '1'
        elif self.is_date:
            return from_excel(self._value, self.base_date)
        elif self.data_type in(Cell.TYPE_INLINE, Cell.TYPE_FORMULA_CACHE_STRING):
            return unicode(self._value)
        elif self.data_type == Cell.TYPE_STRING:
            return unicode(self.string_table[int(self._value)])
        return self._value

    @value.setter
    def value(self, value):
        if hasattr(self, '_value'):
            raise AttributeError("Cell is read only")
        if value is None:
            self.data_type = Cell.TYPE_NULL
        elif self.data_type == Cell.TYPE_NUMERIC:
            value = float(value)
        self._value = value

EMPTY_CELL = ReadOnlyCell(None, None, None, Cell.TYPE_NULL, None)
