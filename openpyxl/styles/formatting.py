from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl
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

from collections import Mapping

from openpyxl.compat import iteritems, OrderedDict
from .colors import Color
from openpyxl.styles import Font


class FormatRule(Mapping):
    """Utility dictionary for formatting rules with specified keys only"""

    __slots__ = ('aboveAverage', 'bottom', 'dxfId', 'equalAverage',
                 'operator', 'percent', 'priority', 'rank', 'stdDev', 'stopIfTrue',
                 'text')

    def update(self, dictionary):
        for k, v in iteritems(dictionary):
            self[k] = v

    def __getitem__(self, key):
        if key not in self.__slots__:
            raise KeyError("{0} is not a valid key for a formatting rule".format(key))
        return getattr(self, key)

    def __setitem__(self, key, value):
        if key not in self.__slots__:
            raise KeyError("{0} is not a valid key for a formatting rule".format(key))
        setattr(self, key, value)

    def __iter__(self):
        return iter(self.__slots__)

    def iterkeys(self):
        for key in self.__slots__:
            if getattr(self, key, None) is not None:
                yield key
            continue

    def keys(self):
        return list(self.iterkeys())

    def itervalues(self):
        for key in self.iterkeys():
            yield self[key]

    def values(self):
        return list(self.itervalues())

    def iteritems(self):
        for key in self.iterkeys():
            yield key, getattr(self, key)

    def items(self):
        return [(key, value) for key, value in self.iteritems()]

    def __len__(self):
        return len(self.keys())


class ColorScale(object):

    __slots__ = (
        "start_type", "_start_value", "start_color",
        "mid_type", "_mid_value", "mid_color",
        "end_type", "_end_value", "end_color"
    )

    valid_types = ('min', 'max', 'num', 'percent', 'percentile', 'formula')

    def __init__(self,
                 start_type=None,
                 start_value=None,
                 start_color=None,
                 mid_type=None,
                 mid_value=None,
                 mid_color=None,
                 end_type=None,
                 end_value=None,
                 end_color=None
                 ):
        self.start_type = start_type
        self.start_value = start_value
        self.start_color = start_color
        self.mid_type = mid_type
        self.mid_value = mid_value
        self.mid_color = mid_color
        self.end_type = end_type
        self.end_value = end_value
        self.end_color = end_color

    @property
    def start_value(self):
        return self._start_value

    @start_value.setter
    def start_value(self, value):
        if value is not None and self.start_type in ('min', 'max'):
            raise ValueError("ColorScale with min or max cannot have values")
        self._start_value = value

    @property
    def end_value(self):
        return self._end_value

    @end_value.setter
    def end_value(self, value):
        if value is not None and self.end_type in ('min', 'max'):
            raise ValueError("ColorScale with min or max cannot have values")
        self._end_value = value

    @property
    def mid_value(self):
        return self._mid_value

    @mid_value.setter
    def mid_value(self, value):
        if value is not None and self.mid_type in ('min', 'max'):
            raise ValueError("ColorScale with min or max cannot have values")
        self._mid_value = value

    @property
    def cfvo(self):
        """Return a dictionary representation"""
        vals = []
        for attr in 'start', 'mid', 'end':
            typ = getattr(self, attr + '_type')
            if typ is None:
                continue
            d = {'type': typ}
            v = getattr(self, attr + '_value')
            if v is not None:
                d['val'] = str(v)
            vals.append(d)
        return vals

    @property
    def colors(self):
        """Return start, mid and end colours"""
        return [v for v in (self.start_color, self.mid_color, self.end_color) if v is not None]


class ConditionalFormatting(object):
    """Conditional formatting rules."""
    rule_attributes = ('aboveAverage', 'bottom', 'dxfId', 'equalAverage', 'operator', 'percent', 'priority', 'rank',
                       'stdDev', 'stopIfTrue', 'text')
    icon_attributes = ('iconSet', 'showValue', 'reverse')

    def __init__(self):
        self.cf_rules = OrderedDict()
        self.max_priority = 0

    def update(self, cfRules):
        """Set the conditional formatting rules from a dictionary.  Intended for use when loading a document.
        cfRules use the structure: {range_string: [rule1, rule2]}, eg:
        {'A1:A4': [{'type': 'colorScale', 'priority': '13', 'colorScale': {'cfvo': [{'type': 'min'}, {'type': 'max'}],
        'color': [Color('FFFF7128'), Color('FFFFEF9C')]}]}
        """
        for range_string, rules in iteritems(cfRules):
            if range_string not in self.cf_rules:
                self.cf_rules[range_string] = []
            self.cf_rules[range_string] += rules
            for rule in rules:
                if int(rule['priority']) > self.max_priority:
                    self.max_priority = int(rule['priority'])

    def fixPriority(self):
        """Fixes any gap in the priority range before writing to the worksheet."""
        self.max_priority = 0
        priorityMap = []
        for range_string, rules in iteritems(self.cf_rules):
            for rule in rules:
                priorityMap.append(int(rule['priority']))
        priorityMap.sort()
        for range_string, rules in iteritems(self.cf_rules):
            for rule in rules:
                priority = priorityMap.index(int(rule['priority'])) + 1
                rule['priority'] = str(priority)
                if 'priority' in rule and priority > self.max_priority:
                    self.max_priority = priority

    def addDxfStyle(self, wb, font, border, fill):
        """Formatting for non color scale conditional formatting uses the dxf style list in styles.xml.  Add a style
        and get the corresponding style id to use in the conditional formatting rule.

        Excel adds a dxf style for each conditional formatting, even if it already exists.

        :param wb: the workbook
        :param font: openpyxl.style.Font
        :param border: openpyxl.style.Border
        :param fill: openpyxl.style.Fill
        :return: dxfId (excel uses a 0 based index for the dxfId)
        """
        if not wb.style_properties:
            wb.style_properties = {'dxf_list': []}
        elif 'dxf_list' not in wb.style_properties:
            wb.style_properties['dxf_list'] = []

        dxf = {}
        if font and isinstance(font, Font):
            # DXF font is limited to color, bold, italic, underline and strikethrough
            dxf['font'] = font
        if border:
            dxf['border'] = [border]
        if fill:
            dxf['fill'] = [fill]

        wb.style_properties['dxf_list'].append(dxf)
        return len(wb.style_properties['dxf_list']) - 1

    def addCustomRule(self, range_string, rule):
        """Add a custom rule.  Rule is a dictionary containing a key called type, and other keys, as found in
        `ConditionalFormatting.rule_attributes`.  The priority will be added automatically.

        For example:
        {'type': 'colorScale', 'colorScale': {'cfvo': [{'type': 'min'}, {'type': 'max'}],
                                              'color': [Color('FFFF7128'), Color('FFFFEF9C')]}
        """
        rule['priority'] = self.max_priority + 1
        self.max_priority += 1
        if range_string not in self.cf_rules:
            self.cf_rules[range_string] = []
        self.cf_rules[range_string].append(rule)

    def add2ColorScale(self, range_string, start_type, start_value, start_rgb, end_type, end_value, end_rgb):
        """
        Add a 2-color scale to the conditional formatting.

        :param range_string: Range of the conditional formatting, eg "B1:B10" or "A1:A1048576" for the whole column.
        :param start_type: Starting color reference - can be: num, percent, percentile, min, max, formula
        :param start_value: Starting value.  Percent expressed in integer from 0 - 100. (Ignored for min / max.)
        :param start_rgb: Start RGB color, such as 'FFAABB11'
        :param end_type: Ending color reference - can be: num, percent, percentile, min, max, formula
        :param end_value: Ending value.
        :param end_rgb: End RGB color, such as 'FFAABB11'
        """
        cs = ColorScale()
        cs.start_type = start_type
        cs.start_value = start_value
        cs.start_color = start_rgb
        cs.end_type = end_type
        cs.end_value = end_value
        cs.end_color = end_rgb
        rule ={'type': 'colorScale', 'colorScale': {'color': cs.colors, 'cfvo': cs.cfvo}}

        self.addCustomRule(range_string, rule)

    def add3ColorScale(self, range_string, start_type, start_value, start_rgb, mid_type, mid_value, mid_rgb, end_type,
                       end_value, end_rgb):
        """Add a 3-color scale to the conditional formatting.  See `add2ColorScale` for parameter descriptions."""
        cs = ColorScale(start_type=start_type, start_value=start_value, start_color=start_rgb,
                       mid_type=mid_type, mid_value=mid_value, mid_color=mid_rgb,
                       end_type=end_type, end_value=end_value, end_color=end_rgb)
        rule ={'type': 'colorScale', 'colorScale': {'color': cs.colors, 'cfvo': cs.cfvo}}
        self.addCustomRule(range_string, rule)

    def addCellIs(self, range_string, operator, formula, stopIfTrue, wb, font, border, fill):
        """Add a conditional formatting of type cellIs.

        Formula is in a list to handle multiple formula's, such as ['a1']

        Valid values for operator are:
        'between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'
        """
        # Excel doesn't use >, >=, etc, but allow for ease of python development
        expand = {">": "greaterThan", ">=": "greaterThanOrEqual", "<": "lessThan", "<=": "lessThanOrEqual",
                  "=": "equal", "==": "equal", "!=": "notEqual"}
        operator = expand.get(operator, operator)

        dxfId = self.addDxfStyle(wb, font, border, fill)
        rule = {'type': 'cellIs', 'dxfId': dxfId, 'operator': operator, 'formula': formula}
        if stopIfTrue:
            rule['stopIfTrue'] = '1'
        self.addCustomRule(range_string, rule)
