from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

from openpyxl.compat import iteritems, OrderedDict

from .rules import CellIsRule, ColorScaleRule, FormatRule


class ConditionalFormatting(object):
    """Conditional formatting rules."""
    rule_attributes = ('aboveAverage', 'bottom', 'dxfId', 'equalAverage', 'operator', 'percent', 'priority', 'rank',
                       'stdDev', 'stopIfTrue', 'text')
    icon_attributes = ('iconSet', 'showValue', 'reverse')

    def __init__(self):
        self.cf_rules = OrderedDict()
        self.max_priority = 0
        self.parse_rules = {}

    def add(self, range_string, cfRule):
        """Add a rule.  Rule is either:
         1. A dictionary containing a key called type, and other keys, as in `ConditionalFormatting.rule_attributes`.
         2. A rule object, such as ColorScaleRule, FormulaRule or CellIsRule

         The priority will be added automatically.
        """
        if isinstance(cfRule, dict):
            rule = cfRule
        else:
            rule = cfRule.rule
        rule['priority'] = self.max_priority + 1
        self.max_priority += 1
        if range_string not in self.cf_rules:
            self.cf_rules[range_string] = []
        self.cf_rules[range_string].append(rule)

    def update(self, cfRules):
        """Set the conditional formatting rules from a dictionary.  Intended for use when loading a document.
        cfRules use the structure: {range_string: [rule1, rule2]}, eg:
        {'A1:A4': [{'type': 'colorScale', 'priority': 13, 'colorScale': {'cfvo': [{'type': 'min'}, {'type': 'max'}],
        'color': [Color('FFFF7128'), Color('FFFFEF9C')]}]}
        """
        for range_string, rules in iteritems(cfRules):
            if range_string not in self.cf_rules:
                self.cf_rules[range_string] = rules
            else:
                self.cf_rules[range_string] += rules

        # Fix any gap in the priority range.
        self.max_priority = 0
        priorityMap = []
        for range_string, rules in iteritems(self.cf_rules):
            for rule in rules:
                priorityMap.append(rule['priority'])
        priorityMap.sort()
        for range_string, rules in iteritems(self.cf_rules):
            for rule in rules:
                priority = priorityMap.index(rule['priority']) + 1
                rule['priority'] = priority
                if 'priority' in rule and priority > self.max_priority:
                    self.max_priority = priority

    def setDxfStyles(self, wb):
        """Formatting for non color scale conditional formatting uses the dxf style list in styles.xml. This scans
        the cf_rules for dxf styles which have not been added - and saves them to the workbook.

        When adding a conditional formatting rule that uses a font, border or fill, this must be called at least once
        before saving the workbook.

        :param wb: the workbook
        """
        if not wb.style_properties:
            wb.style_properties = {'dxf_list': []}
        elif 'dxf_list' not in wb.style_properties:
            wb.style_properties['dxf_list'] = []

        for rules in self.cf_rules.values():
            for rule in rules:
                if 'dxf' in rule:
                    dxf = {}
                    if 'font' in rule['dxf'] and isinstance(rule['dxf']['font'], Font):
                        # DXF font is limited to color, bold, italic, underline and strikethrough
                        dxf['font'] = rule['dxf']['font']
                    if 'border' in rule['dxf'] and isinstance(rule['dxf']['border'], Borders):
                        dxf['border'] = rule['dxf']['border']
                    if 'fill' in rule['dxf'] and isinstance(rule['dxf']['fill'], Fill):
                        dxf['fill'] = rule['dxf']['fill']

                    wb.style_properties['dxf_list'].append(dxf)
                    rule.pop('dxf')
                    rule['dxfId'] = len(wb.style_properties['dxf_list']) - 1

