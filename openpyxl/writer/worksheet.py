# file openpyxl/writer/worksheet.py

"""Write worksheets to xml representations."""

# Python stdlib imports
from StringIO import StringIO  # cStringIO doesn't handle unicode

# package imports
from openpyxl.cell import column_index_from_string
from openpyxl.shared.xmltools import Element, SubElement, XMLGenerator, \
        get_document_content, start_tag, end_tag, tag


def row_sort(cell):
    """Translate column names for sorting."""
    return column_index_from_string(cell.column)


def write_worksheet(worksheet, string_table, style_table):
    """Write a worksheet to an xml file."""
    xml_file = StringIO()
    doc = XMLGenerator(xml_file, 'utf-8')
    start_tag(doc, 'worksheet',
            {'xml:space': 'preserve',
            'xmlns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'})
    start_tag(doc, 'sheetPr')
    tag(doc, 'outlinePr',
            {'summaryBelow': '%d' % (worksheet.show_summary_below),
            'summaryRight': '%d' % (worksheet.show_summary_right)})
    end_tag(doc, 'sheetPr')
    tag(doc, 'dimension', {'ref': '%s' % worksheet.calculate_dimension()})
    start_tag(doc, 'sheetViews')
    start_tag(doc, 'sheetView', {'workbookViewId': '0'})
    tag(doc, 'selection', {'activeCell': worksheet.active_cell,
            'sqref': worksheet.selected_cell})
    end_tag(doc, 'sheetView')
    end_tag(doc, 'sheetViews')
    tag(doc, 'sheetFormatPr', {'defaultRowHeight': '15'})
    write_worksheet_cols(doc, worksheet)
    write_worksheet_data(doc, worksheet, string_table, style_table)
    write_worksheet_hyperlinks(doc, worksheet)
    end_tag(doc, 'worksheet')
    doc.endDocument()
    xml_string = xml_file.getvalue()
    xml_file.close()
    return xml_string


def write_worksheet_cols(doc, worksheet):
    """Write worksheet columns to xml."""
    if worksheet.column_dimensions:
        start_tag(doc, 'cols')
        for column_string, columndimension in \
                worksheet.column_dimensions.iteritems():
            col_index = column_index_from_string(column_string)
            col_def = {}
            col_def['collapsed'] = str(columndimension.style_index)
            col_def['min'] = str(col_index)
            col_def['max'] = str(col_index)
            if columndimension.width != \
                    worksheet.default_column_dimension.width:
                col_def['customWidth'] = 'true'
            if not columndimension.visible:
                col_def['hidden'] = 'true'
            if columndimension.outline_level > 0:
                col_def['outlineLevel'] = str(columndimension.outline_level)
            if columndimension.collapsed:
                col_def['collapsed'] = 'true'
            if columndimension.auto_size:
                col_def['bestFit'] = 'true'
            if columndimension.width > 0:
                col_def['width'] = str(columndimension.width)
            else:
                col_def['width'] = '9.10'
            tag(doc, 'col', col_def)
        end_tag(doc, 'cols')


def write_worksheet_data(doc, worksheet, string_table, style_table):
    """Write worksheet data to xml."""
    start_tag(doc, 'sheetData')
    max_column = worksheet.get_highest_column()
    style_id_by_hash = style_table
    cells_by_row = {}
    for cell in worksheet.get_cell_collection():
        cells_by_row.setdefault(cell.row, []).append(cell)
    for row_idx in sorted(cells_by_row):
        row_dimension = worksheet.row_dimensions[row_idx]
        attrs = {'r': '%d' % row_idx,
                 'spans': '1:%d' % max_column}
        if row_dimension.height > 0:
            attrs['ht'] = str(row_dimension.height)
            attrs['customHeight'] = '1'
        start_tag(doc, 'row', attrs)
        row_cells = cells_by_row[row_idx]
        sorted_cells = sorted(row_cells, key=row_sort)
        for cell in sorted_cells:
            value = cell._value
            coordinate = cell.get_coordinate()
            attributes = {'r': coordinate}
            attributes['t'] = cell.data_type
            if coordinate in worksheet._styles:
                attributes['s'] = '%d' % style_id_by_hash[
                        worksheet._styles[coordinate].__crc__()]
            start_tag(doc, 'c', attributes)
            if cell.data_type == cell.TYPE_STRING:
                tag(doc, 'v', body='%s' % string_table[value])
            elif cell.data_type == cell.TYPE_FORMULA:
                tag(doc, 'f', body='%s' % value[1:])
                tag(doc, 'v')
            elif cell.data_type == cell.TYPE_NUMERIC:
                tag(doc, 'v', body='%s' % value)
            else:
                tag(doc, 'v', body='%s' % value)
            end_tag(doc, 'c')
        end_tag(doc, 'row')
    end_tag(doc, 'sheetData')


def write_worksheet_hyperlinks(doc, worksheet):
    """Write worksheet hyperlinks to xml."""
    write_hyperlinks = False
    for cell in worksheet.get_cell_collection():
        if cell.hyperlink_rel_id is not None:
            write_hyperlinks = True
            break
    if write_hyperlinks:
        start_tag(doc, 'hyperlinks')
        for cell in worksheet.get_cell_collection():
            if cell.hyperlink_rel_id is not None:
                attrs = {'display': cell.hyperlink,
                        'ref': cell.get_coordinate(),
                        'r:id': cell.hyperlink_rel_id}
                tag(doc, 'hyperlink', attrs)
        end_tag(doc, 'hyperlinks')


def write_worksheet_rels(worksheet):
    """Write relationships for the worksheet to xml."""
    root = Element('Relationships', {'xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships'})
    if worksheet.relationships:
        for rel in worksheet.relationships:
            attrs = {'Id': rel.id, 'Type': rel.type, 'Target': rel.target}
            if rel.target_mode:
                attrs['TargetMode'] = rel.target_mode
            SubElement(root, 'Relationship', attrs)
    return get_document_content(root)
