Working with styles
===================

Basic Font Colors
-----------------
::

    from openpyxl.workbook import Workbook
    from openpyxl.style import Color

    wb = Workbook()
    ws = wb.active
    ws['A1'] = 'This is red'
    ws['A1'].style.font.color.index = Color.RED


Edit Print Settings
-------------------
::

    from openpyxl.workbook import Workbook

    wb = Workbook()
    ws = wb.active

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_TABLOID
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToHeight = 0
    ws.page_setup.fitToWidth = 1
    ws.page_setup.horizontalCentered = True
    ws.page_setup.verticalCentered = True


Merge / Unmerge cells
---------------------
::

    from openpyxl.workbook import Workbook

    wb = Workbook()
    ws = wb.active

    ws.merge_cells('A1:B1')
    ws.unmerge_cells('A1:B1')

    # or
    ws.merge_cells(start_row=2,start_col=1,end_row=2,end_col=4)
    ws.unmerge_cells(start_row=2,start_col=1,end_row=2,end_col=4)


Header / Footer
---------------
::

    from openpyxl.workbook import Workbook

    wb = Workbook()
    ws = wb.worksheets[0]

    ws.header_footer.center_header.text = 'My Excel Page'
    ws.header_footer.center_header.font_size = 14
    ws.header_footer.center_header.font_name = "Tahoma,Bold"
    ws.header_footer.center_header.font_color = "CC3366"

    # Or just
    ws.header_footer.right_footer.text = 'My Right Footer'


Conditional Formatting
----------------------

There are many types of conditional formatting - below are some examples for setting this within an excel file.

::

    from openpyxl import Workbook
    from openpyxl.style import Color, Fill
    from openpyxl.styles.formatting import ColorScaleRule, CellIsRule, FormulaRule
    wb = Workbook()
    ws = wb.active

    # Create fill
    redFill = Fill()
    redFill.start_color.index = 'FFEE1111'
    redFill.end_color.index = 'FFEE1111'
    redFill.fill_type = Fill.FILL_SOLID

    # Add a two-color scale
    # add2ColorScale(range_string, start_type, start_value, start_rgb, end_type, end_value, end_rgb)
    # Takes colors in excel 'FFRRGGBB' style.
    ws.conditional_formatting.add('A1:A10', ColorScaleRule(start_type='min', start_rgb='FFAA0000',
                                                           end_type='max', end_rgb='FF00AA00'))

    # Add a three-color scale
    ws.conditional_formatting.add('B1:B10', ColorScaleRule(start_type='percentile', start_value=10, start_rgb='FFAA0000',
                                                           mid_type='percentile', mid_value=50, mid_rgb='FF0000AA',
                                                           end_type='percentile', end_value=90, end_rgb='FF00AA00')

    # Add a conditional formatting based on a cell comparison
    # addCellIs(range_string, operator, formula, stopIfTrue, wb, font, border, fill)
    # Format if cell is less than 'formula'
    ws.conditional_formatting.addCellIs('C2:C10', CellIsRule(operator='lessThan', formula=['C$1'], stopIfTrue=True,
                                                             fill=redFill)

    # Format if cell is between 'formula'
    ws.conditional_formatting.addCellIs('D2:D10', CellIsRule(operator='between', formula=['1','5'], stopIfTrue=True,
                                                             fill=redFill)

    # Format using a formula
    ws.conditional_formatting.add('E1:E10',  FormulaRule(formula=['ISBLANK(E1)'], stopIfTrue=True, fill=redFill})

    # Aside from the 2-color and 3-color scales, format rules take fonts, borders and fills for styling:
    myFont = Font()
    myBorder = Borders()
    ws.conditional_formatting.add('E1:E10',  FormulaRule(formula=['E1=0'], font=myFont, border=myBorder, fill=redFill})

    # Custom formatting
    # There are many types of conditional formatting - it's possible to add additional types directly:
    ws.conditional_formatting.add('E1:E10',  {'type': 'expression', 'dxf': {'fill': redFill},
                                              'formula': ['ISBLANK(E1)'], 'stopIfTrue': '1'})

    # Before writing, call setDxfStyles before saving when adding a conditional format that has a font/border/fill
    ws.conditional_formatting.setDxfStyles(wb)
    wb.save("test.xlsx")


