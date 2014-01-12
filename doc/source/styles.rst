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
    ws.conditional_formatting.add2ColorScale('A1:A10', 'min', None, 'FFAA0000', 'max', None, 'FF00AA00')

    # Add a three-color scale
    ws.conditional_formatting.add3ColorScale('B1:B10', 'percentile', 10, 'FFAA0000'
                                                       'percentile', 50, 'FF0000AA', 'percentile', 90, 'FF00AA00')

    # Add a conditional formatting based on a cell comparison
    # addCellIs(range_string, operator, formula, stopIfTrue, wb, font, border, fill)
    # Format if cell is less than 'formula'
    ws.conditional_formatting.addCellIs('C2:C10', 'lessThan', ['C$1'], True, wb, None, None, redFill)
    # Format if cell is between 'formula'
    ws.conditional_formatting.addCellIs('D2:D10', 'between', ['1','5'], True, wb, None, None, redFill)

    # Custom formatting
    # There are many types of conditional formatting - it's possible to add additional types, through addCustomRule
    dxfId = ws.conditional_formatting.addDxfStyle(wb, None, None, None)
    ws.conditional_formatting.addCustomRule('E1:E10',  {'type': 'expression', 'dxfId': dxfId,
                                                        'formula': ['ISBLANK(E1)'], 'stopIfTrue': '1'})

    # Check
    wb.save("test.xlsx")


