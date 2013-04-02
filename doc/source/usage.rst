Simple usage
=======================

Write a workbook
------------------
::

    from openpyxl import Workbook

    from openpyxl.cell import get_column_letter

    wb = Workbook()

    dest_filename = r'empty_book.xlsx'

    ws = wb.worksheets[0]

    ws.title = "range names"

    for col_idx in xrange(1, 40):
        col = get_column_letter(col_idx)
        for row in xrange(1, 600):
            ws.cell('%s%s'%(col, row)).value = '%s%s' % (col, row)

    ws = wb.create_sheet()

    ws.title = 'Pi'

    ws.cell('F5').value = 3.14

    wb.save(filename = dest_filename)

Read an existing workbook
-----------------------------
::

    from openpyxl import load_workbook

    wb = load_workbook(filename = r'empty_book.xlsx')

    sheet_ranges = wb.get_sheet_by_name(name = 'range names')

    print sheet_ranges.cell('D18').value # D18


Using number formats
----------------------
::

    import datetime
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.worksheets[0]

    # set date using a Python datetime
    ws.cell('A1').value = datetime.datetime(2010, 7, 21)

    print ws.cell('A1').style.number_format.format_code # returns 'yyyy-mm-dd'

    # set percentage using a string followed by the percent sign
    ws.cell('B1').value = '3.14%'

    print ws.cell('B1').value # returns 0.031400000000000004

    print ws.cell('B1').style.number_format.format_code # returns '0%'


Inserting an image
-------------------
::

    from openpyxl import Workbook
    from openpyxl.drawing import Image

    wb = Workbook()
    ws = wb.get_active_sheet()
    ws.cell('A1').value = 'You should see a logo below'

    # create an image instance
    img = Image('logo.png')

    # place it if required
    img.drawing.left = 200
    img.drawing.top = 100

    # you could also 'anchor' the image to a specific cell
    # img.anchor(ws.cell('B12'))

    # add to worksheet
    ws.add_image(img)
    wb.save('logo.xlsx')


Validating cells
----------------
::

    from openpyxl import Workbook
    from openpyxl.datavalidation import DataValidation, ValidationType

    # Create the workbook and worksheet we'll be working with
    wb = Workbook()
    ws = wb.get_active_sheet()

    # Create a data-validation object with list validation
    dv = DataValidation(ValidationType.LIST, '"Dog,Cat,Bat"', allow_blank=True)

    # Optionally set a custom error message
    dv.set_error_message('Your entry is not in the list', 'Invalid Entry')

    # Optionally set a custom prompt message
    dv.set_prompt_message('Please select from the list', 'List Selection')

    # Add the data-validation object to the worksheet
    ws.add_data_validation(dv)

    # Create some cells, and add them to the data-validation object
    c1 = ws.cell("A1")
    c1.value = "Dog"
    dv.add_cell(c1)
    c2 = ws.cell("A2")
    c2.value = "An invalid value"
    dv.add_cell(c2)

    # Or, apply the validation to a range of cells
    dv.ranges.append('B1:B1048576')

    # Write the sheet out.  If you now open the sheet in Excel, you'll find that
    # the cells have data-validation applied.
    wb.save("test.xlsx")


Other validation examples
-------------------------

Any whole number:
::
    dv = DataValidation(ValidationType.WHOLE)

Any whole number above 100:
::
    dv = DataValidation(ValidationType.WHOLE,
                        ValidationOperator.GREATER_THAN,
                        100)

Any decimal number:
::
    dv = DataValidation(ValidationType.DECIMAL)

Any decimal number between 0 and 1:
::
    dv = DataValidation(ValidationType.DECIMAL,
                        ValidationOperator.BETWEEN,
                        0, 1)

Any date:
::
    dv = DataValidation(ValidationType.DATE)

or time:
::
    dv = DataValidation(ValidationType.TIME)

Any string at most 15 characters:
::
    dv = DataValidation(ValidationType.TEXT_LENGTH,
                        ValidationOperator.LESS_THAN_OR_EQUAL,
                        15)

Custom rule:
::
    dv = DataValidation(ValidationType.CUSTOM,
                        None,
                        "=SOMEFORMULA")

See http://www.contextures.com/xlDataVal07.html for custom rules

