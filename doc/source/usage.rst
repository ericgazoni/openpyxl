Simple usage
=======================

Write a workbook
------------------
::

    from openpyxl import Workbook

    from openpyxl.cell import get_column_letter

    wb = Workbook()

    dest_filename = r'empty_book.xlsx'

    ws = wb.active

    ws.title = "range names"

    for col_idx in xrange(1, 40):
        col = get_column_letter(col_idx)
        for row in xrange(1, 600):
            ws.cell('%s%s'%(col, row)).value = '%s%s' % (col, row)

    ws = wb.create_sheet()

    ws.title = 'Pi'

    ws['F5'] = 3.14

    wb.save(filename = dest_filename)


Read an existing workbook
-------------------------
::

    from openpyxl import load_workbook

    wb = load_workbook(filename = r'empty_book.xlsx')

    sheet_ranges = wb['range names']

    print sheet_ranges['D18'].value # D18


.. note ::

    There are several flags that can be used in load_workbook.

    - `guess_types` will enable (default) or disable type inference when
      reading cells.

    - `data_only` controls whether cells with formulae have either the
      formula (default) or the value stored the last time Excel read the sheet.

    - `keep_vba` controls whether any Visual Basic elements are preserved or
      not (default). If they are preserved they are still not editable.


.. warning ::

    openpyxl does currently not read all possible items in an Excel file so
    images and charts will be lost from existing files if they are opened and
    saved with the same name.


Using number formats
--------------------
::

    import datetime
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    # set date using a Python datetime
    ws['A1'] = datetime.datetime(2010, 7, 21)

    print ws['A1'].style.number_format.format_code # returns 'yyyy-mm-dd'

    # set percentage using a string followed by the percent sign
    ws['B1'] = '3.14%'

    print ws['B1'].value # returns 0.031400000000000004

    print ws['B1'].style.number_format.format_code # returns '0%'


Using formulae
--------------
::

    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active

    # add a simple formula
    ws["A1"] = "=SUM(1, 1)"
    wb.save("formula.xlsx")

.. warning::
    NB function arguments *must* be separated by commas and not other
    punctuation such as semi-colons



Inserting an image
-------------------
::

    from openpyxl import Workbook
    from openpyxl.drawing import Image

    wb = Workbook()
    ws = wb.active
    ws['A1'] = 'You should see three logos below'
    ws['A2'] = 'Resize the rows and cells to see anchor differences'

    # create image instances
    img = Image('logo.png')
    img2 = Image('logo.png')
    img3 = Image('logo.png')

    # place image relative to top left corner of spreadsheet
    img.drawing.top = 100
    img.drawing.left = 150

    # the top left offset needed to put the image
    # at a specific cell can be automatically calculated
    img2.anchor(ws['D12'])

    # one can also position the image relative to the specified cell
    # this can be advantageous if the spreadsheet is later resized
    # (this might not work as expected in LibreOffice)
    img3.anchor(ws['G20'], anchortype='oneCell')

    # afterwards one can still add additional offsets from the cell
    img3.drawing.left = 5
    img3.drawing.top = 5

    # add to worksheet
    ws.add_image(img)
    ws.add_image(img2)
    ws.add_image(img3)
    wb.save('logo.xlsx')


Validating cells
----------------
::

    from openpyxl import Workbook
    from openpyxl.datavalidation import DataValidation, ValidationType

    # Create the workbook and worksheet we'll be working with
    wb = Workbook()
    ws = wb.active

    # Create a data-validation object with list validation
    dv = DataValidation(ValidationType.LIST, formula1='"Dog,Cat,Bat"', allow_blank=True)

    # Optionally set a custom error message
    dv.set_error_message('Your entry is not in the list', 'Invalid Entry')

    # Optionally set a custom prompt message
    dv.set_prompt_message('Please select from the list', 'List Selection')

    # Add the data-validation object to the worksheet
    ws.add_data_validation(dv)

    # Create some cells, and add them to the data-validation object
    c1 = ws["A1"]
    c1.value = "Dog"
    dv.add_cell(c1)
    c2 = ws["A2"]
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

.. note::
    See http://www.contextures.com/xlDataVal07.html for custom rules

