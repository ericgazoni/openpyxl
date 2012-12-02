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
    ws = book.get_active_sheet()
    ws.cell('A1').value = 'You should see a logo below'

    # create an image instance
    img = Image('logo.png')

    # place it if required
    img.drawing.left = 200
    img.drawing.top = 100

    # add to worksheet
    ws.add_image(img)
    wb.save('logo.xlsx')
