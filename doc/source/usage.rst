Simple usage
=======================

Write a workbook 
------------------
::

    from openpyxl.workbook import Workbook
    from openpyxl.writer.excel import ExcelWriter
    
    from openpyxl.cell import get_column_letter
    
    wb = Workbook()
    
    ew = ExcelWriter(workbook = wb)
    
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
    
    ew.save(filename = dest_filename)
    
Read an existing workbook 
-----------------------------
::

    from openpyxl.reader.excel import load_workbook

    wb = load_workbook(filename = r'empty_book.xlsx')
    
    sheet_ranges = wb.get_sheet_by_name(name = 'range names')
    
    print sheet_ranges.cell('D18').value # D18


Using number formats
----------------------
::

    import datetime
    from openpyxl.workbook import Workbook
    
    wb = Workbook()
    ws = wb.worksheets[0]
    
    # set date using a Python datetime
    ws.cell('A1').value = datetime.datetime(2010, 7, 21)
    
    print ws.cell('A1').style.number_format.format_code # returns 'yyyy-mm-dd'
    
    # set percentage using a string followed by the percent sign
    ws.cell('B1').value = '3.14%'
    
    print ws.cell('B1').value # returns 0.031400000000000004
    
    print ws.cell('B1').style.number_format.format_code # returns '0%'