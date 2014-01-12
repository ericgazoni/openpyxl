Optimized reader
================

Sometimes, you will need to open or write extremely large XLSX files,
and the common routines in openpyxl won't be able to handle that load.
Hopefully, there are two modes that enable you to read and write unlimited
amounts of data with (near) constant memory consumption.

Introducing :class:`openpyxl.reader.iter_worksheet.IterableWorksheet`::

    from openpyxl import load_workbook
    wb = load_workbook(filename = 'large_file.xlsx', use_iterators = True)
    ws = wb.get_sheet_by_name(name = 'big_data') # ws is now an IterableWorksheet

    for row in ws.iter_rows(): # it brings a new method: iter_rows()

        for cell in row:

            print cell.internal_value

.. warning::

    * As you can see, we are using cell.internal_value instead of .value.
    * :class:`openpyxl.reader.iter_worksheet.IterableWorksheet` are read-only
    * cell, range, rows, columns methods and properties are disabled

Cells returned by iter_rows() are not regular :class:`openpyxl.cell.Cell` but
:class:`openpyxl.reader.iter_worksheet.RawCell`.

Optimized writer
================

Here again, the regular :class:`openpyxl.worksheet.Worksheet` has been replaced
by a faster alternative, the :class:`openpyxl.writer.dump_worksheet.DumpWorksheet`.
When you want to dump large amounts of data, you might find optimized writer helpful::

    from openpyxl import Workbook
    wb = Workbook(optimized_write = True)

    ws = wb.create_sheet()

    # now we'll fill it with 10k rows x 200 columns
    for irow in xrange(10000):
        ws.append(['%d' % i for i in xrange(200)])

    wb.save('new_big_file.xlsx') # don't forget to save !

.. warning::

    * Those worksheet only have an append() method, it's not possible to access independent cells directly (through cell() or range()). They are write-only.
    * It is able to export unlimited amount of data (even more than Excel can handle actually), while keeping memory usage under 10Mb.
    * A workbook using the optimized writer can only be saved once. After that, every attempt to save the workbook or append() to an existing worksheet will raise an :class:`openpyxl.shared.exc.WorkbookAlreadySaved` exception.


