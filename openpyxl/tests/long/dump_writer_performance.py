import openpyxl
import tempfile


def test_large_append():
    wb = openpyxl.Workbook(optimized_write=True)
    ws = wb.create_sheet()
    row = ('this is some text', 3.14)
    total_rows = int(2e4)
    for idx in xrange(total_rows):
        if not idx % 10000:
            print "%.2f%%" % (100 * (float(idx) / float(total_rows)))
        ws.append(row)
    wb.save(tempfile.TemporaryFile(mode='wb'))
