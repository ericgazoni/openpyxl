import os
import sys
import timeit

import openpyxl


def writer(optimised, cols, rows):
    """
    Create a worksheet with variable width rows. Because data must be
    serialised row by row it is often the width of the rows which is most
    important.
    """
    wb = openpyxl.Workbook(optimized_write=optimised)
    ws = wb.create_sheet()
    row = range(rows)
    for idx in xrange(rows):
        if not (idx + 1) % rows/10:
            progress = "." * ((idx + 1) / (1 + rows/10))
            sys.stdout.write("\r" + progress)
            sys.stdout.flush()
        ws.append(row)
    folder = os.path.split(__file__)[0]
    print
    wb.save(os.path.join(folder, "files", "large.xlsx"))


def timer(fn, **kw):
    """
    Create a timeit call to a function and pass in keyword arguments.
    The function is called twice, once using the standard workbook, then with the optimised one.
    Time from the best of three is taken.
    """
    result = []
    cols = kw.get("cols", 0)
    rows = kw.get("rows", 0)
    for opt in (False, True):
        kw.update(optimised=opt)
        print "{} cols {} rows, Worksheet is {}".format(cols, rows,
                                                        opt and "optimised" or "not optimised")
        times = timeit.repeat("{}(**{})".format(fn.func_name, kw),
                              setup="from __main__ import {}".format(fn.func_name),
                              number = 1,
                              repeat = 3
        )
        print "{:.2f}s".format(min(times))
        result.append(min(times))
    std, opt = result
    print "Optimised takes {:.2%} time\n".format(opt/std)
    return std, opt


if __name__ == "__main__":
    timer(writer, cols=100, rows=100)
    timer(writer, cols=1000, rows=100)
    timer(writer, cols=4000, rows=100)
    timer(writer, cols=8192, rows=100)
    timer(writer, cols=10, rows=10000)
    timer(writer, cols=4000, rows=1000)
