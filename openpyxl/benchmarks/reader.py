import os
import sys
import timeit

import openpyxl


def reader(optimised):
    """
    Loop through all cells of a workbook
    """
    folder = os.path.split(__file__)[0]
    src = os.path.join(folder, "files", "large.xlsx")
    wb = openpyxl.load_workbook(src, use_iterators=optimised)
    ws = wb.get_active_sheet()
    if optimised:
        rows = ws.iter_rows()
    else:
        rows = ws.rows
    for row in rows:
        for col in row:
            pass

def timer(fn):
    """
    Create a timeit call to a function and pass in keyword arguments.
    The function is called twice, once using the standard workbook, then with the optimised one.
    Time from the best of three is taken.
    """
    result = []
    for opt in (False, True):
        print "Workbook is {}".format(opt and "optimised" or "not optimised")
        times = timeit.repeat("{}({})".format(fn.func_name, opt),
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
    timer(reader)
