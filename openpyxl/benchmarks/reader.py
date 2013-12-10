import os
import sys
import timeit

import openpyxl


def reader(optimised):
    """
    Create a worksheet with variable width rows. Because data must be
    serialised row by row it is often the width of the rows which is most
    important.
    """
    folder = os.path.split(__file__)[0]
    src = os.path.join(folder, "files", "large.xlsx")
    wb = openpyxl.load_workbook(src, use_iterators=optimised)


def timer(fn):
    """
    Create a timeit call to a function and pass in keyword arguments.
    The function is called twice, once using the standard workbook, then with the optimised one.
    Time from the best of three is taken.
    """
    result = []
    for opt in (False, True):
        print "Worksbook is {}".format(opt and "optimised" or "not optimised")
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
