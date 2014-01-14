"""
Very simple memory use analysis
"""
import os
import openpyxl

from memory_profiler import memory_usage


def test_memory_use():
    """Naive test that assumes memory use will never be more than 120 % of
    that for first 50 rows"""
    folder = os.path.split(__file__)[0]
    src = os.path.join(folder, "files", "very_large.xlsx")
    wb = openpyxl.load_workbook(src, use_iterators=True)
    ws = wb.get_active_sheet()

    initial_use = None

    for n, line in enumerate(ws.iter_rows()):
        if n % 50 == 0:
            use = memory_usage(proc=-1, interval=1)[0]
            if initial_use is None:
                initial_use = use
            assert use/initial_use < 1.2
            print n, use

if __name__ == '__main__':
    test_memory_use()
