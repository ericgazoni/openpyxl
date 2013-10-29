Development
================

With the ongoing development of openpyxl, there is occasional information
useful to assist developers.  This documentation contains examples for
making the development process easier.


Benchmarking
-----------------

As openpyxl does not include any internal memory benchmarking tools, the python `pympler` package was used
during the testing of styles to profile the memory usage in :def:`openpyxl.reader.excel.read_style_table()`::

    # in openpyxl/reader/style.py
    from pympler import muppy, summary

    def read_style_table(xml_source):
      ...
      if cell_xfs is not None:  # ~ line 47
          initialState = summary.summarize(muppy.get_objects())  # Capture the initial state
          for index, cell_xfs_node in enumerate(cell_xfs_nodes):
             ...
             table[index] = new_style
          finalState = summary.summarize(muppy.get_objects())  # Capture the final state
          diff = summary.get_diff(initialState, finalState)  # Compare
          summary.print_(diff)


`pympler.summary` prints to the console a report of object memory usage, allowing the comparison of different
methods and examination of memory usage.  A useful future development would be to construct a benchmarking package to
measure the performance of different components.