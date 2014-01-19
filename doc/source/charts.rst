Charts
======

.. warning::

    Openpyxl currently supports chart creation within a worksheet only. Charts in
    existing workbooks will be lost.


Chart types
-----------

The following charts are available:

* Bar Chart
* Line Chart
* Scatter Chart
* Pie Chart


Creating a chart
----------------

Charts are composed of at least one series of one or more data points. Series
themselves are comprised of references to cell ranges.

>>> from openpyxl import Workbook
>>> wb = Workbook()
>>> ws = wb.active
>>> for i in range(10):
>>>     ws.append(i)
>>>
>>> from openpyxl.charts import BarChart, Reference, Series
>>> values = Reference(ws, (0, 0), (9, 0))
>>> series = Series(values, title="First series of values")
>>> chart = BarChart()
>>> chart.append(series)
>>> ws.add_chart()
>>> wb.save("SampleChart.xlsx")
