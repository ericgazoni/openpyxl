"""Simple test charts"""
from datetime import date
import os
from openpyxl import Workbook
from openpyxl.chart import Chart, Serie, Reference, BarChart, PieChart, LineChart

wb = Workbook()
ws = wb.get_active_sheet()
ws.title = "Numbers"
for i in range(10):
    ws.append([i])
chart = BarChart()
values = Reference(ws, (0, 0), (9, 0))
series = Serie(values)
chart.add_serie(series)
ws.add_chart(chart)

ws = wb.create_sheet(1, "Negative")
for i in range(-5, 5):
    ws.append([i])
chart = BarChart()
values = Reference(ws, (0, 0), (9, 0))
series = Serie(values)
chart.add_serie(series)
ws.add_chart(chart)

ws = wb.create_sheet(2, "Letters")
for idx, l in enumerate("ABCDEFGHIJ"):
    ws.append([l, idx, idx])
chart = BarChart()
labels = Reference(ws, (0, 0), (9, 0))
values = Reference(ws, (0, 1), (9, 1))
series = Serie(values, labels=labels)
chart.add_serie(series)
#  add second series
values = Reference(ws, (0, 2), (9, 2))
series = Serie(values, labels=labels)
chart.add_serie(series)
ws.add_chart(chart)

ws = wb.create_sheet(3, "Dates")
for i in range(1, 10):
    ws.append([date(2013, i, 1), i])
chart = BarChart()
values = Reference(ws, (0, 1), (9, 1))
labels = Reference(ws, (0, 0), (9, 0))
labels.number_format = 'd-mmm'
series = Serie(values, labels=labels)
chart.add_serie(series)
ws.add_chart(chart)

ws = wb.create_sheet(4, "Pie")
for i in range(1, 5):
    ws.append([i])
chart = PieChart()
values = Reference(ws, (0, 0), (9, 0))
series = Serie(values, labels=values)
chart.add_serie(series)
ws.add_chart(chart)

ws = wb.create_sheet(5, "Line")
for i in range(1, 5):
    ws.append([i])
chart = LineChart()
values = Reference(ws, (0, 0), (4,0))
series = Serie(values)
chart.add_serie(series)
ws.add_chart(chart)

folder = os.path.split(__file__)[0]
wb.save(os.path.join(folder, "files", "charts_gen.xlsx"))
