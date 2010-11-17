#-*- coding: iso-8859-1 -*-

import random

from openpyxl.workbook import Workbook
from openpyxl.writer.excel import save_workbook
from openpyxl.drawing import Shape
from openpyxl.style import Color, Border
from openpyxl.chart import LineChart, BarChart, Serie, ErrorBar, Reference


wb = Workbook()

ws = wb.get_active_sheet()

ws.title = u'data'

cell = ws.cell('H2')
cell.value = u'jean-rené'
cell.style.font.bold = True
cell.style.font.size = '15'
cell.style.borders.top.border_style = Border.BORDER_THIN
cell.style.borders.top.color.index = Color.DARKYELLOW

for i in range(4):
    ws.cell(row=i, column=0).value = chr(65+i)
    
for i in range(4):
    ws.cell(row=i, column=1).value = i
for i in range(5,9):
    ws.cell(row=i-5, column=2).value = i
    
for i in range(4):
    ws.cell(row=i, column=3).value = random.random()
    
serie1 = Serie(Reference(ws, (0,1), (3,1)), 
    labels=Reference(ws, (0,0), (3,0)), 
    legend=Reference(ws, (0,0)))
    
serie1.color = Color.DARKGREEN
serie1.error_bar = ErrorBar(ErrorBar.PLUS_MINUS, Reference(ws, (0,3), (3,3)))

serie2 = Serie(Reference(ws, (0,2), (3,2)), labels=Reference(ws, (0,0), (3,0)))

chart = LineChart()
chart.add_serie(serie1)
chart.add_serie(serie2)

# chart container dimensions in pixels
chart.drawing.left = 10
chart.drawing.top = 150
chart.drawing.height = 200
chart.drawing.width = 500

# chart area in percentage of the container
chart.width = .7
chart.height = .7
chart.margin_top = .2

# shapes are positionned in graph coordinates
s = Shape(((0,0), (2,-0.5)))
chart.add_shape(s)

s = Shape(((3,-.5), (4,1.5)))
s.border_color = Color.RED
s.border_width = 2
s.text = u'lolo'
chart.add_shape(s)

ws.add_chart(chart)

##chart = BarChart()
##chart.add_serie(serie)
##chart.drawing.height = 200
##chart.drawing.width = 500
##
##sheet2 = wb.create_sheet()
##sheet2.title = 'sheet2'
##sheet2.add_chart(chart)
##
##ws.add_chart(chart)

wb.save(r'c:\temp\xl_xml\toto.xlsx')
