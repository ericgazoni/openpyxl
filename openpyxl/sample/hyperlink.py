# see # 88 This generated file is unreadable because relation ids are duplicated.

import os
import openpyxl


here = os.path.split(__file__)[0]

wb = openpyxl.workbook.Workbook()
ws = wb.worksheets[0]
img = openpyxl.drawing.Image('logo.png')
ws.add_image(img)
ws.cell(row=0,column=0).value = "TEXT"
ws.cell(row=0,column=0).hyperlink = "http://google.com"
wb.save(os.path.join(here, "files", 'logo.xlsx'))
