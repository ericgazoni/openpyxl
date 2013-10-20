from openpyxl import load_workbook

wb = load_workbook("files/concatenate.xlsx")
ws = wb.get_active_sheet()

b1 = ws.cell('B1')
a6 = ws.cell('A6')

assert b1.value == '=CONCATENATE(A1,A2)'
assert b1.data_type == 'f'

assert a6.value == '=SUM(A4:A5)'
assert a6.data_type == 'f'

# test iterator

wb = load_workbook("files/concatenate.xlsx", True)
ws = wb.get_active_sheet()

for row in ws.iter_rows():
    for col in row:
        if col.coordinate == 'B1':
            b1 = col
        elif col.coordinate == 'A6':
            a6 = col

assert b1.internal_value == '=CONCATENATE(A1,A2)'
assert b1.data_type == 'f'

assert a6.internal_value == '=SUM(A4:A5)'
assert a6.data_type == 'f'
