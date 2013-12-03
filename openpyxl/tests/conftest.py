# Fixtures (pre-configured objects) for tests

import pytest

# objects under test

@pytest.fixture
def NumberFormat():
    from openpyxl.styles import NumberFormat
    return NumberFormat


@pytest.fixture
def Workbook():
    from openpyxl import Workbook
    return Workbook


@pytest.fixture
def ws(Workbook):
    wb = Workbook()
    ws = wb.get_active_sheet()
    ws.title = 'data'
    return ws


@pytest.fixture
def ten_row_sheet(ws):
    for i in range(10):
        ws.cell(row=i, column=0).value = i
    return ws


@pytest.fixture
def ten_column_sheet(ws):
    ws.append(list(range(10)))
    return ws
