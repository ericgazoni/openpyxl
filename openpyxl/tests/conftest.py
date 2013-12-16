# Fixtures (pre-configured objects) for tests

import pytest

# objects under test

@pytest.fixture
def NumberFormat():
    """NumberFormat Class"""
    from openpyxl.styles import NumberFormat
    return NumberFormat


@pytest.fixture
def Workbook():
    """Workbook Class"""
    from openpyxl import Workbook
    return Workbook

# Charts

@pytest.fixture
def Chart():
    """Chart class"""
    from openpyxl.chart import Chart
    return Chart


@pytest.fixture
def GraphChart():
    """GraphicChart class"""
    from openpyxl.chart import GraphChart
    return GraphChart


@pytest.fixture
def Axis():
    """Axis class"""
    from openpyxl.chart import Axis
    return Axis


@pytest.fixture
def PieChart():
    """PieChart class"""
    from openpyxl.chart import PieChart
    return PieChart


@pytest.fixture
def LineChart():
    """LineChart class"""
    from openpyxl.chart import LineChart
    return LineChart


@pytest.fixture
def BarChart():
    """BarChart class"""
    from openpyxl.chart import BarChart
    return BarChart


@pytest.fixture
def ScatterChart():
    """ScatterChart class"""
    from openpyxl.chart import ScatterChart
    return ScatterChart


@pytest.fixture
def Reference():
    """Reference class"""
    from openpyxl.chart import Reference
    return Reference


@pytest.fixture
def Series():
    """Serie class"""
    from openpyxl.chart import Series
    return Series


@pytest.fixture
def ErrorBar():
    """ErrorBar class"""
    from openpyxl.chart import ErrorBar
    return ErrorBar


@pytest.fixture
def Image():
    """Image class"""
    from openpyxl.drawing import Image
    return Image


# utility fixtures

@pytest.fixture
def ws(Workbook):
    """Empty worksheet titled 'data'"""
    wb = Workbook()
    ws = wb.get_active_sheet()
    ws.title = 'data'
    return ws


@pytest.fixture
def ten_row_sheet(ws):
    """Worksheet with values 0-9 in the first column"""
    for i in range(10):
        ws.cell(row=i, column=0).value = i
    return ws


@pytest.fixture
def ten_column_sheet(ws):
    """Worksheet with values 0-9 in the first row"""
    ws.append(list(range(10)))
    return ws

from openpyxl.shared.xmltools import Element

@pytest.fixture
def root_xml():
    """Root XML element <test>"""
    return Element("test")


### Markers ###

def pytest_runtest_setup(item):
    if isinstance(item, item.Function):
        try:
            from PIL import Image
        except ImportError:
            Image = False
        if item.get_marker("pil_required") and Image is False:
            pytest.skip("PIL must be installed")
        elif item.get_marker("pil_not_installed") and Image:
            pytest.skip("PIL is installed")
