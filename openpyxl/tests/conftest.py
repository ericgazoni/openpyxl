# Fixtures (pre-configured objects) for tests
import sys
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


@pytest.fixture
def Worksheet():
    """Worksheet Class"""
    from openpyxl.worksheet import Worksheet
    return Worksheet


# Charts

@pytest.fixture
def Chart():
    """Chart class"""
    from openpyxl.charts.chart import Chart
    return Chart


@pytest.fixture
def GraphChart():
    """GraphicChart class"""
    from openpyxl.charts.chart import GraphChart
    return GraphChart


@pytest.fixture
def Axis():
    """Axis class"""
    from openpyxl.charts.axis import Axis
    return Axis


@pytest.fixture
def PieChart():
    """PieChart class"""
    from openpyxl.charts import PieChart
    return PieChart


@pytest.fixture
def LineChart():
    """LineChart class"""
    from openpyxl.charts import LineChart
    return LineChart


@pytest.fixture
def BarChart():
    """BarChart class"""
    from openpyxl.charts import BarChart
    return BarChart


@pytest.fixture
def ScatterChart():
    """ScatterChart class"""
    from openpyxl.charts import ScatterChart
    return ScatterChart


@pytest.fixture
def Reference():
    """Reference class"""
    from openpyxl.charts import Reference
    return Reference


@pytest.fixture
def Series():
    """Serie class"""
    from openpyxl.charts import Series
    return Series


@pytest.fixture
def ErrorBar():
    """ErrorBar class"""
    from openpyxl.charts import ErrorBar
    return ErrorBar


@pytest.fixture
def Image():
    """Image class"""
    from openpyxl.drawing import Image
    return Image


# Styles

@pytest.fixture(autouse=True)
def FormatRule():
    """Formatting rule class"""
    from openpyxl.styles.formatting import FormatRule
    return FormatRule


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
        elif item.get_marker("not_py33"):
            pytest.skip("Ordering is not a given in Python 3")
        elif item.get_marker("lxml_required"):
            pytest.skip("LXML is required for some features such as schema validation")
