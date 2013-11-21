
import pytest

@pytest.fixture
def number_format():
    from openpyxl.styles import NumberFormat
    return NumberFormat


def test_format_comparisions(number_format):
    format1 = number_format()
    format2 = number_format()
    format3 = number_format()
    format1.format_code = 'm/d/yyyy'
    format2.format_code = 'm/d/yyyy'
    format3.format_code = 'mm/dd/yyyy'
    assert format1 == format2
    assert format1 == 'm/d/yyyy' and format1 != 'mm/dd/yyyy'
    assert format3 != 'm/d/yyyy' and format3 == 'mm/dd/yyyy'
    assert format1 != format3


def test_builtin_format(number_format):
    fmt = number_format()
    fmt.format_code = '0.00'
    assert fmt.builtin_format_code(2) == fmt.format_code
