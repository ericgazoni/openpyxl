

def test_format_comparisions(NumberFormat):
    format1 = NumberFormat()
    format2 = NumberFormat()
    format3 = NumberFormat()
    format1.format_code = 'm/d/yyyy'
    format2.format_code = 'm/d/yyyy'
    format3.format_code = 'mm/dd/yyyy'
    assert format1 == format2
    assert format1 == 'm/d/yyyy' and format1 != 'mm/dd/yyyy'
    assert format3 != 'm/d/yyyy' and format3 == 'mm/dd/yyyy'
    assert format1 != format3


def test_builtin_format(NumberFormat):
    fmt = NumberFormat()
    fmt.format_code = '0.00'
    assert fmt.builtin_format_code(2) == fmt.format_code
