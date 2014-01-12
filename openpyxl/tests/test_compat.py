import pytest


@pytest.mark.parametrize("value, result",
                         [
                          ('s', 's'),
                          (2.0/3, '0.666666666666667'),
                          (1, '1'),
                          (None, 'None')
                         ]
                         )
def test_safe_string(value, result):
    from openpyxl.writer.charts import safe_string
    assert safe_string(value) == result
    v = safe_string('s')
    assert v == 's'
