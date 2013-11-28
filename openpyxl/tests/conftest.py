# Fixtures (pre-configured objects) for tests

import pytest

@pytest.fixture
def NumberFormat():
    from openpyxl.styles import NumberFormat
    return NumberFormat
