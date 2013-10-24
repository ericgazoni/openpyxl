from openpyxl.style import Style, Font
from nose.tools.trivial import eq_


def test_style_builder():
    expected = Style(font=Font(size=31))

    base = Style(font=Font(size=13))
    fixture = base.copy(font=Font(size=31))

    eq_(expected, fixture)


