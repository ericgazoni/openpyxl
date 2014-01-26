import pytest

from openpyxl import units


@pytest.mark.parametrize("value, expected",
                         [
                             (-120, -0.08333333333333334),
                             (0, 0),
                             (240, 0.16666666666666669),
                             (1440, 1),
                             (5000, 3.4722222222222223)
                         ]
                         )
def test_dxa_to_inch(value, expected):
    FUT = units.dxa_to_inch
    assert FUT(value) == expected


@pytest.mark.parametrize("value, expected",
                         [
                             (-10, -14400),
                             (0, 0),
                             (1, 1440),
                             (2.37, 3412),
                             (9, 12960),
                         ]
                         )
def test_inch_to_dxa(value, expected):
    FUT =  units.inch_to_dxa
    assert FUT(value) == expected


@pytest.mark.parametrize("value, expected",
                         [
                             (-120, -0.2116666666666667),
                             (0, 0),
                             (240, 0.4233333333333334),
                             (1440, 2.54),
                             (5000,  8.819444444444445)
                         ]
                         )
def test_dxa_to_cm(value, expected):
    FUT =  units.dxa_to_cm
    assert FUT(value) == expected


@pytest.mark.parametrize("value, expected",
                         [
                             (-10, -5669),
                             (0, 0),
                             (1, 566),
                             (10.0, 5669),
                             (1000, 566929),
                         ]
                         )
def test_cm_to_dxa(value, expected):
    FUT =  units.cm_to_dxa
    assert FUT(value) == expected


@pytest.mark.parametrize("value, expected",
                         [
                             (-10, -95250),
                             (0, 0),
                             (1, 9525),
                             (10.0, 95250),
                             (1000, 9525000),
                         ]
                         )
def test_pixels_to_EMU(value, expected):
    FUT = units.pixels_to_EMU
    assert FUT(value) == expected


@pytest.mark.parametrize("value, expected",
                         [
                            (0, 0),
                            (1000, 0),
                            (5000, 1),
                            (9525, 1),
                         ]
                         )
def test_EMU_to_pixels(value, expected):
    FUT = units.EMU_to_pixels
    assert FUT(value) == expected


@pytest.mark.parametrize("value, expected",
                         [
                             (-100000, -0.2778),
                             (0, 0),
                             (200000, 0.5556),
                             (360000, 1),
                            (500000, 1.3889),
                         ]
                         )
def test_EMU_to_cm(value, expected):
    FUT = units.EMU_to_cm
    assert FUT(value) == expected


@pytest.mark.parametrize("value, expected",
                         [
                             (-10, -3600000),
                             (0, 0),
                             (1, 360000),
                             (3.23, 1162800),
                         ]
                         )
def test_cm_to_EMU(value, expected):
    FUT = units.cm_to_EMU
    assert FUT(value) == expected


@pytest.mark.parametrize("value, expected",
                         [
                             (-100000, -0.1094),
                             (0, 0),
                             (200000, 0.2187),
                             (914400, 1),
                            (500000, 0.5468),
                         ]
                         )
def test_EMU_to_inch(value, expected):
    FUT = units.EMU_to_inch
    assert FUT(value) == expected


@pytest.mark.parametrize("value, expected",
                         [
                             (-10, -9144000),
                             (0, 0),
                             (1, 914400),
                             (3.23, 2953512),
                         ]
                         )
def test_inch_to_EMU(value, expected):
    FUT = units.inch_to_EMU
    assert FUT(value) == expected



@pytest.mark.parametrize("value, expected",
                         [
                             (-10, -7.5),
                             (0, 0),
                             (1, 0.75),
                             (96, 72),
                             (144, 108),
                         ]
                         )
def test_pixels_to_points(value, expected):
    FUT = units.pixels_to_points
    assert FUT(value) == expected


@pytest.mark.parametrize("value, expected",
                         [
                             (-10, -13),
                             (0, 0),
                             (1, 2),
                             (10.0, 14),
                             (72, 96),
                         ]
                         )
def test_points_to_pixels(value, expected):
    FUT = units.points_to_pixels
    assert FUT(value) == expected


@pytest.mark.parametrize("value, expected",
                         [
                             (-10, -600000),
                             (0, 0),
                             (1, 60000),
                             (10.0, 600000),
                             (1000, 60000000),
                         ]
                         )
def test_degrees_to_angle(value, expected):
    FUT = units.degrees_to_angle
    assert FUT(value) == expected


@pytest.mark.parametrize("value, expected",
                         [
                             (-10, 0),
                             (0, 0),
                             (10, 0),
                             (50000, 0.83),
                             (60000, 1),
                         ]
                         )
def test_angle_to_degrees(value, expected):
    FUT = units.angle_to_degrees
    assert FUT(value) == expected


@pytest.mark.parametrize("value, expected",
                         [
                             ("#FFFFF", "#FFFFF"),
                         ]
                         )
def test_short_color(value, expected):
    FUT = units.short_color
    assert FUT(value) == expected
