import pytest

from openpyxl import units


@pytest.mark.parametrize("value, expected",
                        [
                            (-10, -446),
                            (0, 0),
                            (1, 44),
                            (10.0, 446),
                            (1000, 44600)
                        ]
                        )
def test_cm_to_pixels(value, expected):
    FUT = units.cm_to_pixels
    assert FUT(value) == expected


@pytest.mark.parametrize("value, expected",
                         [
                             (-10, -1),
                             (0, 0),
                             (1, 0),
                             (10.0, 0),
                             (1000, 0),
                         ]
                         )
def test_pixels_to_cm(value, expected):
    FUT = units.pixels_to_cm
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
                            (-10, 0),
                            (0, 0),
                            (1, 0),
                            (10.0, 0),
                            (1000, 0),
                         ]
                         )
def test_EMU_to_pixels(value, expected):
    FUT = units.EMU_to_pixels
    assert FUT(value) == expected


@pytest.mark.parametrize("value, expected",
                         [
                             (-10, 0),
                             (0, 0),
                             (1, 0),
                             (10.0, 0),
                            (1000, 0),
                         ]
                         )
def test_EMU_to_cm(value, expected):
    FUT = units.EMU_to_cm
    assert FUT(value) == expected


@pytest.mark.parametrize("value, expected",
                         [
                             (-10, -6.7777777),
                             (0, 0),
                             (1, 0.67777777),
                             (10.0, 6.7777777),
                             (1000, 677.7777699999999),
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
                             (1000, 1334),
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
                             (1, 0),
                             (10.0, 0),
                             (1000, 0),
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
