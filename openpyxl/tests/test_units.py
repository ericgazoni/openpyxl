from openpyxl import units


def test_cm_to_pixels():
    FUT = units.cm_to_pixels
    assert FUT(1) == 44


def test_pixels_to_cm():
    FUT = units.pixels_to_cm
    assert FUT(1) == 0


def test_pixels_to_EMU():
    FUT = units.pixels_to_EMU
    assert FUT(1) == 9525


def test_EMU_to_pixels():
    FUT = units.EMU_to_pixels
    assert FUT(1) == 0


def test_EMU_to_cm():
    FUT = units.EMU_to_cm
    assert FUT(1) == 0


def test_pixels_to_points():
    FUT = units.pixels_to_points
    assert FUT(1) == 0.67777777


def test_points_to_pixels():
    FUT = units.points_to_pixels
    assert FUT(1) == 2


def test_degrees_to_angle():
    FUT = units.degrees_to_angle
    assert FUT(1) == 60000


def test_angle_to_degrees():
    FUT = units.angle_to_degrees
    assert FUT(1) == 0


def test_short_color():
    FUT = units.short_color
    assert FUT("#FFFFF") == "#FFFFF"
