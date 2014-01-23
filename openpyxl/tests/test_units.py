from openpyxl import units


def test_cm_to_pixels():
    FUT = units.cm_to_pixels


def test_pixels_to_cm():
    FUT = units.pixels_to_cm


def test_pixels_to_EMU():
    FUT = units.pixels_to_EMU


def test_EMU_to_pixels():
    FUT = units.EMU_to_pixels


def test_EMU_to_cm():
    FUT = units.EMU_to_cm


def test_pixels_to_points():
    FUT = units.pixels_to_points


def test_points_to_pixels():
    FUT = units.points_to_pixels


def test_degrees_to_angle():
    FUT = units.degrees_to_angle


def test_angle_to_degrees():
    FUT = units.angle_to_degrees


def test_short_color():
    FUT = units.short_color
