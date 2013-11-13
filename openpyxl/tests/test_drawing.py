# Copyright (c) 2010-2013 openpyxl
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#
# @license: http://www.opensource.org/licenses/mit-license.php
# @author: see AUTHORS file
import pytest


def test_bounding_box():
    from openpyxl.drawing import bounding_box
    w, h = bounding_box(80, 80, 90, 100)
    assert w == 72
    assert h == 80


class TestDrawing(object):

    def setup(self):
        from openpyxl.drawing import Drawing
        self.drawing = Drawing()

    def test_ctor(self):
        d = self.drawing
        assert d.coordinates == ((1, 2), (16, 8))
        assert d.width == 21
        assert d.height == 192
        assert d.left == 0
        assert d.top == 0
        assert d.count == 0
        assert d.rotation == 0
        assert d.resize_proportional is False
        assert d.description == ""
        assert d.name == ""

    def test_width(self):
        d = self.drawing
        d.width = 100
        d.height = 50
        assert d.width == 100

    def test_proportional_width(self):
        d = self.drawing
        d.resize_proportional = True
        d.width = 100
        d.height = 50
        assert (d.width, d.height) == (5, 50)

    def test_height(self):
        d = self.drawing
        d.height = 50
        d.width = 100
        assert d.height == 50

    def test_proportional_height(self):
        d = self.drawing
        d.resize_proportional = True
        d.height = 50
        d.width = 100
        assert (d.width, d.height) == (100, 1000)

    def test_set_dimension(self):
        d = self.drawing
        d.resize_proportional = True
        d.set_dimension(100, 50)
        assert d.width == 6
        assert d.height == 50

        d.set_dimension(50, 500)
        assert d.width == 50
        assert d.height == 417

    def test_get_emu(self):
        d = self.drawing
        dims = d.get_emu_dimensions()
        assert dims == (0, 0, 200025, 1828800)


class DummyDrawing(object):

    """Shapes need charts which need drawings"""

    width = 10
    height = 20


class DummyChart(object):

    """Shapes need a chart to calculate their coordinates"""

    width = 100
    height = 100

    def __init__(self):
        self.drawing = DummyDrawing()

    def _get_margin_left(self):
        return 10

    def _get_margin_top(self):
        return 5

    def get_x_units(self):
        return 25

    def get_y_units(self):
        return 15


class TestShape(object):

    def setup(self):
        from openpyxl.drawing import Shape
        self.shape = Shape(chart=DummyChart())

    def test_ctor(self):
        s = self.shape
        assert s.axis_coordinates == ((0, 0), (1, 1))
        assert s.text is None
        assert s.scheme == "accent1"
        assert s.style == "rect"
        assert s.border_color == "000000"
        assert s.color == "FFFFFF"
        assert s.text_color == "000000"
        assert s.border_width == 0

    def test_border_color(self):
        s = self.shape
        s.border_color = "BBBBBB"
        assert s.border_color == "BBBBBB"

    def test_color(self):
        s = self.shape
        s.color = "000000"
        assert s.color == "000000"

    def test_text_color(self):
        s = self.shape
        s.text_color = "FF0000"
        assert s.text_color == "FF0000"

    def test_border_width(self):
        s = self.shape
        s.border_width = 50
        assert s.border_width == 50

    def test_coordinates(self):
        s = self.shape
        s.coordinates = ((0, 0), (60, 80))
        assert s.axis_coordinates == ((0, 0), (60, 80))
        assert s.coordinates == (1, 1, 1, 1)

    def test_pct(self):
        s = self.shape
        assert s._norm_pct(10) == 1
        assert s._norm_pct(0.5) == 0.5
        assert s._norm_pct(-10) == 0


class TestShadow(object):

    def setup(self):
        from openpyxl.drawing import Shadow
        self.shadow = Shadow()

    def test_ctor(self):
        s = self.shadow
        assert s.visible == False
        assert s.blurRadius == 6
        assert s.distance == 2
        assert s.direction == 0
        assert s.alignment == "br"
        assert s.color.index == "FF000000"
        assert s.alpha == 50


try:
    from PIL import Image
except ImportError:
    Image = False

import os
from .helper import DATADIR


class DummySheet(object):
    """Required for images"""

    def point_pos(self, vertical, horizontal):
        return vertical, horizontal


class DummyCell(object):
    """Required for images"""

    column = "A"
    row = "1"
    anchor = (0, 0)

    def __init__(self):
        self.parent = DummySheet()


class TestImage(object):

    def setup(self):
        self.img = img = os.path.join(DATADIR, "plain.png")

    def make_one(self):
        from openpyxl.drawing import Image
        return Image

    @pytest.mark.skipif(Image, reason="PIL is installed installed")
    def test_import(self):
        Image = self.make_one()
        with pytest.raises(ImportError):
            i = Image._import_image(self.img)

    @pytest.mark.skipif(Image is False, reason="PIL must be installed")
    def test_ctor(self):
        Image = self.make_one()
        i = Image(img=self.img)
        assert i.nochangearrowheads == True
        assert i.nochangeaspect == True
        d = i.drawing
        assert d.coordinates == ((0, 0), (1, 1))
        assert d.width == 118
        assert d.height == 118

    @pytest.mark.skipif(Image is False, reason="PIL must be installed")
    def test_anchor(self):
        Image = self.make_one()
        i = Image(self.img)
        c = DummyCell()
        vals = i.anchor(c)
        assert vals == (('A', '1'), (118, 118))
