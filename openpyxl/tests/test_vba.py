# Copyright (c) 2010-2014 openpyxl
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

# Python stdlib imports
import zipfile

# compatibility imports
from openpyxl.compat import BytesIO

# package imports
from openpyxl.reader.excel import load_workbook
from openpyxl.writer.excel import save_virtual_workbook


def test_save_vba(datadir):
    path = str(datadir.join('reader').join('vba-test.xlsm'))
    wb = load_workbook(path, keep_vba=True)
    buf = save_virtual_workbook(wb)
    files1 = set(zipfile.ZipFile(path, 'r').namelist())
    files2 = set(zipfile.ZipFile(BytesIO(buf), 'r').namelist())
    assert files1.issubset(files2), "Missing files: %s" % ', '.join(files1 - files2)


def test_save_without_vba(datadir):
    path = str(datadir.join('reader').join('vba-test.xlsm'))
    vbFiles = set(['xl/activeX/activeX2.xml', 'xl/drawings/_rels/vmlDrawing1.vml.rels',
                   'xl/activeX/_rels/activeX1.xml.rels', 'xl/drawings/vmlDrawing1.vml', 'xl/activeX/activeX1.bin',
                   'xl/media/image1.emf', 'xl/vbaProject.bin', 'xl/activeX/_rels/activeX2.xml.rels',
                   'xl/worksheets/_rels/sheet1.xml.rels', 'customUI/customUI.xml', 'xl/media/image2.emf',
                   'xl/ctrlProps/ctrlProp1.xml', 'xl/activeX/activeX2.bin', 'xl/activeX/activeX1.xml',
                   'xl/ctrlProps/ctrlProp2.xml', 'xl/drawings/drawing1.xml'])

    wb = load_workbook(path, keep_vba=False)
    buf = save_virtual_workbook(wb)
    files1 = set(zipfile.ZipFile(path, 'r').namelist())
    files2 = set(zipfile.ZipFile(BytesIO(buf), 'r').namelist())
    difference = files1.difference(files2)
    assert difference.issubset(vbFiles), "Missing files: %s" % ', '.join(difference - vbFiles)

def test_save_same_file(tmpdir, datadir):
    p1 = datadir.join('reader').join('vba-test.xlsm')
    p2 = tmpdir.join('vba-test.xlsm')
    p1.copy(p2)
    path = str(p2)
    wb = load_workbook(path, keep_vba=True)
    wb.save(path)
