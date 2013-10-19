# file openpyxl/tests/helper.py

# Copyright (c) 2010-2011 openpyxl
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

import os
from zipfile import ZipFile

from lxml.etree import XMLSchema
from lxml.etree import tostring, fromstring, parse

# Provide schema based validators, lxml required
# use schema.validate(Element) or schema.assertValid(Element) for messages

SCHEMA_FOLDER = os.path.join(os.path.dirname(__file__), 'schemas')

sheet_src = os.path.join(SCHEMA_FOLDER, 'sml.xsd')
sheet_schema = XMLSchema(file=sheet_src)

chart_src = os.path.join(SCHEMA_FOLDER, 'dml-chart.xsd')
chart_schema = XMLSchema(file=chart_src)

drawing_src = os.path.join(SCHEMA_FOLDER, 'dml-spreadsheetDrawing.xsd')
drawing_schema = XMLSchema(file=drawing_src)

sml_files = ['xl/styles.xml']  # , 'xl/workbook.xml']


def validate_archive(file_path):
    zipfile = ZipFile(file_path)
    try:
        for entry in zipfile.infolist():
            filename = entry.filename
            f = zipfile.open(entry)
            root = parse(f).getroot()
            if filename in sml_files or filename.startswith('xl/worksheets/sheet'):
                if root.get('{http://www.w3.org/XML/1998/namespace}space'):
                    # not allowed by schema
                    del root.attrib['{http://www.w3.org/XML/1998/namespace}space']
                sheet_schema.assertValid(root)
    finally:
        zipfile.close()
