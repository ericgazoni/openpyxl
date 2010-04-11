'''
Copyright (c) 2010 openpyxl

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.

@license: http://www.opensource.org/licenses/mit-license.php
@author: Eric Gazoni
'''

from xml.etree.cElementTree import fromstring, QName
from openpyxl.shared.ooxml import NAMESPACES, ARC_CORE
from openpyxl.workbook import DocumentProperties
from openpyxl.shared.date_time import W3CDTF_to_datetime

def read_properties_core(archive):

    properties = DocumentProperties()

    content = archive.get_from_name(arc_name = ARC_CORE)

    root = fromstring(text = content)

    properties.creator = root.find(QName(NAMESPACES['dc'], 'creator').text).text
    properties.last_modified_by = root.find(QName(NAMESPACES['cp'], 'lastModifiedBy').text).text

    properties.created = W3CDTF_to_datetime(root.find(QName(NAMESPACES['dcterms'], 'created').text).text)
    properties.modified = W3CDTF_to_datetime(root.find(QName(NAMESPACES['dcterms'], 'modified').text).text)

    return properties
