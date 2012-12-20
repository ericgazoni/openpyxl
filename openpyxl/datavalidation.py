# file openpyxl/datavalidations.py

# Copyright (c) 2010-2012 openpyxl
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

"""
  <xsd:complexType name="CT_DataValidations">
    <xsd:sequence>
      <xsd:element name="dataValidation" type="CT_DataValidation" minOccurs="1"
        maxOccurs="unbounded"/>
    </xsd:sequence>
    <xsd:attribute name="disablePrompts" type="xsd:boolean" use="optional" default="false"/>
    <xsd:attribute name="xWindow" type="xsd:unsignedInt" use="optional"/>
    <xsd:attribute name="yWindow" type="xsd:unsignedInt" use="optional"/>
    <xsd:attribute name="count" type="xsd:unsignedInt" use="optional"/>
  </xsd:complexType>
  <xsd:complexType name="CT_DataValidation">
    <xsd:sequence>
      <xsd:element name="formula1" type="ST_Formula" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="formula2" type="ST_Formula" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="type" type="ST_DataValidationType" use="optional" default="none"/>
    <xsd:attribute name="errorStyle" type="ST_DataValidationErrorStyle" use="optional"
      default="stop"/>
    <xsd:attribute name="imeMode" type="ST_DataValidationImeMode" use="optional" default="noControl"/>
    <xsd:attribute name="operator" type="ST_DataValidationOperator" use="optional" default="between"/>
    <xsd:attribute name="allowBlank" type="xsd:boolean" use="optional" default="false"/>
    <xsd:attribute name="showDropDown" type="xsd:boolean" use="optional" default="false"/>
    <xsd:attribute name="showInputMessage" type="xsd:boolean" use="optional" default="false"/>
    <xsd:attribute name="showErrorMessage" type="xsd:boolean" use="optional" default="false"/>
    <xsd:attribute name="errorTitle" type="s:ST_Xstring" use="optional"/>
    <xsd:attribute name="error" type="s:ST_Xstring" use="optional"/>
    <xsd:attribute name="promptTitle" type="s:ST_Xstring" use="optional"/>
    <xsd:attribute name="prompt" type="s:ST_Xstring" use="optional"/>
    <xsd:attribute name="sqref" type="ST_Sqref" use="required"/>
  </xsd:complexType>
  <xsd:simpleType name="ST_DataValidationType">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="none"/>
      <xsd:enumeration value="whole"/>
      <xsd:enumeration value="decimal"/>
      <xsd:enumeration value="list"/>
      <xsd:enumeration value="date"/>
      <xsd:enumeration value="time"/>
      <xsd:enumeration value="textLength"/>
      <xsd:enumeration value="custom"/>
    </xsd:restriction>
  </xsd:simpleType>
  <xsd:simpleType name="ST_DataValidationOperator">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="between"/>
      <xsd:enumeration value="notBetween"/>
      <xsd:enumeration value="equal"/>
      <xsd:enumeration value="notEqual"/>
      <xsd:enumeration value="lessThan"/>
      <xsd:enumeration value="lessThanOrEqual"/>
      <xsd:enumeration value="greaterThan"/>
      <xsd:enumeration value="greaterThanOrEqual"/>
    </xsd:restriction>
  </xsd:simpleType>
  <xsd:simpleType name="ST_DataValidationErrorStyle">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="stop"/>
      <xsd:enumeration value="warning"/>
      <xsd:enumeration value="information"/>
    </xsd:restriction>
  </xsd:simpleType>
  <xsd:simpleType name="ST_DataValidationImeMode">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="noControl"/>
      <xsd:enumeration value="off"/>
      <xsd:enumeration value="on"/>
      <xsd:enumeration value="disabled"/>
      <xsd:enumeration value="hiragana"/>
      <xsd:enumeration value="fullKatakana"/>
      <xsd:enumeration value="halfKatakana"/>
      <xsd:enumeration value="fullAlpha"/>
      <xsd:enumeration value="halfAlpha"/>
      <xsd:enumeration value="fullHangul"/>
      <xsd:enumeration value="halfHangul"/>
    </xsd:restriction>
  </xsd:simpleType>
"""
class DataValidation(object):
    def __init__(self, type="list", formula1=""):
        self.type = type
        self.formula1 = formula1
        self.forumla2 = None
        self.attribute_map = {}
        self.cells = []

    def set_validation_on_cell(self, cell_address):
        self.cells.append(cell_address)
        
    def generate_attributes_map(self):
        # Update the sqref to ensure it points at all cells we're interested in
        self.attribute_map['sqref'] = collapse_cell_addresses(self.cells)
        return self.attribute_map
        