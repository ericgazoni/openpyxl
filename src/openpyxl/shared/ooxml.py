# file openpyxl/shared/ooxml.py

PACKAGE_PROPS = 'docProps'
PACKAGE_XL = 'xl'
PACKAGE_RELS = '_rels'
PACKAGE_THEME = PACKAGE_XL + '/' + 'theme'
PACKAGE_WORKSHEETS = PACKAGE_XL + '/' + 'worksheets'

ARC_CONTENT_TYPES = '[Content_Types].xml'
ARC_ROOT_RELS = PACKAGE_RELS + '/.rels'
ARC_WORKBOOK_RELS = PACKAGE_XL + '/' + PACKAGE_RELS + '/workbook.xml.rels'
ARC_CORE = PACKAGE_PROPS + '/core.xml'
ARC_APP = PACKAGE_PROPS + '/app.xml'
ARC_WORKBOOK = PACKAGE_XL + '/workbook.xml'
ARC_STYLE = PACKAGE_XL + '/styles.xml'
ARC_THEME = PACKAGE_THEME + '/theme1.xml'
ARC_SHARED_STRINGS = PACKAGE_XL + '/sharedStrings.xml'

NAMESPACES = {
'cp' : 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
'dc' : 'http://purl.org/dc/elements/1.1/',
'dcterms' : 'http://purl.org/dc/terms/',
'dcmitype' : 'http://purl.org/dc/dcmitype/',
'xsi' : 'http://www.w3.org/2001/XMLSchema-instance',
'vt' : 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
'xml' : 'http://www.w3.org/XML/1998/namespace'
}
