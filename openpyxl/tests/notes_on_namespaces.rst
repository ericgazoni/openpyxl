Correct use of namespaces when generating
=========================================


Current situation
-----------------

Most namespace tags are directly generated in openpyxl without explicit or
reliable use of namespaces, eg. `Element("c:valAx")`. These should be
replaced using qualified tagnames:
`Element("{http://schemas.openxmlformats.org/drawingml/2006/chart}valAx").
This, together with registered namespace prefixes ensures that tags are always
correctly generated and that serialised XML is valid.
