Comments
========

.. warning::

    Openpyxl currently supports the reading and writing of comment text only.
    Formatting information is lost.


Adding a comment to a cell
--------------------------

Comments have a text attribute and an author attribute, which must both be set

>>> from openpyxl import Workbook
>>> from openpyxl.comments import Comment
>>> wb = Workbook()
>>> ws = wb.get_active_sheet()
>>> ws.cell(coordinate="A1").comment = Comment('This is the comment text', 'Comment Author')
>>> ws.cell(coordinate="A1").text
'This is the comment text'
>>> ws.cell(coordinate="A1").author
'Comment Author'

You cannot assign the same Comment object to two different cells. Doing so raises an AttributeError.

>>> from openpyxl import Workbook
>>> from openpyxl.comments import Comment
>>> wb=Workbook()
>>> ws=wb.get_active_sheet()
>>> comment = Comment("Text", "Author")
>>> ws.cell(coordinate="A1").comment = comment
>>> ws.cell(coordinate="B2").comment = comment
AttributeError: Comment already assigned to A1 in worksheet Sheet. Cannot assign a comment to more than one cell

Loading and saving comments
----------------------------

Comments present in a workbook when loaded are stored in the comment attribute of their respective cells automatically.
Formatting information such as font size, bold and italics are lost, as are the original dimensions and position of the comment's container box.

Comments remaining in a workbook when it is saved are automatically saved to the workbook file.
