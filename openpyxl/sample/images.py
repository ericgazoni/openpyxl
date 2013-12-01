import os

from openpyxl import Workbook
from openpyxl.drawing import Image
from openpyxl.tests.helper import DATADIR

wb = Workbook()
ws = wb.get_active_sheet()
ws.cell('A1').value = 'You should see a logo below'

# create an image instance
pth = os.path.split(__file__)[0]
img = Image(os.path.join(DATADIR, 'plain.png'))

# place it if required
img.drawing.left = 200
img.drawing.top = 100

# you could also 'anchor' the image to a specific cell
# img.anchor(ws.cell('B12'))

# add to worksheet
ws.add_image(img)
wb.save(os.path.join(pth, 'files', 'logo.xlsx'))
