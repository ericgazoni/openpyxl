# file openpyxl/tests/test_zip.py

import os.path as osp

from zipfile import ZipFile, ZIP_DEFLATED

from openpyxl.tests.helper import BaseTestCase, TMPDIR
from openpyxl.shared.zip import ZipArchive

class TestZip(BaseTestCase):

    def test_write_zip(self):

        filename = osp.join(TMPDIR, 'test.zip')

        inner_filename = 'file.a'
        inner_content = "here is the content"


        z = ZipArchive(filename = filename, mode = 'w')

        z.add_from_string(inner_filename, inner_content)

        z.close()


        test_zip = ZipFile(file = filename,
                           mode = 'r',
                           compression = ZIP_DEFLATED,
                           allowZip64 = True)

        self.assertTrue(inner_filename in test_zip.namelist())

        self.assertEqual(test_zip.read(inner_filename), inner_content)

        test_zip.close()

    def test_read_zip(self):

        filename = osp.join(TMPDIR, 'test.zip')

        inner_filename = 'file.a'
        inner_content = "here is the content"

        # write the zip file
        z = ZipArchive(filename = filename, mode = 'w')
        z.add_from_string(inner_filename, inner_content)
        z.close()

        # read it again
        z = ZipArchive(filename = filename)
        read_content = z.get_from_name(inner_filename)
        z.close()

        self.assertEqual(read_content, inner_content)
