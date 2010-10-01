# file openpyxl/shared/zip.py
from zipfile import ZipFile, ZIP_DEFLATED

class ZipArchive(object):

    def __init__(self, filename, mode = 'r'):

        self._filename = filename
        try :
            self._zipfile = ZipFile(file = filename,
                                    mode = mode,
                                    compression = ZIP_DEFLATED,
                                    allowZip64 = False)
        except:
            self._zipfile = ZipFile(file = filename,
                                    mode = mode,
                                    compression = ZIP_DEFLATED)

    def is_in_archive(self, arc_name):

        try:
            self._zipfile.getinfo(name = arc_name)
            return True
        except KeyError:
            return False

    def add_from_string(self, arc_name, content):

        self._zipfile.writestr(arc_name, content)

    def add_from_file(self, arc_name, content):

        self._zipfile.write(content, arc_name)

    def get_from_name(self, arc_name):

        return self._zipfile.read(arc_name)

    def close(self):

        self._zipfile.close()
