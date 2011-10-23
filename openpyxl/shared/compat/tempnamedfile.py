import tempfile
import os

def NamedTemporaryFile(mode, suffix, prefix, delete=False):

    try:
        return tempfile.NamedTemporaryFile(mode=mode, suffix=suffix, prefix=prefix, delete=delete)
    except TypeError:
        handle, filename = tempfile.mkstemp(suffix=suffix, prefix=prefix)
        os.close(handle)
        fobj = open(filename, mode)

        return fobj
