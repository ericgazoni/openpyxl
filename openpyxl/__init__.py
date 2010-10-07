# file openpyxl/__init__.py

"""Maybe these should go in setup?"""

# package imports
from . import cell
from . import namedrange
from . import style
from . import workbook
from . import worksheet
from . import reader
from . import shared
from . import writer

# constants
__major__ = 1       # for major interface/format changes
__minor__ = 1       # for minor interface/format changes
__release__ = 8     # for tweaks, bug-fixes, or development

__version__ = '%d.%d.%d' % (__major__,
                            __minor__,
                            __release__)

__author__ = 'Eric Gazoni'
__license__ = 'MIT/Expat'
__author_email__ = 'eric.gazoni@gmail.com'
__maintainer_email__ = 'openpyxl-users@googlegroups.com'
__url__ = 'http://bitbucket.org/ericgazoni/openpyxl/wiki/Home'
__downloadUrl__ = "http://bitbucket.org/ericgazoni/openpyxl/downloads"

__all__ = ('reader', 'shared', 'writer', )
