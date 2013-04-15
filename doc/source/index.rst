:mod:`openpyxl` - A Python library to read/write Excel 2007 xlsx/xlsm files
=============================================================================

.. module:: openpyxl
.. moduleauthor:: Eric Gazoni

:Author: Eric Gazoni
:Source code: http://bitbucket.org/ericgazoni/openpyxl/src
:Issues: http://bitbucket.org/ericgazoni/openpyxl/issues
:Generated: |today|
:License: MIT/Expat
:Version: |release|

Introduction
------------

OpenPyxl is a Python library to read/write Excel 2007 xlsx/xlsm files.

It was born from lack of existing library to read/write natively from Python the new Open Office XML format.

All kudos to the PHPExcel team as openpyxl is a Python port of PHPExcel http://www.phpexcel.net/

User List
---------

Official user list can be found on http://groups.google.com/group/openpyxl-users

Contribute
----------

Any help will be greatly appreciated, there are just a few requirements to get your code checked in the public repository:

    * Forks are the now prefered way to contribute, but don't forget to make a pull request if you want your code to be included in the main branch :)
    * long diffs posted in the body of a tracker request will not be looked at (more than 30 rows of non-syntax highlighted code is simply unreadable).     
    * every non-trivial change must come with at least a unit test (that tests the new behavior, obviously :p). There are plenty of examples in the /test directory if you lack know-how or inspiration.


Installation
------------

The best method to install openpyxl is using a PyPi client such as easy_install (setuptools) or pip::

    $ pip install openpyxl

or ::

    $ easy_install install openpyxl

.. note::

    To install from sources (there is nothing to build, openpyxl is 100% pure Python), you can download an archive from https://bitbucket.org/ericgazoni/openpyxl/downloads (look in the "tags" tab).
    After extracting the archive, you can do::

    $ python setup.py install 

.. warning::
    
    To be able to include images (jpeg,png,bmp,...) into an openpyxl file, you will also need the 'PIL' library that can be installed with::

    $ pip install pillow

    or browse https://pypi.python.org/pypi/Pillow/, pick the latest version and head to the bottom of the page for Windows binaries.


Usage examples
------------------

Tutorial
++++++++

.. toctree::

    tutorial

Cookbook
++++++++

.. toctree::

       usage

Read/write large files
++++++++++++++++++++++

.. toctree::

    optimized

API Documentation
------------------

.. toctree::

       api

Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`
