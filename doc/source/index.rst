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

How to Contribute Code
----------------------

Any help will be greatly appreciated, just follow those steps:

    1. Please start a new fork (https://bitbucket.org/ericgazoni/openpyxl/fork) for each independant feature, don't try to fix all problems at the same time, it's easier for those who will review and merge your changes ;-)
    2. Hack hack hack
    3. Don't forget to add unit tests for your changes ! (YES, even if it's a one-liner, or there is a high probability your work will not be taken into consideration). There are plenty of examples in the /test directory if you lack know-how or inspiration.
    4. If you added a whole new feature, or just improved something, you can be proud of it, so add yourself to the AUTHORS file :-)
    5. Let people know about the shiny thing you just implemented, update the docs !
    6. When it's done, just issue a pull request (click on the large "pull request" button on *your* repository) and wait for your code to be reviewed, and, if you followed all theses steps, merged into the main repository.

.. note:

This is an open-source project, maintained by volunteers on their spare time, so while we try to work on this project as often as possible, sometimes life gets in the way. Please be patient.

Other ways to help
------------------

There are several ways to contribute, even if you can't code (or can't code well):

- triaging bugs on the bug tracker: closing bugs that have already been closed, are not relevant, cannot be reproduced, ...
- updating documentation in virtually every area: many large features have been added (mainly about charts and images at the moment) but without any documentation, it's pretty hard to do anything with it
- proposing compatibility fixes for different versions of Python: we try to support 2.5 to 3.3, so if it does not work on your environment, let us know :-)

Installation
------------

The best method to install openpyxl is using a PyPi client such as easy_install (setuptools) or pip::

    $ pip install openpyxl

or ::

    $ easy_install install openpyxl

.. note::

    To install from sources (there is nothing to build, openpyxl is 100% pure Python), you can download an archive from `bitbucket`_ (look in the "tags" tab).

    After extracting the archive, you can do::

    $ python setup.py install

.. _bitbucket: https://bitbucket.org/ericgazoni/openpyxl/downloads

.. warning::

    To be able to include images (jpeg,png,bmp,...) into an openpyxl file, you will also need the 'PIL' library that can be installed with::

    $ pip install pillow

    or browse https://pypi.python.org/pypi/Pillow/, pick the latest version and head to the bottom of the page for Windows binaries.


Getting the source
------------------

Source code is hosted on bitbucket.org. You can get it using a Mercurial client and the following URLs:

    * $ hg clone \https://bitbucket.org/ericgazoni/openpyxl -r |release|

or to get the latest development version:

    * $ hg clone \https://bitbucket.org/ericgazoni/openpyxl


Usage examples
--------------

Tutorial
++++++++

.. toctree::

    tutorial

Cookbook
++++++++

.. toctree::

    usage

Charts
++++++

.. toctree::

    charts

Read/write large files
++++++++++++++++++++++

.. toctree::

    optimized

Working with styles
+++++++++++++++++++

.. toctree::

    styles

API Documentation
------------------

.. toctree::

       api

Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`
