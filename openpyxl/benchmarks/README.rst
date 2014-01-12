Benchmarking
============


Purpose
-------

Openpyxl provides optimised readers and writers for dealing with large files.
It is important to know when these optimisations make sense and that they
continue to do so over time.

In addition, openpyxl supports different XML-backends for the parsing and
creation of XML. It is important that the fastest backend is always used and
provides measurable performance improvements.


Approach
--------

Sample files exist for benchmarking. Performance will vary considerably from
machine to machine, therefore, results should be normalised with the standard
workbook using cElementTree.
