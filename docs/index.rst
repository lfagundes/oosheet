==========
OO, Sheet!
==========

OOSheet is a Python module for manipulating LibreOffice spreadsheet documents and creating macros.

Using Python, you interact with an LibreOffice instance to develop and test your code. When you're finished you can insert your python script inside the document to run it as macro, if this is what you desire.

OOSheet API was inspired by `JQuery <http://jquery.com/>`_. It uses selectors in same way that you would use in LibreOffice, with cascading method calls for quick development. Part of the API was also inspired by `Django's <http://djangoproject.com/>`_ object-relational mapper.

Why OOSheet?
============

When you use a lot of spreadsheets and have some programming skills, writing scripts to make your life easier are something that will come to your mind. If you start to look for your options with LibreOffice, you'll see that OO.org Basic is ugly and weird, while Python support is powerful but very tricky. And finally, your routines are likely to be coupled to the structure of your spreadsheet, so the best place to have your macros would be inside your documents, but OO.org provides no way for you to do that.

If you see yourself in need of automating routines in a spreadsheet and like Python, OOSheet is surely for you. 

If your situation is not really like this but you're considering using PyUno for anything, it's very likely that you'll find OOSheet useful in some way, even if your document is not a spreadsheet. The base class OODoc may be a good general wrapper for PyUno, and OOPacker class can be used to insert your python script in any OO.org document.

For using OOSheet you need a running instance of LibreOffice. If you just want to generate a document, for example, as a result of a web system in which user will download some automatically generated spreadsheet, then this module is probably not what you're looking for. It could be used though, if you're willing to manage a running LibreOffice process.

Download / Install
==================

Just type::

    $ pip install oosheet

You can get the `pip command here`_.  

.. _pip command here: http://pip.openplans.org/

You'll need git and python uno. If you use a Debian-based GNU/Linux distribution (like Ubuntu), you can do this with::

    $ sudo aptitude install python-uno

OOSheet was developed and tested on Python 2.6 and LibreOffice 3.3. It should work in other versions, though. If you try it in other environments, please report results to author.

Source
======

The OOSheet source can be downloaded as a tar.gz file from http://pypi.python.org/pypi/oosheet

Using `git <http://git-scm.com/>`_ you can clone the source from http://github.com/lfagundes/oosheet.git

OOSheet is free and open for usage under the `MIT license`_.

.. _MIT license: http://en.wikipedia.org/wiki/MIT_License



Documentation
=============

Contents:

.. toctree::
   :maxdepth: 2

   using-oosheet
   working-with-any-document-type
   macros
   issues

.. _oosheet-api:

API Reference
=============

.. toctree::
    :glob:
    
    api/oodoc
    api/oosheet
    api/oopacker

Contributing
============

Please submit `bugs and patches <http://github.org/lfagundes/issues/>`_, preferably with tests.  All contributors will be acknowledged.  Thanks!

Credits
=======

OOSheet was created by Luis Fagundes and sponsored by `hacklab/ <http://hacklab.com.br/>`_.

`Fudge <http://farmdev.com/projects/fudge/>`_ project also take credits for a good documentation structure, on which this one was based.

Oscar Garcia suggested the user selection feature, implemented in v0.9.4.

Thomas Lundqvist sent implementation of data_array property, implemented in v0.9.6.

Luc Jean sent patch that allows OOSheet to run on Windows Vista with default Python interpreter, v0.9.7

Changelog
=========

- 1.1

  - Much better performance when acessing via sockets
  - Development now based on LibreOffice instead of OpenOffice.org
  - [NEW] Objects can be accessed as arrays
  - [NEW] find() methods searches selection for matching cells
  - [NEW] oosheet-launch reminds that complicated launching command line

- 1.0

  - [NEW] Iterators for cells, rows and columns
  - [MOD] Refactor for working with all types of documents

- 0.9.7

  - [FIX] works with default Python interpreter in Windows Vista

- 0.9.6

  - [NEW] data_array property returns selection's values as 2d-tuple

- 0.9.5

  - [FIX] flatten() breaking fields with formulas based on zero values formatted as strings (ex: R$ 0,00)

- 0.9.4

  - [NEW] S() now gets user's selection
  - [FIX] Date format is not automatically set if current format is already a date

- 0.9.3.1

  - Fixes packaging problem

- 0.9.3

  - [NEW] flatten() method
  - [NEW] sheet and cell protection

- 0.9.2

  - [NEW] grow_DIRECTION_until and shrink_DIRECTION_until methods
  - [NEW] last_row, last_column, first_row, first_column properties
  - [FIX] api documentation error in shift_until()
  - [FIX] __repr__ of empty selector raised error

- 0.9.1

  - Documentation changed to include installation instructions with pip and link to website
  - Code is same as 0.9.0

- 0.9.0

  - first release

