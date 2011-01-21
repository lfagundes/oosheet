==========
OO, Sheet!
==========

OOSheet is a Python module for manipulating OpenOffice.org spreadsheet documents and creating macros.

Using Python, you interact with an OpenOffice.org instance to develop and test your code. When you're finished you can insert your python script inside the document to run it as macro, if this is what you desire.

OOSheet API was inspired by `JQuery <http://jquery.com/>`_. It uses selectors in same way that you would use in OpenOffice.org, with cascading method calls for quick development. Part of the API was also inspired by `Django's <http://djangoproject.com/>`_ object-relational mapper.


Download / Install
==================

Just type::

    $ pip install oosheet

You can get the `pip command here`_.  OOSheet was tested on Python 2.6

.. _pip command here: http://pip.openplans.org/

.. _oosheet-source:

Source
======

The OOSheet source can be downloaded as tar.gz file from http://pypi.python.org/pypi/oosheet

Using `git <http://git-scm.com/>`_ you can clone the source from http://github.com/lfagundes/oosheet.git. 

OOSheet is free and open for usage under the `MIT license`_.

.. _MIT license: http://en.wikipedia.org/wiki/MIT_License

Contents
========

.. toctree::
   :maxdepth: 2

   using-oosheet
   document-macros
   other-oo-documents

.. _oosheet-api:

API Reference
=============

.. toctree::
    :glob:
    
    api/*

Contributing
============

Please submit `bugs and patches <http://github.org/lfagundes/issues/>`_, preferably with tests.  All contributors will be acknowledged.  Thanks!

Credits
=======

OOSheet was created by Luis Fagundes and sponsored by `hacklab/ <http://hacklab.com.br/>`_.

`Fudge <http://farmdev.com/projects/fudge/>`_ project also take credits for a good documentation structure, on which this one was based.

Changelog
=========

- 0.9.0

  - First release






