==========
OO, Sheet!
==========

OOSheet is a Python module for manipulating OpenOffice.org spreadsheet documents and creating macros.

Using Python, you interact with an OpenOffice.org instance to develop and test your code. When you're finished you can insert your python script inside the document to run it as macro, if this is what you desire.

OOSheet API was inspired by `JQuery <http://jquery.com/>`_. It uses selectors in same way that you would use in OpenOffice.org, with cascading method calls for quick development. Part of the API was also inspired by `Django's <http://djangoproject.com/>`_ object-relational mapper.

Why OOSheet?
============

This library was developed in the following scenario: there is a spreadsheet used to manage some business, and as the spreadsheet gets more complex, the need of some more automation raises, and some people might even start saying the word "ERP" around, but there's a big gap between that good naive spreadsheet and a software that would solve all the problems. When you consider macros, you'll see that OO.org Basic is ugly and weird, while Python support is powerful but very tricky. And finally, your routines are likely to be coupled to the structure of your spreadsheet, so the best place to have your macros would be inside your documents, but OO.org provides no way for you to do that.

If you see yourself in need of automation of routines of a spreadsheet, OOSheet is surely for you. 

If your situation is not really like this but you're considering using PyUno for anything, it's very likely that you'll find OOSheet useful in some way, even if your document is not a spreadsheet. The base class OODoc may be a good general wrapper for PyUno, and OOPacker class can be used to insert your python script in any OO.org document.

For using OOSheet you need a running instance of OpenOffice.org. If you just want to generate a document, for example, as a result of a web system in which user will download some automatically generated spreadsheet, then this module is probably not what you're looking for. It would make more sense to generate the document directly. It could be used though, if you're willing to manage a running OpenOffice.org process.

Download / Install
==================

For now, cloning the source from github is the only way. Using `git <http://git-scm.com/>`_:

    $ git clone http://github.com/lfagundes/oosheet.git
    $ cd oosheet
    $ python setup.py install

You'll need git and python uno. If you use a Debian-based GNU/Linux distribution (like Ubuntu), you can do this with:

    $ sudo aptitude install python-uno

OOSheet was developed and tested on Python 2.6 and OpenOffice.org 3.2. It should work in other versions, though. If you try it in other environments, please report results to author.

OOSheet is free and open for usage under the `MIT license`_.

.. _MIT license: http://en.wikipedia.org/wiki/MIT_License


Welcome to oosheet's documentation!
===================================

Contents:

.. toctree::
   :maxdepth: 2

   using-oosheet
   macros

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

No formal releases yet. You're a pioneer! But don't get intimidated, the code is stable and test coverage good.


Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`

