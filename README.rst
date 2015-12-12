========================================================
PyLOO - Python Libre/Open Office interface API (UNO)
========================================================

Description
-----------
There are a lot of python libraries for dealing with Libre/OpenOffice via 
API (UNO). As for me, one of the most interesting projects is pyoo. It supports 
a lot of feature from open/save documents up to cell merging and working with 
charts and diagrams. But none of them implements number of functions wich 
I need.

I have to generate different documents for number of projects such as 
accounting system, estate management, document circulation and others. The 
easiest way is to use standard office software. End users can create their own 
template without great efforts.

It's easy to create template but how to know where I have to insert data into 
the template. I can use ``Cell`` indexes (column, row) or name, 
example: ``E5``. 

But what if I want to use more then one template for report. For example one 
report for landscape format another for portrait. There is no warranty that 
I have to set the same value into the same Cell in different templates. 
So I have to store rules for different templates somewhere. I have found 
easier way and that is ``NamedRange``.

``NamedRange`` is a name for a cell or cell range on a sheet. ``NamedRange`` 
is unique for entire document. 

One more unfound feature is inserting rows. Any report or invoice has table 
with header and footprint. So I need to insert rows into table area and keep 
row format (font, cell merging etc).

Main features:
.............
  * Opening and creation of spreadsheet documents
  * Saving documents to all formats available in OpenOffice
  * Insert remove sheets
  * Insert rows
  * Set/get value by NamedRange
  * Set/get value by Cell address or name

You can find an example of the document with NamedRanges and how to work 
with it in the examples folder.


Requirements
------------

PyLOO runs on Python 3.

The only dependency is the Python-UNO library (imported as a module uno).
It is often installed with the office suite. On Debian based systems it can be
installed as python-uno or python3-uno package.

Obviously you will also need OpenOffice or LibreOffice Calc. On Debian systems
it is available as libreoffice-calc package.



Install
-------

You can copy the pyloo.py file somewhere to your ``PYTHONPATH``.



Usage
-----


Starting OpenOffice / LibreOffice

PyLOO requires a running OpenOffice or LibreOffice instance which it can
connect to. On Ubuntu you can start LibreOffice from a command line using a 
command similar to: ::

$ soffice --accept="socket,host=localhost,port=2002;urp;" --norestore --nologo --nodefault # --headless

The LibreOffice will be listening for localhost connection on port 2002.
Alternatively a named pipe can be used: ::

$ soffice --accept="pipe,name=hello;urp;" --norestore --nologo --nodefault # --headless

If the --headless option is used then no user interface is visible even when a
document is opened.

For more information run: ::

$ soffice --help

It is recommended to start directly the soffice binary. There can be various 
scripts (called for example libreoffice) which will run the soffice binary but 
you may not get the correct PID of the running program.



Documentation
-------------

You can find documentation here: ::

./doc/index.html

Examples: ::

 ./src/examples/example.py



Testing
-------

Automated integration tests cover most of the code.

The test suite assumes that OpenOffice or LibreOffice is running and it is 
listening on localhost port 2002.

All tests are in the test.py file: ::

$ python3 example.py



License
-------

This library is released under the ``GPL-V3`` license. 
See the ``LICENSE`` file.

Copyright (c) 2015.
