========================================================
PyLOO - Python Libre/Open Office interface API (UNO)
========================================================

PyLOO allows you to generate different report documents on the base of
prepared OpenOffice or LibreOffice templates. PyLOO is used in accounting
system for generating invoices and other documents.



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

This library is released under the GPL-V3 license. See the ``LICENSE`` file.

Copyright (c) 2015 thepurple.
