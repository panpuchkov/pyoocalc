# -*- coding: utf-8 -*-

"""
PyOOCalc - Python Libre/Open Office Calc interface API (UNO)

Requirements for Ubuntu users:

sudo apt-get install python-uno

Optional requirements:

sudo apt-get install libreoffice-script-provider-python

Copyright (c) 2015

@author: Yurii Puchkov
@organization: http://arilot.com/
@license: GPL v3
@contact: panpuchkov@gmail.com
"""

import logging
import os
import sys

sys.path.append('./../')
import pyoocalc

logging.basicConfig(stream=sys.stdout,
                    level=logging.INFO,
                    format='%(message)s')
logger = logging.getLogger("example")

# open document
doc = None
try:
    doc = pyoocalc.Document(autostart=True)
except OSError as e:
    logger.error("ERROR: {0} - {1}", e.errno, e.strerror)
except pyoocalc.NoConnectException as e:
    logger.error("ERROR: The OpenOffice.org process is not started or does \
not listen on the resource.\n\
{0}\n\n\
Start LibreOffice/OpenOffice in listening mode, \
example:\n\
    soffice \
--accept=\"socket,host=localhost,port=2002;urp;\"\n".format(e.Message))

if doc:
    # Get PyOOCalc version
    logger.info("PyOOCalc version: {0}".format(doc.version))

    # Open document
    doc.open_document(os.getcwd() + "/example.ods")

    # Get document fields
    fields = doc.fields
    logger.info("Fields count: {0}".format(fields.count))

    # Get field "HEADER"
    field = fields.field("HEADER")
    logger.info("Document header is: {0}".format(field.is_null))

    # Set values
    field = fields.field("TABLE_NAME")
    field.set_value("Test table name")
    logger.info("New table name is: {0}".format(field.value()))

    ########################################
    # Get table column fields
    field1 = fields.field("FIELD_1")
    field2 = fields.field("FIELD_2")
    field3 = fields.field("FIELD_3")
    field4 = fields.field("FIELD_4")

    # Set number of rows and step
    num_rows = 5
    step = 2
    # Insert rows into the table
    if num_rows > 0:
        field1.insert_rows(num_rows=num_rows-1, step=step, columns_to_copy=200)
    # Insert data into the cells by field name and offset
    for i in range(1, num_rows + 1):
        field1.set_value("F1.{0}".format(str(i)), 0, i * step - (step - 1))
        field2.set_value("F2.{0}".format(str(i)), 0, i * step - (step - 1))
        field3.set_value("F3.{0}".format(str(i)), 0, i * step - (step - 1))
        field4.set_value("F4.{0}".format(str(i)), 0, i * step - (step - 1))

    # Insert and remove spreadsheets
    doc.sheets.insert_spreadsheet("Test1", 0)
    doc.sheets.insert_spreadsheet("Test2", 2)
    doc.sheets.remove_spreadsheet("Test2")

    # Get sheet by index and set and get cell value
    sheet = doc.sheets.sheet(0)
    sheet.set_cell_value_by_index("value1", 1, 0,)
    logger.info("Cell 'value1' : {0}".format(sheet.cell_value_by_index(1, 0)))

    # Get sheet by name and set and get cell value
    sheet = doc.sheets.sheet("Test1")
    sheet.set_cell_value_by_index("value2", 0, 1,)
    logger.info("Cell 'value2' : {0}".format(sheet.cell_value_by_index(0, 1)))
