#!/usr/bin/python3 
# -*- coding: utf-8 -*- 

##
# sudo apt-get install python-uno
# sudo apt-get install libreoffice-script-provider-python
# 

# common modules
import sys, os
sys.path.append('./../')

# user modules
import pylibratm

# open document
oTemplate = pylibratm.TemplateManager().Template()
if oTemplate:
    file_name = os.getcwd() + "/example.ods"
    oTemplate.openDocument( file_name )

# Get document fields
oFields = oTemplate.fields()

# Get field "HEADER"
oField = oFields.field("HEADER")
print ( "Document header is: " + str(oField.isNull()) )

# Set values
oField = oFields.field("TABLE_NAME")
oField.setValue("Test table name")
print ("New table name is: " + oField.value())

# Insert rows
oField = oFields.field("FIELD_1")
oField.insertRow(nRow = 1, nStep = 1, nNumColumns = 1, nOffset = 0)
oField.setValue("F1.1", 0, 1)
oField.setValue("F1.2", 0, 2)


# oField = oFields.field("G1")
# oField.setValue("G1")
