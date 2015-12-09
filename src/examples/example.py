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
# oTemplate = pylibratm.TemplateManager().template()
oTemplate = pylibratm.Template()
if oTemplate:
    file_name = os.getcwd() + "/example.ods"
    oTemplate.open_document( file_name )

# Get document fields
oFields = oTemplate.fields()

# Get field "HEADER"
oField = oFields.field("HEADER")
print ( "Document header is: " + str(oField.is_null()) )

# Set values
oField = oFields.field("TABLE_NAME")
oField.set_value("Test table name")
print ("New table name is: " + oField.value())

# Insert rows
oField1 = oFields.field("FIELD_1")
# oField.insert_row(nRow = 1, nStep = 1, nNumColumns = 1, nOffset = 0)
oField1.set_value("F1.1", 0, 1)
oField1.set_value("F1.2", 0, 2)


oField4 = oFields.field("FIELD_4")
# oField.insert_row(nRow = 1, nStep = 1, nNumColumns = 1, nOffset = 0)
oField4.set_value("2", 0, 1)
oField4.set_value("3", 0, 2)

# oField = oFields.field("G1")
# oField.setValue("G1")
