#!/usr/bin/python3 
# -*- coding: utf-8 -*- 

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
print ( "Documetn header is: " + str(oField.isNull()) )

# Set values
oField = oFields.field("TABLE_NAME")
oField.setValue("Test table name")
print ("New table name is: " + oField.value())
