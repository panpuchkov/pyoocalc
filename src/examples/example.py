#!/usr/bin/python3 
# -*- coding: utf-8 -*- 

##
# sudo apt-get install python-uno
# sudo apt-get install libreoffice-script-provider-python
# 

# common modules
import sys 
import os
sys.path.append('./../')

# user modules
import pylibratm

# open document
# template = pylibratm.TemplateManager().template()
template = pylibratm.Template()
if template:
    file_name = os.getcwd() + "/example.ods"
    file_name_saved = os.getcwd() + "/example_saved.ods"
    template.open_document( file_name )

# Get document fields
fields = template.fields()

# Get field "HEADER"
field = fields.field("HEADER")
print ( "Document header is: " + str(field.is_null()) )

# Set values
field = fields.field("TABLE_NAME")
field.set_value("Test table name")
print ("New table name is: " + field.value())

# Insert rows
field1 = fields.field("FIELD_1")
field2 = fields.field("FIELD_2")
field3 = fields.field("FIELD_3")
field4 = fields.field("FIELD_4")

num_rows = 5
step = 2
if num_rows > 0:
    field1.insert_rows(num_rows=num_rows-1, step=step, columns_to_copy=200)
for i in range(1, num_rows + 1):
    field1.set_value("F1." + str(i), 0, i * step - (step -1))
    field2.set_value("F2." + str(i), 0, i * step - (step -1))
    field3.set_value("F3." + str(i), 0, i * step - (step -1))
    field4.set_value("F4." + str(i), 0, i * step - (step -1))

template.insert_spreadsheet("Test1", 0)
template.insert_spreadsheet("Test2", 2)
template.remove_spreadsheet("Test2")
