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
    file_name = os.getcwd() + "/example2.ods"
    template.open_document( file_name )

#     file_name1 = os.getcwd() + "/example1.ods"
#     template.new_document()
#     template.save_document(file_name1, )

exit()

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
# field1.insert_rows(1)
# field1.insert_rows(1)
# field1.insert_rows(1)
# field1.set_value("F1.1", 0, 1)
# field1.set_value("F1.2", 0, 2)

if template:
    print (template)
    template.save_document()
    template.close_document()
