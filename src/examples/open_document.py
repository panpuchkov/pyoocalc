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
