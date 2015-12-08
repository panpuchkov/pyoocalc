#!/usr/bin/python 
# -*- coding: utf-8 -*- 

from  pylibratm import *

################################################################################
# example
#template("/home/pub/test.ods")
#g_oFields = fields()
#field("ARILOT_GOODS_NAME")

################################################################################
# example
# import subprocess
# subprocess.Popen(["/usr/bin/libreoffice", '-accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager'], shell = False)

#oTemplate = ArilotTemplateManager().Template()

#oTemplate.openDocument("/home/pub/test.ods")
#oFields = oTemplate.fields()
#oField = oFields.field("ARILOT_GOODS_NAME")
#print oField.value()
#oField.setValue("abc")
#print oField.value()
#oField.insertRow(1, 1, 3, 1)
#oField.insertColumn(1, 1, 5, 0)
#oTemplate.saveDocument()



#exit(0)



################################################################################
################################################################################
################################################################################

class dictFields:
    m_oFieldDict = dict()

    def __init__(self):
        self.clear()
        
    def field(self, strName):
        global g_oField
        global g_oFields
        g_oField = self.m_oFieldDict.get(strName)
        if None == g_oField:
            g_oField = g_oFields.field(strName)
            if None != g_oField:
                self.m_oFieldDict[strName] = g_oField
        return g_oField
    
    def clear(self):
        self.m_oFieldDict.clear()

def template(strDocname):
    global g_oTemplate
    global g_oFields
    global g_oField
    global g_oFieldDict

    g_oTemplate = None
    g_oFields = None
    g_oField = None
    g_oFieldDict.clear()
    
    g_oTemplate = ArilotTemplateManager().Template()
    if g_oTemplate:
        g_oTemplate.openDocument(strDocname)
        g_oFields = g_oTemplate.fields()
    return g_oTemplate

def fields():
    global g_oTemplate
    global g_oFields
    g_oFields = None
    if None == g_oFields:
        g_oFields = g_oTemplate.fields()
    return g_oFields

def field(strName):
    global g_oFields
    global g_oField
    #global g_oFieldDict
    
    g_oField = None
    g_oField = g_oFields.field(strName)
    if g_oFields:
        g_oField = g_oFieldDict.field(strName)
    return g_oField

g_oTemplate = None
g_oFields = None
g_oField = None
g_oFieldDict = dictFields()

################################################################################

################################################################################
# example
# open document
oTemplate = ArilotTemplateManager().Template()
if oTemplate:
    oTemplate.openDocument("/home/pub/test.ods")



#print g_oFieldDict.field("ARILOT_GOODS_1")
#g_oFieldDict.field("ARILOT_GOODS_NAME")
#print g_oFieldDict.field("ARILOT_GOODS_1")
#print g_oFieldDict.field("ARILOT_GOODS_NAME")

#field("ARILOT_GOODS_NAME1")

#print field("ARILOT_GOODS_NAME").value(-1, 2)
#print field("ARILOT_GOODS_NAME").setValue("abc", 1, 0)
#print g_oFieldDict


#print field("ARILOT_GOODS_NAME").value(-1, 0)
#print field("ARILOT_GOODS_NAME").setValue(0, 1, "new value")
