#!/usr/bin/python
# -*- coding: utf-8 -*-

"""
Python library for operating with LibreOffice/OpenOffice.org Calc.

Requirements for Ubuntu users:

sudo apt-get install python-uno

Optional requirements:

sudo apt-get install libreoffice-script-provider-python

@author: Yurii Puchkov
@organization: http://arilot.com/
@license: GPL v3.0
@contact: panpuchkov@gmail.com
"""

###############################################################################
import uno
import unohelper

from com.sun.star.uno import RuntimeException
from com.sun.star.lang import IllegalArgumentException
from com.sun.star.connection import NoConnectException

###############################################################################

__version__ = "$Revision$"
# $Source$

###############################################################################
# example
# /usr/bin/libreoffice \
# -accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager"

###############################################################################
###############################################################################
###############################################################################


class Field:

    """
    Document field.

    Operations on the existing field.
    """

    def __init__(self, fields, name):
        self._fields = fields
        self._is_null = True

#         LibreOffice variables.
        self._oSheet = None
        self._oNamedRanges = None
        self._oRange = None
        self._oCell = None
        self._oCellAddress = None

        if self._fields:
            if self._fields._oNamedRanges.hasByName(name):
                self._oRange = self._fields._oNamedRanges.getByName(name)
                self._oCellAddress = self._oRange.getReferencePosition()
                oSheets = self._fields.template().document().getSheets()
                self._oSheet = oSheets.getByIndex(self._oCellAddress.Sheet)
                self._is_null = False
            else:
                self._oRange = None
                self._is_null = True

    def fields(self):
        """
        Get fields object.

        @rtype:   Fields
        @return:  Fields object
        """
        return self._fields

    def is_null(self):
        """
        Checking if a field is null.

        @rtype:   boolean
        @return:  Value insertion result
        """
        return self._is_null

    def set_value(self, value, column=0, row=0):
        """
        Set filed value at position Column/Row

        @type  value: string
        @param value: Cell value

        @type  column: integer
        @param column: column index

        @type  row: integer
        @param row: row index

        @rtype:   boolean
        @return:  Value insertion result
        """
        result = True
        if self._oRange:
            self._oCell = self._oSheet.getCellByPosition(
                self._oCellAddress.Column + column,
                self._oCellAddress.Row + row)
            if self._oCell:
                self._oCell.setString(value)
        else:
            result = False
        return result

    def value(self, column=0, row=0):
        """
        Get filed value at position Column/Row

        @type  column: integer
        @param column: column index

        @type  row: integer
        @param row: row index

        @rtype:   string
        @return: Document cell value in string format. Regardless of document\
                    cell type.
        """
        value = ""
        if self._oRange:
            self._oCell = self._oSheet.getCellByPosition(
                self._oCellAddress.Column + column,
                self._oCellAddress.Row + row)
            if self._oCell:
                value = self._oCell.getString()
        return value

    def insert_rows(self, row=0, step=1, num_columns=1, offset=0):
        """
        Insert rows

        Insert new rows at the specified position relatively to cell.
        After the new row insertion the content of the current rows is copied
        to the new rows.

        @type  row: integer
        @param row: Row index relative to current Field. Default value=0.

        @type  step: integer
        @param step: Step of rows insertion.

        @type  num_columns: integer
        @param num_columns: Number of columns to insert. Default value=1

        @type  offset: integer
        @param offset: Rows offset. Relatively to current field.
                        Default value=0

        @rtype:   boolean
        @return:  Operation result
        """
        result = False
#         if self._oCell and self._fields:
        if self._fields:
            # oCellAddress = self._oCell.CellAddress
            # oSheets = self._fields.template().document().getSheets()
            # oSheet = oSheets.getByIndex(oCellAddress.Sheet)
            if self._oSheet:
                self._oSheet.Rows.insertByIndex(
                    self._oCellAddress.Row + row, num_columns)
# self._oCellAddress.Row + row + step + offset, num_columns)

                self._oCellAddress = self._oRange.getReferencePosition()
                self._oCellAddress.Column = 0
                print (self._oCellAddress.Column)
                print (self._oCellAddress.Row)

                oCellAddress_src = CellRangeAddress()
#                 oCellAddress_Source = self._oCellAddress.Sheet()

#                 self._oSheet.copyRange(self._oCellAddress,
#                                             oCellAddress_Source)

#                 CellAddress cAddress;
#                 cAddress.Sheet = m_cAddress.Sheet;
#                 cAddress.Column = 0;
#                 cAddress.Row = crAddress.StartRow;
#                 crAddress.StartRow -= nStep;
#                 crAddress.EndRow -= nStep;
#                 xCellRangeMovement->copyRange(cAddress, crAddress);
#                 xCellRangeMovement->copyRange(cAddress, crAddress);
            result = True
        return result

    def insert_columns(self, column=0, step=1, num_rows=1, offset=0):
        """
        Insert rows

        Insert new column at the specified position relatively to cell.
        After the new column insertion the content of the current columns is
        copied to the new columns.

        @type  column: integer
        @param column: Row index. Default value=1.

        @type  step: integer
        @param step: Default value = 1.    // FIXME

        @type  num_rows: integer
        @param num_rows: Number of columns to copy. Default value=1

        @type  offset: integer
        @param offset: Rows offset. Relatively to current field.
                        Default value=0

        @rtype:   boolean
        @return:  Operation result
        """
        result = True
        if self._oCell and self._fields:
            # oCellAddress = self._oCell.CellAddress
            # oSheets = self._fields.template().document().getSheets()
            # oSheet = oSheets.getByIndex(oCellAddress.Sheet)
            if self._oSheet:
                self._oSheet.Columns.insertByIndex(
                    self._oCellAddress.Column + offset, num_rows)
        else:
            result = False
        return result

    def remove(self):
        """
        Remove field name.

        Remove field name from the document.
        Not implemented yet.
        """
        return None

###############################################################################
###############################################################################
###############################################################################


class Fields:

    """
    Document fields.
    Search and manage fields.
    """

    def __init__(self, template):
        self._template = template
        self._field = None

#         LibreOffice variables.
        self._oSheets = None
        self._oNamedRanges = None

        if self._template:
            self._oNamedRanges = self._template.document().NamedRanges

    def template(self):
        """
        Get template object.

        @rtype:   Template
        @return:  Template object
        """
        return self._template

    def field(self, name):
        """
        Get document field by name

        @type  name: string
        @param name: Field name

        @rtype:   Field object
        @return:  Field object
        """
        self._field = Field(self, name)
        if self._field.is_null() is None:
            self._field = None
        return self._field

    def insert_spreadsheet(self, name, index):
        """
        Insert spreadsheet.

        Not implemented yet.
        """
        result = False
        return result

    def add(self, name):
        """
        Get field.

        Not implemented yet.
        """
        return None

    def count(self):
        """
        Get fields count.

        Not implemented yet.
        """
        count = -1
        return count

###############################################################################
###############################################################################
###############################################################################


class Template:

    def __init__(self, connection_string="\
uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"):
        self._fields = None
        self._connection_string = connection_string

#         LibreOffice variables.
        self._oResolver = None
        self._oContext = None
        self._oDesktop = None
        self._oDocument = None
        self._oLocal = uno.getComponentContext()

        if self._oLocal:
            self._oResolver = \
                self._oLocal.ServiceManager.createInstanceWithContext(
                    "com.sun.star.bridge.UnoUrlResolver", self._oLocal)
            try:
                if self._oResolver:
                    self._oContext = self._oResolver.resolve(
                        self._connection_string)
                    self._oDesktop = self._oContext.ServiceManager.\
                        createInstanceWithContext(
                                            "com.sun.star.frame.Desktop",
                                            self._oContext)
            except NoConnectException as e:
                print ("The OpenOffice.org process is not started or does not\
listen on the resource (" + e.Message + ")")
            except IllegalArgumentException as e:
                print ("The url is invalid ( " + e.Message + ")")
            except RuntimeException as e:
                print ("An unknown error occurred: " + e.Message)
            except:
                print ("Unknown exception")

    def document(self):
        """
        LibreOffice/OpenOffice Calc document object.

        Required for Fileds and Field classes. Do not use it directly.
        """
        return self._oDocument

    def fields(self):
        """
        Get fields document's object.

        @rtype:   Fields
        @return:  Fields object
        """
        self._fields = Fields(self)
        return self._fields

    def version(self):
        """
        Get library version.

        @rtype:   string
        @return:  PyLibra version
        """
        return "0.0.1"

    def close_document(self):
        """
        Close document.

        Close current document.

        @rtype:   boolean
        @return:  Operation result
        """
        result = False
        try:
            if self._oDocument:
                self._oDocument.close(True)
                result = True
        except:
            print ("Unknown exception")
        return result

    def save_document(self, doc_name=""):
        """
        Save document.

        @type  doc_name: string
        @param doc_name: Document name. If no document name defined the current
                        name is used.

        @rtype:   boolean
        @return:  Operation result
        """
        result = False
        if self._oDocument:
            if 0 == len(doc_name):
                self._oDocument.store()
            else:
                full_file_name = unohelper.systemPathToFileUrl(doc_name)
                try:
                    self._oDocument.storeToURL(full_file_name)
                    result = True
                except IllegalArgumentException as e:
                    print ("The url (" + full_file_name + ") "
                           "is invalid ( " + e.Message + ")")
                except:
                    print ("Unknown exception")
        return result

    def open_document(self, doc_name):
        """
        Open document.

        @type  doc_name: string
        @param doc_name: Document name.

        @rtype:   boolean
        @return:  Operation result
        """
        result = False
        if len(doc_name) > 0 and self._oDesktop:
            full_file_name = unohelper.systemPathToFileUrl(doc_name)
            try:
                self._oDocument = self._oDesktop.loadComponentFromURL(
                    full_file_name, "_blank", 0, ())
                result = True
            except IllegalArgumentException as e:
                print (e)
            except:
                print ("Unknown exception")
        return result

    def new_document(self):
        """
        Create new document.

        Create new document.
        Not implemented yet.
        """
        result = False
        try:
            if self._oDesktop:
                self._oDocument = self._oDesktop.loadComponentFromURL(
                                    "private:factory/scalc", "_blank", 0, ())
                result = True
        except:
            print ("Unknown exception")
        return result
