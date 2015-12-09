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

import uno

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
        self._fields = None
        self._is_null = True
        self._oSheet = None
        self._oNamedRanges = None
        self._oRange = None
        self._oCell = None
        self._oCellAddress = None

        self._fields = fields

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
        """
        return self._fields

    def is_null(self):
        """
        Checking if a field is null.
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

    def insert_row(self, row=1, step=1, num_columns=1, offset=0):
        """
        Insert rows

        Insert new rows at the specified position relatively to cell.
        After the new row insertion the content of the current rows is copied
        to the new rows.

        @type  row: integer
        @param row: Row index. Default value=1.

        @type  step: integer
        @param step: Default value = 1.    // FIXME

        @type  num_columns: integer
        @param num_columns: Number of columns to copy. Default value=1

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
                self._oSheet.Rows.insertByIndex(
                    self._oCellAddress.Row + offset, num_columns)
        else:
            result = False
        return result

    def insert_column(self, column, step=1, num_rows=1, offset=0):
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
        self._field = None
        self._template = None

        self._oSheets = None
        self._oNamedRanges = None

        self._template = None
        self.attach_template(template)
        if self._template:
            self._oNamedRanges = self._template.document().NamedRanges

    def template(self):
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
        result = False
        # FIXME, Not implemented yet.
        return result

    def add(self, name):
        result = False
        # FIXME, Not implemented yet.
        return result

    def attach_template(self, oTemplate):
        self._template = oTemplate

    def count(self):
        nCount = -1
        # FIXME, Not implemented yet.
        return nCount

###############################################################################
###############################################################################
###############################################################################


class Template:

    def __init__(self, connection_string="\
uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"):
        self._fields = None

        self._oLocal = None
        self._oResolver = None
        self._oContext = None
        self._oDesktop = None
        self._oDocument = None

        self._oLocal = uno.getComponentContext()
        if self._oLocal:
            self._oResolver = \
                self._oLocal.ServiceManager.createInstanceWithContext(
                    "com.sun.star.bridge.UnoUrlResolver", self._oLocal)
            if self._oResolver:
                self._oContext = self._oResolver.resolve(connection_string)
                if self._oContext:
                    self._oDesktop = self\
                        ._oContext.ServiceManager.createInstanceWithContext(
                            "com.sun.star.frame.Desktop", self._oContext)

    def document(self):
        return self._oDocument

    def fields(self):
        """
        Get fields document's object.
        """
        self._fields = Fields(self)
        return self._fields

    def version(self):
        return "0.0.1"

    def close_document(self):
        """
        Close document.

        Close current document.
        Not implemented yet.
        """
        result = False
        return result

    def save_document(self, doc_name=""):
        result = True
        if self._oDocument:
            if 0 == len(doc_name):
                self._oDocument.store()
            else:
                strFullFileName = "file://" + doc_name
                self._oDocument.storeToURL(strFullFileName)
        else:
            result = False
        return result

    def open_document(self, doc_name):
        result = True
        if len(doc_name) > 0 and self._oDesktop:
            strFullFileName = "file://" + doc_name
            self._oDocument = self._oDesktop.loadComponentFromURL(
                strFullFileName, "_blank", 0, ())
        else:
            result = False
        return result

    def new_document(self, doc_name):
        """
        Create new document.

        Create new document.
        Not implemented yet.
        """
        result = False
        return result
