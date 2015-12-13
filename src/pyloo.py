#!/usr/bin/python
# -*- coding: utf-8 -*-

"""
PyLOO - Python Libre/Open Office interface API (UNO)

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

###############################################################################
import uno
import unohelper

# Exceptions
from com.sun.star.uno import RuntimeException
from com.sun.star.lang import IllegalArgumentException
from com.sun.star.connection import NoConnectException
from com.sun.star.io import IOException

# Other office interfaces
from com.sun.star.table import CellRangeAddress
from com.sun.star.beans import PropertyValue

# Office eNums
from com.sun.star.table.CellContentType import TEXT, EMPTY, VALUE, FORMULA

###############################################################################
# Start LibreOffice/OpenOffice Calc in listening mode:
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
        """
        Constructor

        @type  fields: Fields
        @param fields: Fields object

        @type  name: string
        @param name: Field name
        """
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
                oSheets = self._fields.document().o_doc().getSheets()
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
        Checking if the field is null.

        @rtype:   bool
        @return:  Value insertion result
        """
        return self._is_null

    def set_value(self, value, column=0, row=0):
        """
        Set filed value at position Column/Row

        @type  value: string
        @param value: Cell value

        @type  column: int
        @param column: column index

        @type  row: int
        @param row: row index

        @rtype:   bool
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

        @type  column: int
        @param column: column index

        @type  row: int
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

    def insert_rows(self, num_rows=1, step=1, columns_to_copy=250):
        """
        Insert rows

        Insert new rows at the specified position relatively to cell.
        After the new row insertion the content of the current rows is copied
        to the new rows.

        @type  num_rows: int
        @param num_rows: Number of rows to insert. Default value=1

        @type  step: int
        @param step: Step of rows insertion.

        @type  columns_to_copy: bool
        @param columns_to_copy: Number of a columns to copy on insert.
                            No copy will be performed if columns_to_copy = 0

        @rtype:   bool
        @return:  Operation result
        """
        result = False
        if self._fields and self._oSheet and num_rows > 0:
            insert_pos_with_step = self._oCellAddress.Row + 1 + step
            self._oSheet.Rows.insertByIndex(
                insert_pos_with_step, (num_rows * step))

            # Copy rows
            if columns_to_copy > 0 and num_rows > 0:
                # Initialize variable as CellRangeAddress object
                oCellRangeAddress_Src = CellRangeAddress()

                # Source address
                oCellRangeAddress_Src.Sheet = self._oCellAddress.Sheet
                oCellRangeAddress_Src.StartColumn = 0
                oCellRangeAddress_Src.EndColumn = columns_to_copy
                oCellRangeAddress_Src.StartRow = \
                    self._oCellAddress.Row + 1
                oCellRangeAddress_Src.EndRow = \
                    oCellRangeAddress_Src.StartRow + step - 1

                # Destination address
                oCellAddress_Dst = self._oCellAddress
                oCellAddress_Dst.Column = 0
                oCellAddress_Dst.Row = oCellAddress_Dst.Row + 1 + step

                for i in range(0, num_rows):
                    self._oSheet.copyRange(oCellAddress_Dst,
                                           oCellRangeAddress_Src)
                    oCellAddress_Dst.Row += step

#                 Restore cell address variable
                self._oCellAddress = self._oRange.getReferencePosition()

            result = True
        return result

###############################################################################
###############################################################################
###############################################################################


class Fields:

    """
    Document fields.
    Search and manage fields (name ranges).
    """

    def __init__(self, document):
        """
        Constructor

        @type  document: Document
        @param document: Document object
        """
        self._document = document
        self._field = None
        self._is_null = True

#         LibreOffice variables.
        self._oNamedRanges = None
        if self._document:
            self._oNamedRanges = self._document.o_doc().NamedRanges
            self._is_null = False

    def is_null(self):
        """
        Checking if the fields (NamedRanges) object is initialized

        @rtype:   bool
        @return:  Fields object state
        """
        return self._is_null

    def count(self):
        """
        Get number of fields (named ranges) in the document.

        @rtype:   int
        @return:  the number of fields in the document.
        """
        count = 0
        if self._oNamedRanges:
            count = self._oNamedRanges.getCount()
        return count

    def document(self):
        """
        Get document object.

        @rtype:   Template
        @return:  Template object
        """
        return self._document

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

    def add(self, name, value, sheet, column, row, type=0):
        """
        Not implemented yet. FIXME
        Adds a new field (named range) to the collection.

        @type  name: string
        @param name: the new name of the named range.

        @type  value: string
        @param value: the formula expression.

        @type  sheet: int
        @param sheet: the formula expression.

        @type  column: int
        @param column: the formula expression.

        @type  row: int
        @param row: the formula expression.

        @type  type: int
        @param type: a combination of flags that specify the type of a named \
                    range, as defined in NamedRangeFlag. This parameter \
                    will be zero for any common named range.

        @rtype:   bool
        @return:  Operation result
        """
        cell_address = uno.createUnoStruct("com.sun.star.table.CellAddress")
        cell_address.Sheet = sheet
        cell_address.Column = column
        cell_address.Row = row
        if self._oNamedRanges:
            self._oNamedRanges.addNewByName(name, value, cell_address, 0)
        return None

    def remove(self, name):
        """
        Not implemented yet. FIXME
        Removes a field (named range) from the collection.

        @type  name: string
        @param name: the new name of the named range.

        @rtype:   bool
        @return:  Operation result
        """
        result = False
        if self._oNamedRanges:
            self._oNamedRanges.removeByName(name)
            result = True
        return result

###############################################################################
###############################################################################
###############################################################################


class Sheet:

    """
    Document sheet.
    Manage sheet and cells.
    """

    def __init__(self, sheets, index_or_name):
        """
        Constructor

        @type  sheets: Sheets
        @param sheets: Sheets object

        @type  index_or_name: int or string
        @param index_or_name: Sheet index or sheet name
        """
        self._sheets = sheets
        self._is_null = True

#         LibreOffice variables.
        self._oSheet = None

        if isinstance(index_or_name, int):
            # get by index
            self._oSheet = self._sheets.o_sheets().getByIndex(index_or_name)
            self._is_null = False
        else:
            # get by name
            self._oSheet = self._sheets.o_sheets().getByName(index_or_name)
            self._is_null = False

    def is_null(self):
        """
        Checking if the sheet is null.

        @rtype:   bool
        @return:  Value insertion result
        """
        return self._is_null

    def set_cell_value_by_index(self, value, col, row, is_formula=False):
        """
        Set cell value.

        @type  value: string
        @param value: Cell value

        @type  col: int
        @param col: Cell column index

        @type  row: int
        @param row: Cell row index

        @type  is_formula: bool
        @param is_formula: Not supported yet

        @rtype:   bool
        @return:  Operation result
        """
        result = False
        oCell = self._oSheet.getCellByPosition(col, row)
        if is_formula:
            oCell.setFormula(value)
            result = True
        elif isinstance(value, int)\
                or isinstance(value, float):
            oCell.setValue(value)
            result = True
        else:
            oCell.setString(value)
            result = True
        return result

    def cell_value_by_index(self, col, row):
        """
        Get cell value.

        @type  col: int
        @param col: Cell column index

        @type  row: int
        @param row: Cell row index

        @rtype:   long, int, float or string
        @return:  Value. Value type depends on document cell value.
        """
        value = None
        oCell = self._oSheet.getCellByPosition(col, row)
        value_type = oCell.getType()
        if VALUE == value_type:
            value = oCell.getValue()
        if FORMULA == value_type:
            value = oCell.getFormula()
        if TEXT == value_type:
            value = oCell.getString()

        return value

###############################################################################
###############################################################################
###############################################################################


class Sheets:

    """
    Document sheets.
    Search and manage sheets.
    """

    def __init__(self, document):
        """
        Constructor

        @type  document: Document
        @param document: Document object
        """
        self._document = document
        self._sheet = None

#         LibreOffice variables.
        self._oSheets = None
        if self._document:
            self._oSheets = self._document.o_doc().getSheets()

    def is_null(self):
        """
        Checking if the document object is initialized

        @rtype:   bool
        @return:  Document object state
        """
        result = False
        if self._oSheets is None:
            result = True
        return result

    def sheet(self, index_or_name):
        """
        Get sheet by index or name.

        @type  value: int, string
        @param value: Sheet index or name

        @rtype:   Sheet
        @return:  Sheet object
        """
        return Sheet(self, index_or_name)

    def count(self):
        """
        Get number of sheets in document.

        @rtype:   int
        @return:  the number of sheets in document.
        """
        count = 0
        if self._oSheets:
            count = self._oSheets.getCount()
        return count

    def o_sheets(self):
        """
        LibreOffice/OpenOffice Calc Spreadsheets object.

        Required for Fields and Field classes. Not recommended use it directly.

        @rtype:   com::sun::star::sheet::XSpreadsheets
        @return:  Libre/Open office Spreadsheets object
        """
        return self._oSheets

    def insert_spreadsheet(self, name, index):
        """
        Inserts a new sheet into the collection.

        @type  name: string
        @param name: The name of the new spreadsheet.

        @type  index: int
        @param index: The index of the new spreadsheet in the collection.

        @rtype:   bool
        @return:  Operation result
        """
        result = False
        if self.o_sheets():
            self.o_sheets().insertNewByName(name, index)
            result = True
        return result

    def remove_spreadsheet(self, name):
        """
        Inserts a new sheet into the collection.

        @type  name: string
        @param name: The name of the removing spreadsheet.

        @rtype:   bool
        @return:  Operation result
        """
        result = False
        if self.o_sheets():
            self.o_sheets().removeByName(name)
            result = True
        return result

###############################################################################
###############################################################################
###############################################################################


class Document:

    def __init__(self, connection_string="\
uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"):
        """
        Constructor

        @type  connection_string: string
        @param connection_string: Libre/Open office initialization string
        """
        self._sheets = None
        self._fields = None
        self._connection_string = connection_string

#         LibreOffice variables.
        self._oResolver = None
        self._oContext = None
        self._oDesktop = None
        self._oDoc = None
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
                raise (e)
            except IllegalArgumentException as e:
                raise (e)
            except RuntimeException as e:
                raise (e)

    def is_null(self):
        """
        Checking if the document object is initialized

        @rtype:   bool
        @return:  Document object state
        """
        result = False
        if self._oDesktop is None:
            result = True
        return result

    def _toProperties(self, **args):
        """
        Converts '**args' arguments to the tuple of 'PropertyValue's

        @rtype:   tuple
        @return:  Fields object
        """
        props = []
        for key in args:
            prop = PropertyValue()
            prop.Name = key
            prop.Value = args[key]
            props.append(prop)
        return tuple(props)

    def o_doc(self):
        """
        LibreOffice/OpenOffice Calc document object.

        Required for Fields and Field classes. Not recommended use it directly.

        @rtype:   com::sun::star::lang::XComponent
        @return:  Libre/Open office document object
        """
        return self._oDoc

    def _open_document(self, doc_name=""):
        """
        Open document.

        @type  doc_name: string
        @param doc_name: Document name.

        @rtype:   bool
        @return:  Operation result
        """
        result = False
        if self._oDesktop:
            try:
                self._oDoc = self._oDesktop.loadComponentFromURL(
                    doc_name, "_blank", 0, ())
                result = True
            except IllegalArgumentException as e:
                raise (e)
            except IOException as e:
                raise (e)
        return result

    def new_document(self):
        """
        Create new document.

        @rtype:   bool
        @return:  Operation result
        """
        return self._open_document("private:factory/scalc")

    def open_document(self, doc_name):
        """
        Open document.

        @type  doc_name: string
        @param doc_name: Document name.

        @rtype:   bool
        @return:  Operation result
        """
        result = False
        if len(doc_name) > 0:
            doc_name = unohelper.systemPathToFileUrl(doc_name)
            result = self._open_document(doc_name)
        return result

    def save_document(self, doc_name="", filter_name=""):
        """
        Save document.

        @type  doc_name: string
        @param doc_name: Document name. If no document name defined the current
                        name is used.

        @type  filter_name: string
        @param filter_name: file type:
                            # ods=""
                            # pdf="calc_pdf_Export"
                            # csv="Text - txt - csv (StarCalc)"
                            # xls="calc_MS_Excel_40"
                            # xlsx="Calc Office Open XML"

        @rtype:   bool
        @return:  Operation result
        """
        result = False
        if self.o_doc():
            if 0 == len(doc_name):
                self.o_doc().store()
            else:
                full_file_name = unohelper.systemPathToFileUrl(doc_name)
                try:
                    self.o_doc().storeToURL(
                                full_file_name,
                                self._toProperties(FilterName=filter_name))
                    result = True
                except IllegalArgumentException as e:
                    raise (e)
                except ErrorCodeIOException as e:
                    raise (e)
                except IOException as e:
                    raise (e)
        return result

    def close_document(self):
        """
        Close document.

        Close current document.

        @rtype:   bool
        @return:  Operation result
        """
        result = False
        try:
            if self.o_doc():
                self.o_doc().close(True)
                self._oDoc = None
                result = True
        except ErrorCodeIOException as e:
            raise (e)
        except IOException as e:
            raise (e)
        return result

    def sheets(self):
        """
        Get Sheets document's object.

        @rtype:   Sheets
        @return:  Sheets object
        """

        if self._sheets is None:
            self._sheets = Sheets(self)
        return self._sheets

    def fields(self):
        """
        Get Fields document's object.

        @rtype:   Fields
        @return:  Fields object
        """
        if self._fields is None:
            self._fields = Fields(self)
        return self._fields

    def version(self):
        """
        Get library version.

        @rtype:   string
        @return:  PyLibra version
        """
        return "0.0.1"
