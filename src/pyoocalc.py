#!/usr/bin/python
# -*- coding: utf-8 -*-

"""
PyOOCalc - Python LibreOffice/OpenOffice Calc interface API (UNO)

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

import os
import subprocess
import time

# Exceptions
from com.sun.star.uno import RuntimeException
from com.sun.star.lang import IllegalArgumentException, DisposedException
from com.sun.star.connection import NoConnectException
from com.sun.star.io import IOException

# Other office interfaces
from com.sun.star.table import CellRangeAddress
from com.sun.star.beans import PropertyValue

# Office eNums
from com.sun.star.table.CellContentType import TEXT, EMPTY, VALUE, FORMULA


###############################################################################
__version__ = "0.0.5"
_MSG_EXCEPT_SIDE_EFFECT = "Assigning a value to the '{0}' is not allowed."

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

        # LibreOffice variables.
        self._oSheet = None
        self._oNamedRanges = None
        self._oRange = None
        self._oCell = None
        self._oCellAddress = None

        if self._fields:
            if 0 == len(name):
                raise ValueError("'name' is an empty string")
            if self._fields._oNamedRanges.hasByName(name):
                self._oRange = self._fields._oNamedRanges.getByName(name)
                self._oCellAddress = self._oRange.getReferencePosition()
                oSheets = self._fields.document.o_doc.getSheets()
                self._oSheet = oSheets.getByIndex(self._oCellAddress.Sheet)
                self._is_null = False
            else:
                self._oRange = None
                self._is_null = True
        else:
            raise ValueError("'fields' value is None")

    @property
    def is_null(self):
        """
        Checking if the field is null.

        @rtype:   bool
        @return:  Value insertion result
        """
        return self._is_null

    @is_null.setter
    def is_null(self, value):
        """
        Side-effect protection. Raising an exception 'ValueError'.
        """
        raise ValueError(_MSG_EXCEPT_SIDE_EFFECT.format("is_null"))

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
        if num_rows <= 0:
            raise ValueError("'num_rows' must be a positive number")
        if columns_to_copy <= 0:
            raise ValueError("'columns_to_copy' must be a positive number")
        if 0 == step:
            raise ValueError("'step' must not be equal to Zero")

        if self._fields and self._oSheet:
            insert_pos_with_step = self._oCellAddress.Row + 1 + step
            self._oSheet.Rows.insertByIndex(
                insert_pos_with_step, (num_rows * step))

            # Copy rows
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

            # Restore cell address variable
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

        # LibreOffice variables.
        self._oNamedRanges = None
        if self._document:
            self._oNamedRanges = self._document.o_doc.NamedRanges
            self._is_null = False
        else:
            raise ValueError("'document' value is None")

    @property
    def is_null(self):
        """
        Checking if the fields (NamedRanges) object is initialized

        @rtype:   bool
        @return:  Fields object state
        """
        return self._is_null

    @is_null.setter
    def is_null(self, value):
        """
        Side-effect protection. Raising an exception 'ValueError'.
        """
        raise ValueError(_MSG_EXCEPT_SIDE_EFFECT.format("is_null"))

    @property
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

    @count.setter
    def count(self, value):
        """
        Side-effect protection. Raising an exception 'ValueError'.
        """
        raise ValueError(_MSG_EXCEPT_SIDE_EFFECT.format("count"))

    @property
    def document(self):
        """
        Get document object.

        @rtype:   Template
        @return:  Template object
        """
        return self._document

    @document.setter
    def document(self, value):
        """
        Side-effect protection. Raising an exception 'ValueError'.
        """
        raise ValueError(_MSG_EXCEPT_SIDE_EFFECT.format("document"))

    def field(self, name):
        """
        Get document field by name

        @type  name: string
        @param name: Field name

        @rtype:   Field object
        @return:  Field object
        """
        self._field = Field(self, name)
        if self._field.is_null is None:
            self._field = None
        return self._field

    def add(self, name, value, sheet, column, row):
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

        @rtype:   bool
        @return:  Operation result
        """
        if 0 == len(name):
            raise ValueError("'name' is an empty string")
        if sheet < 0:
            raise ValueError("'sheet' must be >= 0")
        if column < 0:
            raise ValueError("'column' must be >= 0")
        if row < 0:
            raise ValueError("'row' must be >= 0")
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
        if 0 == len(name):
            raise ValueError("'name' is an empty string")
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

        # LibreOffice variables.
        self._oSheet = None

        if sheets:
            if isinstance(index_or_name, int):
                if index_or_name < 0:
                    raise ValueError("'index_or_name' must be >= 0")
                # get by index
                self._oSheet = self._sheets.o_sheets.getByIndex(index_or_name)
                self._is_null = False
            else:
                if 0 == len(index_or_name):
                    raise ValueError("'index_or_name' is an empty string")
                # get by name
                self._oSheet = self._sheets.o_sheets.getByName(index_or_name)
                self._is_null = False
        else:
            raise ValueError("'sheets' value is None")

    @property
    def is_null(self):
        """
        Checking if the sheet is null.

        @rtype:   bool
        @return:  Value insertion result
        """
        return self._is_null

    @is_null.setter
    def is_null(self, value):
        """
        Side-effect protection. Raising an exception 'ValueError'.
        """
        raise ValueError(_MSG_EXCEPT_SIDE_EFFECT.format("is_null"))

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
        if col < 0:
            raise ValueError("'col' must be >= 0")
        if row < 0:
            raise ValueError("'row' must be >= 0")
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

    def cell_value_by_index(self, col, row, val_type="AUTO"):
        """
        Get cell value.

        @type  col: int
        @param col: Cell column index

        @type  row: int
        @param row: Cell row index

        @type  val_type: string
        @param val_type: Data type of return value

        @rtype:   long, int, float or string
        @return:  Value. Value type depends on val_type parameter
        """
        if col < 0:
            raise ValueError("'col' must be >= 0")
        if row < 0:
            raise ValueError("'row' must be >= 0")
        value = None
        oCell = self._oSheet.getCellByPosition(col, row)

        if val_type == "AUTO":
            # Return data type is set based on getType()
            value_type = oCell.getType()
            if VALUE == value_type:
                value = oCell.getValue()
            if FORMULA == value_type:
                value = oCell.getFormula()
            if TEXT == value_type:
                value = oCell.getString()
        elif val_type == "VALUE":
            # Force return data type to be cell value
            value = oCell.getValue()
        elif val_type == "FORMULA":
            # Force return data type to be cell formula
            value = oCell.getFormula()
        elif val_type == "STRING":
            # Force return data type to be cell string
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

        # LibreOffice variables.
        self._oSheets = None
        if self._document:
            self._oSheets = self._document.o_doc.getSheets()
        else:
            raise ValueError("'document' value is None")

    @property
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

    @is_null.setter
    def is_null(self, value):
        """
        Side-effect protection. Raising an exception 'ValueError'.
        """
        raise ValueError(_MSG_EXCEPT_SIDE_EFFECT.format("is_null"))

    def sheet(self, index_or_name):
        """
        Get sheet by index or name.

        @type  index_or_name: int, string
        @param index_or_name: Sheet index or name

        @rtype:   Sheet
        @return:  Sheet object
        """
        return Sheet(self, index_or_name)

    @property
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

    @count.setter
    def count(self, value):
        """
        Side-effect protection. Raising an exception 'ValueError'.
        """
        raise ValueError(_MSG_EXCEPT_SIDE_EFFECT.format("count"))

    @property
    def o_sheets(self):
        """
        LibreOffice/OpenOffice Calc Spreadsheets object.

        Required for Fields and Field classes. Not recommended use it directly.

        @rtype:   com::sun::star::sheet::XSpreadsheets
        @return:  Libre/Open office Spreadsheets object
        """
        return self._oSheets

    @o_sheets.setter
    def o_sheets(self, value):
        """
        Side-effect protection. Raising an exception 'ValueError'.
        """
        raise ValueError(_MSG_EXCEPT_SIDE_EFFECT.format("o_sheets"))

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
        if index < 0:
            raise ValueError("'index' must be >= 0")
        if 0 == len(name):
            raise ValueError("'name' is an empty string")
        result = False
        if self.o_sheets:
            self.o_sheets.insertNewByName(name, index)
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
        if 0 == len(name):
            raise ValueError("'name' is an empty string")
        result = False
        if self.o_sheets:
            self.o_sheets.removeByName(name)
            result = True
        return result

###############################################################################
###############################################################################
###############################################################################


class Document:
    def __init__(self,
                 autostart=False,
                 office='soffice \
--accept="socket,host=localhost,port=2002;urp;"',
                 connection_string="\
uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext",
                 timeout=30,
                 attempt_period=0.1):
        """
        Constructor

        @type  autostart: bool
        @param autostart: Auto Starts Libre/Open Office with a listening socket

        @type  office: string
        @param office: Libre/Open Office startup string

        @type  connection_string: string
        @param connection_string: Libre/Open office initialization string

        @type  timeout: int
        @param timeout: Timeout for starting Libre/Open Office in seconds

        @type  attempt_period: int
        @param attempt_period: Timeout between attempts in seconds
        """
        self._sheets = None
        self._fields = None
        self._connection_string = connection_string

        # LibreOffice variables.
        self._oResolver = None
        self._oContext = None
        self._oDesktop = None
        self._oDoc = None
        self._oLocal = uno.getComponentContext()

        if self._oLocal:
            self._oResolver = \
                self._oLocal.ServiceManager.createInstanceWithContext(
                    "com.sun.star.bridge.UnoUrlResolver", self._oLocal)
            if autostart:
                self._autostart_office(office, timeout, attempt_period)
            else:
                self._init_doc()

    def __enter__(self):
        """
        PEP 0343 - The “with” statement
        The specification, background, and examples for the Python with
        statement.

        The with statement will bind this method’s return value to the
        target(s) specified in the as clause of the statement, if any.
        """
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """
        PEP 0343 - The “with” statement
        The specification, background, and examples for the Python with
        statement.

        Exit the runtime context related to this object.
        """
        # Nothing to do
        pass

    def _autostart_office(self, office, timeout, attempt_period):
        """
        Starts Libre/Open Office with a listening socket.

        @type  timeout: int
        @param timeout: Timeout for starting Libre/Open Office in seconds

        @type  attempt_period: int
        @param attempt_period: Timeout between attempts in seconds

        @type  office: string
        @param office: Libre/Open Office startup string
        """
        #######################################################################
        def start_office_instance(_office):
            """
            Starts Libre/Open Office with a listening socket.

            @type  office: string
            @param office: Libre/Open Office startup string
            """
            # Fork to execute Office
            if os.fork():
                return

            # Start OpenOffice.org and report any errors that occur.
            try:
                retcode = subprocess.call(_office, shell=True)
                if retcode < 0:
                    raise OSError(retcode, "Office was terminated by signal")
                elif retcode > 0:
                    raise OSError(retcode, "Office returned")
            except OSError as ose:
                raise OSError(ose)

            # Terminate this process when Office has closed.
            raise SystemExit()

        #######################################################################
        waiting = False
        try:
            self._init_doc()
        except NoConnectException as e:
            waiting = True
            start_office_instance(office)
        except DisposedException as e:
            waiting = True

        if waiting:
            exception = None
            steps = int(timeout/attempt_period)
            for i in range(steps + 1):
                try:
                    self._init_doc()
                    break
                except (NoConnectException, DisposedException) as e:
                    exception = e
                    time.sleep(attempt_period)
            else:
                if exception:
                    raise NoConnectException(exception)
                else:
                    raise NoConnectException()

    def _init_doc(self):
        """
        Initialize Libre/Open Office connection
        """
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

    @property
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

    @is_null.setter
    def is_null(self, value):
        """
        Side-effect protection. Raising an exception 'ValueError'.
        """
        raise ValueError(_MSG_EXCEPT_SIDE_EFFECT.format("is_null"))

    def _to_properties(self, **args):
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

    @property
    def o_doc(self):
        """
        LibreOffice/OpenOffice Calc document object.

        Required for Fields and Field classes. Not recommended use it directly.

        @rtype:   com::sun::star::lang::XComponent
        @return:  Libre/Open office document object
        """
        return self._oDoc

    @o_doc.setter
    def o_doc(self, value):
        """
        Side-effect protection. Raising an exception 'ValueError'.
        """
        raise ValueError(_MSG_EXCEPT_SIDE_EFFECT.format("o_doc"))

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
        else:
            raise ValueError("'doc_name' is an empty string")
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
        if self._oDoc:
            if 0 == len(doc_name):
                self._oDoc.store()
            else:
                full_file_name = unohelper.systemPathToFileUrl(doc_name)
                try:
                    self._oDoc.storeToURL(
                                full_file_name,
                                self._to_properties(FilterName=filter_name))
                    result = True
                except IllegalArgumentException as e:
                    raise IllegalArgumentException(e)
                except ErrorCodeIOException as e:
                    raise ErrorCodeIOException(e)
                except IOException as e:
                    raise IOException(e)
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
            if self._oDoc:
                self._oDoc.close(True)
                self._oDoc = None
                result = True
        except ErrorCodeIOException as e:
            raise ErrorCodeIOException(e)
        except IOException as e:
            raise IOException(e)
        return result

    @property
    def sheets(self):
        """
        Get Sheets document's object.

        @rtype:   Sheets
        @return:  Sheets object
        """

        if self._sheets is None:
            self._sheets = Sheets(self)
        return self._sheets

    @sheets.setter
    def sheets(self, value):
        """
        Side-effect protection. Raising an exception 'ValueError'.
        """
        raise ValueError(_MSG_EXCEPT_SIDE_EFFECT.format("sheets"))

    @property
    def fields(self):
        """
        Get Fields document's object.

        @rtype:   Fields
        @return:  Fields object
        """
        if self._fields is None:
            self._fields = Fields(self)
        return self._fields

    @fields.setter
    def fields(self, value):
        """
        Side-effect protection. Raising an exception 'ValueError'.
        """
        raise ValueError(_MSG_EXCEPT_SIDE_EFFECT.format("fields"))

    @property
    def version(self):
        """
        Get library version.

        @rtype:   string
        @return:  library version
        """
        return __version__

    @version.setter
    def version(self, value):
        """
        Side-effect protection. Raising an exception 'ValueError'.
        """
        raise ValueError(_MSG_EXCEPT_SIDE_EFFECT.format("version"))
