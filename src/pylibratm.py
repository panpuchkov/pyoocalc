#!/usr/bin/python 
# -*- coding: utf-8 -*- 

"""
@author: Yurii Puchkov
@organization: http://arilot.com/
@license: GPL v3.0
@contact: panpuchkov@gmail.com
"""

import uno

################################################################################
# example
#/usr/bin/libreoffice -accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager"

################################################################################
################################################################################
################################################################################
class ArilotField:
	"""
	Document field.

	Operations on the existing field.
	"""
	m_bIsNull = True
	m_oFields = None
	m_oField = None

	m_oSheet = None
	m_oNamedRanges = None
	m_oRange = None
	m_oCell = None
	m_oCellAddress = None

	def __init__(self, oFields, strName):
		self.m_oFields = oFields
		
		if self.m_oFields:
			if self.m_oFields.m_oNamedRanges.hasByName(strName):
				self.m_oRange = self.m_oFields.m_oNamedRanges.getByName(strName)
				self.m_oCellAddress = self.m_oRange.getReferencePosition()
				oSheets = self.m_oFields.m_oTemplate.m_oDocument.getSheets()
				self.m_oSheet = oSheets.getByIndex(self.m_oCellAddress.Sheet)
				self.m_bIsNull = False
			else:
				self.m_oRange = None
				self.m_bIsNull = True

	def __del__(self):
		self.release()

	def isNull(self):
		"""
		Checking if a field is null.
		"""
		return self.m_bIsNull

	def setValue(self, strValue, nColumn = 0, nRow = 0):
		"""
		Set filed value at position Column/Row
		
		@type  strValue: string
		@param strValue: Cell value
		
		@type  nColumn: integer
		@param nColumn: column index
		
		@type  nRow: integer
		@param nRow: row index
		
		@rtype:   boolean
		@return:  Value insertion result
		"""
		bResult = True
		if self.m_oRange:
			self.m_oCell = self.m_oSheet.getCellByPosition(self.m_oCellAddress.Column + nColumn, self.m_oCellAddress.Row + nRow)
			if self.m_oCell:
				self.m_oCell.setString(strValue)
		else:
			bResult = False
		return bResult

	def value(self, nColumn = 0, nRow = 0):
		"""
		Get filed value at position Column/Row
		
		@type  nColumn: integer
		@param nColumn: column index
		
		@type  nRow: integer
		@param nRow: row index
		
		@rtype:   string
		@return:  Document cell value in string format. Regardless of document cell type.
		"""
		strValue = ""
		if self.m_oRange:
			self.m_oCell = self.m_oSheet.getCellByPosition(self.m_oCellAddress.Column + nColumn, self.m_oCellAddress.Row + nRow)
			if self.m_oCell:
				strValue = self.m_oCell.getString()
		return strValue

	def insertRow(self, nRow, nStep = 1, nNumColumns = 1, nOffset = 0):
		"""
		Insert rows
		
		Insert new rows at the specified position relatively to cell. 
		After the new row insertion the content of the current rows is copied 
		to the new rows.

		@type  nRow: integer
		@param nRow: row index
		
		@type  nStep: integer
		@param nStep: Default value = 1.	// FIXME 
		
		@type  nNumColumns: integer
		@param nNumColumns: Number of columns to copy. Default value = 1 
		
		@type  nOffset: integer
		@param nOffset: Rows offset. Relatively to current field. Default value = 0
		
		@rtype:   boolean
		@return:  Operation result
		"""
		bResult = True
		if self.m_oCell and self.m_oFields:
			#oCellAddress = self.m_oCell.CellAddress
			#oSheets = self.m_oFields.m_oTemplate.m_oDocument.getSheets()
			#oSheet = oSheets.getByIndex(oCellAddress.Sheet)
			if self.m_oSheet:
				self.m_oSheet.Rows.insertByIndex(self.m_oCellAddress.Row + nOffset, nNumColumns) 
		else:
			bResult = False
		return bResult

	def insertColumn(self, nColumn, nStep = 1, nNumRows = 1, nOffset = 0):
		bResult = True
		if self.m_oCell and self.m_oFields:
			#oCellAddress = self.m_oCell.CellAddress
			#oSheets = self.m_oFields.m_oTemplate.m_oDocument.getSheets()
			#oSheet = oSheets.getByIndex(oCellAddress.Sheet)
			if self.m_oSheet:
				self.m_oSheet.Columns.insertByIndex(self.m_oCellAddress.Column + nOffset, nNumRows) 
		else:
			bResult = False
		return bResult

	def remove(self):
		# FIXME, Not implemented yet.
		return None

	def release(self):
		# FIXME, Not implemented yet.
		return None

################################################################################
################################################################################
################################################################################
class ArilotFields:
	"""
	Document fields.
	Search and manage fields.
	"""
	m_oField = None
	m_oTemplate = None

	m_oSheets = None
	m_oNamedRanges = None

	def __init__(self, oTemplate):
		self.m_oTemplate = None
		self.attachTemplate(oTemplate)
		if self.m_oTemplate:
			self.m_oNamedRanges = self.m_oTemplate.m_oDocument.NamedRanges

	def __del__(self):
		self.release()

	def field(self, strName):
		"""
		Get document field by name
		
		@type  strName: string
		@param strName: Field name
		
		@rtype:   Field object
		@return:  Field object

		"""
		self.m_oField = ArilotField(self, strName)
		if None == self.m_oField.isNull():
			self.m_oField = None
		return self.m_oField

	def insertSpreadsheet(self, strName, nIndex):
		bResult = False
		# FIXME, Not implemented yet.
		return bResult

	def add(self, strName):
		bResult = False
		# FIXME, Not implemented yet.
		return bResult

	def attachTemplate(self, oTemplate):
		self.m_oTemplate = oTemplate

	def count(self):
		nCount = -1;
		# FIXME, Not implemented yet.
		return nCount

	def release(self):
		# FIXME, Not implemented yet.
		return None

################################################################################
################################################################################
################################################################################
class ArilotTemplate:
	m_oFields = None

	m_oLocal = None
	m_oResolver = None
	m_oContext = None
	m_oDesktop = None
	m_oDocument = None

	def __init__(self, strConnectionString = "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"):
		self.m_oLocal = uno.getComponentContext()
		if self.m_oLocal:
			self.m_oResolver = self.m_oLocal.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", self.m_oLocal)
			if self.m_oResolver:
				self.m_oContext = self.m_oResolver.resolve(strConnectionString)
				if self.m_oContext:
					self.m_oDesktop = self.m_oContext.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.m_oContext)

	def __del__(self):
		self.release()

	def fields(self):
		self.m_oFields = ArilotFields(self)
		return self.m_oFields

	def version(self):
		return "0.0.1"

	def closeDocument(self):
		bResult = False
		# FIXME, Not implemented yet.
		return bResult

	def saveDocument(self, strDocname = ""):
		bResult = True
		if self.m_oDocument:
			if 0 == len(strDocname):
				self.m_oDocument.store()
			else:
				strFullFileName = "file://" + strDocname
				self.m_oDocument.storeToURL(strFullFileName)
		else:
			bResult = False
		return bResult

	def openDocument(self, strDocname):
		bResult = True
		if len(strDocname) > 0 and self.m_oDesktop:
			strFullFileName = "file://" + strDocname
			self.m_oDocument = self.m_oDesktop.loadComponentFromURL(strFullFileName ,"_blank", 0, ())
		else:
			bResult = False
		return bResult

	def newDocument(self, strDocname):
		bResult = False
		# FIXME, Not implemented yet.
		return bResult

	def release(self):
		# FIXME, Not implemented yet.
		return None

################################################################################
################################################################################
################################################################################
class ArilotTemplateManager:
	m_oTemplate = None

	def __init__(self):
		self.m_oTemplate = ArilotTemplate()

	def Template(self):
		return self.m_oTemplate

	def templateManagerName(self):
		return "pylibratm"

