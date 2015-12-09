import uno
from com.sun.star.sheet.CellInsertMode import DOWN
from com.sun.star.sheet.FillDirection import TO_BOTTOM
from com.sun.star.sheet.FillMode import SIMPLE
from com.sun.star.sheet.CellDeleteMode import UP
from com.sun.star.sheet.CellFlags import FORMULA

# Start with a few generic spreadsheet functions deriving a range from another range
def getRangeByAddress(obj, a):
    '''Retrieve a range by range address within doc, sheet or parent range.'''
    if obj.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
        # use the sheet specified by given address
        oSheet = obj.Sheets.getByIndex(a.Sheet)
    else:
        # use address relatively to given parent object (range or sheet)
        oSheet = obj
    try:
        return oSheet.getCellRangeByPosition(
            a.StartColumn,
            a.StartRow,
            a.EndColumn,
            a.EndRow
            )
    except:
        return None

def getOffsetRange(oRg, nRowOffset=0, nColOffset=0, nRowResize=0, nColResize=0):
    '''Analogue to the spreadsheet's OFFSET function'''
    addr = oRg.getRangeAddress()
    addr.StartRow = addr.StartRow + nRowOffset
    addr.EndRow = addr.EndRow + nRowOffset
    addr.StartColumn = addr.StartColumn + nColOffset
    addr.EndColumn = addr.EndColumn + nColOffset
    if nRowResize > 0 : addr.EndRow = addr.StartRow + nRowResize -1
    if nColResize > 0 : addr.EndColumn = addr.StartColumn + nColResize -1
    return getRangeByAddress(oRg.getSpreadsheet(), addr)


def getCurrentRegion(oRange):
    """Get current region around given range."""
    oCursor = oRange.getSpreadsheet().createCursorByRange(oRange)
    oCursor.collapseToCurrentRegion()
    return oCursor

def getCurrentColumnsAddress(oRange):
    """Get address of intersection between range and current region's columns"""
    oCurrent = getCurrentRegion(oRange)
    oAddr = oRange.getRangeAddress()
    oCurrAddr = oCurrent.getRangeAddress()
    oAddr.StartColumn = oCurrAddr.StartColumn
    oAddr.EndColumn = oCurrAddr.EndColumn
    return oAddr

def getMergedArea(oRange):
    oCursor = oRange.getSpreadsheet().createCursorByRange(oRange)
    oCursor.collapseToMergedArea()
    return oCursor

class MergedRanges:
    def __init__(self, ):
        pass

class InsertRange:
    def __init__(self, ctx):
        smgr = ctx.ServiceManager
        dtp = smgr.createInstanceWithContext("com.sun.star.frame.Desktop",ctx)
        doc = dtp.getCurrentComponent()
        oSel = doc.getCurrentSelection()
        # Exception if oSel is not a single range on a sheet
        self.Sheet = oSel.getSpreadsheet()
        self.Area = getCurrentColumnsAddress(oSel)
        self.View = doc.getCurrentController()

        self.GlobalSettings = smgr.createInstance('com.sun.star.sheet.GlobalSheetSettings')
        self.Expand = self.GlobalSettings.ExpandReferences
        
        oRg = getRangeByAddress(self.Sheet, self.Area)

        self.TopRow = getOffsetRange(oRg, nRowOffset= -1) #may be None

        # Having the range of interest, collect the merged ranges:
        self.MergedAreas = doc.createInstance('com.sun.star.sheet.SheetCellRanges')
        # merged ranges have their own c.f.r.
        eCFR = oRg.CellFormatRanges.createEnumeration()
        while eCFR.hasMoreElements():
            n = eCFR.nextElement()
            r = getMergedArea(n)
            c = r.getCellByPosition(0,0)
            if c.getIsMerged():
                self.MergedAreas.addRangeAddress(r.getRangeAddress(),False)

        # Initialize additional references to merged cells as simple list of top-left cells. 
        # The object references move with the insertions.
        # Cells can be moved from the stored references to the new top-left cells then.
        self.MergedCells = {}
        for i in range(self.MergedAreas.getCount()):
            c = self.MergedAreas.getByIndex(i).getCellByPosition(0,0)
            qx = c.queryIntersection(self.Area)
            if qx.getCount() > 0 : self.MergedCells[i] = c
        


    def moveMergedUp(self,):
        '''Move the displayed top-left cell of a merged range to the new top-left cell.
        To be called after insertion'''
        while self.MergedCells:
            i,c = self.MergedCells.popitem()
            src = c.getRangeAddress()
            oM = self.MergedAreas.getByIndex(i)
            tgt = oM.getCellByPosition(0,0).getCellAddress()
            self.Sheet.moveRange(tgt, src)

    def moveMergedDown(self,):
        '''Move the displayed top-left cell of a merged range to the new top-left cell.
        To be called before deletion'''
        while self.MergedCells:
            i,c = self.MergedCells.popitem()
            oM = self.MergedAreas.getByIndex(i)
            addr = oM.getRangeAddress()
            if self.Area.EndRow <= addr.EndRow:
                src = c.getRangeAddress()
                tgt = uno.createUnoStruct('com.sun.star.table.CellAddress')
                tgt.Sheet = src.Sheet
                tgt.Column = src.StartColumn
                tgt.Row = self.Area.EndRow + 1
                self.Sheet.moveRange(tgt, src)


    def unmerge(self,):
        self.merge(False)

    def remerge(self,):
        self.merge(True)

    def merge(self, bM):
        for i in range(self.MergedAreas.getCount()):
            oNext = self.MergedAreas.getByIndex(i)
            oNext.merge(bM)

    def dragDownFormulas(self,):
        '''drag down formula ranges and additional merges from row above oAddr'''
        if self.TopRow:
            oFmlRanges = self.TopRow.queryContentCells(FORMULA)
            e = oFmlRanges.createEnumeration()
            while e.hasMoreElements():
                n = e.nextElement()
                oDrag = getOffsetRange(
                    n, nRowResize = self.Area.EndRow-self.Area.StartRow +2
                    )
                oDrag.fillSeries(TO_BOTTOM, SIMPLE,0,0,0)

    def removeRange(self,):
        self.Sheet.removeRange(self.Area, UP)

    def insertListRows(self,):
        '''Let references expand automagically on insertion directly below the list'''
        self.GlobalSettings.ExpandReferences = True
        self.Sheet.insertCells(self.Area, DOWN)
        self.GlobalSettings.ExpandReferences = self.Expand

    def stripMerged(self, oRanges,):
        e = oRanges.createEnumeration()
        lst = []
        while e.hasMoreElements():
            n = e.nextElement()
            qx = self.MergedAreas.queryIntersection(n.getRangeAddress())
            if qx.getCount() : lst.append(qx.getRangeAddresses())
        if lst:
            oRanges.removeRangeAddresses(tuple(lst))
        return oRanges

    def selectBlanks(self,):
        oCurrent = getRangeByAddress(self.Sheet, self.Area)
        # quirk: queryEmptyCells() returns an empty collection for an entirely blank range
        oBlanks = oCurrent.queryEmptyCells()
        if oBlanks.getCount() == 0: #if all or none are blank
            oBlanks = oCurrent.queryIntersection(self.Area)
        
        #oSel = self.stripMerged(oBlanks)
        self.View.select(oBlanks)

    def selectRows(self,):
        oSel = getRangeByAddress(self.Sheet, self.Area)
        self.View.select(oSel)


def insertListRows(*args):
    '''Unmerge ranges intersection with selected rows within current region,
    insert new row(s) into current region, drag down formula cells, remerge,
    select blank cells of new row(s)'''
    try:
        oLCR = InsertRange(uno.getComponentContext())
    except:
        # fail silently
        return
    oLCR.unmerge()
    oLCR.insertListRows()
    oLCR.dragDownFormulas()
    #oLCR.dragDownMergedAreas()
    oLCR.moveMergedUp()
    oLCR.remerge()
    oLCR.selectBlanks()

    
def removeListRows(*args):
    '''Unmerge ranges intersectinv with selected rows within current region,
    remove selected rows within current region, remerge, select address of removed cells'''
    try:
        oLCR = InsertRange(uno.getComponentContext())
    except:
        # fail silently
        return
    
    oLCR.unmerge()
    oLCR.moveMergedDown()
    oLCR.removeRange()
    oLCR.remerge()
    oLCR.selectRows()

g_exportedScripts = insertListRows, removeListRows
