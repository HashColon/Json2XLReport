# DataTable is for large data table output.
# DataTable is not include in printing area
# DataTable is placed in the right side of report

import collections
import numbers
from typing import Iterable
from XLReport.XLBase import DrawUnderLine, SetColWidth
from XLReport.XLBase import Styles as BaseStyles
import XLReport.XLReport as Report
import XLReport.XLArray as Array

Styles = {
    'NameRow': {
        'Font': BaseStyles['Font']['Title'],
        'Align': BaseStyles['Align']['Left'],
        'Line': BaseStyles['Border']['DoubleUnder']
    },
    'Header': {
        'Font': BaseStyles['Font']['DataBold'],
        'Align': BaseStyles['Align']['Center'],
        'Fill': BaseStyles['Fill']['LightShade'],
        'Line': BaseStyles['Border']['ThinUnder']
    },
    'Index': {
        'Font': BaseStyles['Font']['DataBold'],
        'Align': BaseStyles['Align']['Center'],
        'Fill': BaseStyles['Fill']['LightShade'],
        'Line': BaseStyles['Border']['ThinRight']
    },
    'Data': {
        'Font': BaseStyles['Font']['Data'],
        'Align': BaseStyles['Align']['Center'],
    }
}
Keys = {
    'Index': "__Index",
    'InsertIndex': "__InsertHeader",
    'Header': "__Header",
    'InsertHeader': "__InsertHeader",
    'ColumnFormat': "__ColumnFormat",
    'Data': "__Data"
}
DatatableIDKey = "__FelineReportDatatable"


def IsDatatable(jsonObj):
    return isinstance(jsonObj, Datatable) or _isDatatable_json(jsonObj)


def _isDatatable_json(jsonObj):
    if isinstance(jsonObj, collections.Mapping) \
            and (DatatableIDKey in jsonObj):
        item = jsonObj[DatatableIDKey]
        if Keys['Data'] in item:
            return True
        else:
            return False
    else:
        return False


def TryDatatableConversion(obj):
    if _isDatatable_json(obj):
        item = obj[DatatableIDKey]
        iIndex = None
        if Keys['Index'] in item:
            iIndex = item[Keys['Index']]
        iInsIndex = True
        if Keys['InsertIndex'] in item:
            iInsIndex = item[Keys['InsertIndex']]
        iHeader = None
        if Keys['Header'] in item:
            iHeader = item[Keys['Header']]
        iInsHeader = True
        if Keys['InsertHeader'] in item:
            iInsHeader = item[Keys['InsertHeader']]
        iColFormat = None
        if Keys['ColumnFormat'] in item:
            iColFormat = item[Keys['ColumnFormat']]
        obj = Datatable(
            item[Keys['Data']],
            index=iIndex, insertIndex=iInsIndex,
            header=iHeader, insertHeader=iInsHeader,
            colWidthFormat=iColFormat)
        return True
    elif isinstance(obj, Datatable):
        return True
    else:
        return False


def InsertDataTable(worksheet, baseRow, baseCol, name, obj):
    if TryDatatableConversion(obj):
        return obj.Insert(worksheet, baseRow, baseCol, name)
    else:
        return False


def FindDatatable(report, dts, currentName=''):
    # search in childs
    for key in report:
        # if item is datatable, add to dts
        if IsDatatable(report[key]):
            fullname = currentName + key
            dts.append([fullname, report[key]])
        # if item is report
        elif Report.IsReport(report[key]):
            fullname = currentName + key + '.'
            FindDatatable(report[key], dts, fullname)


class Datatable:
    def __init__(self, data, index=None, insertIndex=True, header=None, insertHeader=True, colWidthFormat=None):
        self.Data = data
        self.Header = header
        self.Index = index
        self.InsertIndex = insertIndex
        self.InsertHeader = insertHeader
        self.WidthFormat = colWidthFormat

    def _getMaxCol(self):
        maxColCnt = -1
        for r in self.Data:
            if hasattr(r, '__len__'):   # check if len is available
                if maxColCnt < len(r):
                    maxColCnt = len(r)
            else:
                if maxColCnt < 1:
                    maxColCnt = 1
        return maxColCnt

    def _getEndColumn(self, baseCol, includeIndex):
        if includeIndex:
            idxCol = 1
        else:
            idxCol = 0
        maxColCnt = self._getMaxCol()
        return baseCol + maxColCnt + idxCol - 1

    def Insert(self, worksheet, baseRow, baseCol, name):
        # get max column numbers of the d4ata
        maxColNum = self._getMaxCol()
        endCol = self._getEndColumn(baseCol, self.InsertIndex)
        # set value/style for name cell
        C_name = worksheet.cell(
            row=baseRow, column=baseCol, value="DATA: "+name)
        C_name.font = Styles['NameRow']['Font']
        C_name.alignment = Styles['NameRow']['Align']
        # draw undeline under name row
        DrawUnderLine(worksheet, row=baseRow,
                      startCol=baseCol, endCol=endCol,
                      lineType=Styles['NameRow']['Line'])
        # insert header
        if self.InsertHeader:
            # if header is none, give column number instead
            if self.Header is None:
                pHeader = list(range(maxColNum))
            else:
                pHeader = self.Header
            # if row index should be printed, add a index col
            if self.InsertIndex:
                pHeader.insert(0, "#")
            # insert header row
            for i, header in enumerate(pHeader):
                r = baseRow + 1
                c = baseCol + i
                C_header = worksheet.cell(row=r, column=c, value=header)
                C_header.font = Styles['Header']['Font']
                C_header.alignment = Styles['Header']['Align']
                C_header.fill = Styles['Header']['Fill']
                C_header.border = Styles['Header']['Line']
        # Set index
        pIndex = []
        if self.InsertIndex:
            if self.Index is None:
                pIndex = range(len(self.Data))
            elif len(self.Index) < len(self.Data):
                pIndex = pIndex + self.Index
                pIndex = pIndex + range(len(self.Index), len(self.Data))
            else:
                pIndex = self.Index
        # set Column width format
        widthFormat = []
        if self.InsertIndex:
            widthFormat.append(max([len(str(idx)) for idx in pIndex])+1.5)
        if Array.IsArray(self.WidthFormat):
            widthFormat = widthFormat + self.WidthFormat
        elif isinstance(self.WidthFormat, numbers.Number):
            widthFormat = widthFormat + [self.WidthFormat] * maxColNum
        for i, wFormat in enumerate(widthFormat):
            SetColWidth(worksheet, baseCol+i, widthFormat[i])
        # Set base row/col for data
        if self.InsertIndex:
            dataCol = baseCol+1
            maxColNum += 1
        else:
            dataCol = baseCol
        if self.InsertHeader:
            dataRow = baseRow+2
        else:
            dataRow = baseRow+1
        # insert index/data row by row
        for rIdx, rowData in enumerate(self.Data):
            currRow = dataRow + rIdx
            # insert index
            if self.InsertIndex:
                C_Index = worksheet.cell(row=currRow, column=baseCol,
                                         value=pIndex[rIdx])
                C_Index.font = Styles['Index']['Font']
                C_Index.alignment = Styles['Index']['Align']
                C_Index.fill = Styles['Index']['Fill']
                C_Index.border = Styles['Index']['Line']
            # insert data
            if isinstance(rowData, Iterable):
                for cIdx, itemData in enumerate(rowData):
                    currCol = dataCol + cIdx
                    C_data = worksheet.cell(
                        row=currRow, column=currCol, value=itemData)
                    C_data.font = Styles['Data']['Font']
                    C_data.alignment = Styles['Data']['Align']
            # insert data(if row is single data)
            else:
                currCol = dataCol
                C_data = worksheet.cell(
                    row=currRow, column=currCol, value=rowData)
                C_data.font = Styles['Data']['Font']
                C_data.alignment = Styles['Data']['Align']

        return maxColNum
