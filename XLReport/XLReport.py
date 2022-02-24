import collections
from XLReport.XLBase import SetColWidth, DrawUnderLine
from XLReport.XLBase import Styles
import XLReport.XLArray as Array
import XLReport.XLImage as Image
import XLReport.XLDatatable as Datatable

ValueStyles = {
    'NameFont': Styles['Font']['Italic'],
    'NameAlign': Styles['Align']['Left'],
    'ValueFont': Styles['Font']['Base'],
    'ValueAlign': Styles['Align']['Right'],
    'Line': Styles['Border']['ThinUnder']
}
ReportStyles = {
    'Font': Styles['Font']['Bold'],
    'Align': Styles['Align']['Left'],
    'Line': Styles['Border']['ThickUnder']
}


def InsertValue(worksheet, baseRow, baseCol, valueCol, name, obj):
    # set value/style for name cell
    nameC = worksheet.cell(row=baseRow, column=baseCol, value=name)
    nameC.font = ValueStyles['NameFont']
    nameC.alignment = ValueStyles['NameAlign']
    # set value for value cell
    valC = worksheet.cell(row=baseRow, column=valueCol, value=obj)
    valC.font = ValueStyles['ValueFont']
    valC.alignment = ValueStyles['ValueAlign']
    SetColWidth(worksheet, valueCol, Styles['RowSize']['ValueCell'])
    # draw underline
    DrawUnderLine(worksheet, baseRow,
                  baseCol, valueCol, ValueStyles['Line'])
    return 1


def IsReport(obj):
    return isinstance(obj, collections.Mapping) \
        and not Array.IsArray(obj) \
        and not Datatable.IsDatatable(obj) \
        and not Image.IsImage(obj)


def CountColReport(report, noname=False):
    colcnt = 0
    tmpcnt = 0
    if noname:
        nameColCnt = 0
    else:
        nameColCnt = 1
    for key in report:
        # if item is image
        if Image.IsImage(report[key]):
            tmpcnt = Image.CountColImage(report[key])
        # if item is datatable
        elif Datatable.IsDatatable(report[key]):
            tmpcnt = 0
        # if item is array
        elif Array.IsArray(report[key]):
            tmpcnt = Array.CountColArray(report[key])
        # if item is report
        elif IsReport(report[key]):
            tmpcnt = CountColReport(report[key])
        # else item is value
        else:
            tmpcnt = 0
        colcnt = max([colcnt, tmpcnt + nameColCnt])
    return colcnt


def CheckArrayExistance(report):
    re = False
    for key in report:
        if Array.IsArray(report[key]):
            return True
        elif IsReport(report[key]):
            re = re | CheckArrayExistance(report[key])
    return re


def IndentSizeReport(report):
    # indentcnt = 0
    # tmpcnt = 0
    # for key in report:
    #     # if item is image
    #     if Image.IsImage(report[key]):
    #         tmpcnt = Image.IndentSizeImage(report[key])
    #     # if item is datatable
    #     elif Datatable.IsDatatable(report[key]):
    #         tmpcnt = 0
    #     # if item is array
    #     elif Array.IsArray(report[key]):
    #         tmpcnt = Array.IndentSizeArray(report[key])
    #     # if item is report
    #     elif IsReport(report[key]):
    #         tmpcnt = IndentSizeReport(report[key])\
    #             + Styles['RowSize']['Indent']
    #     # else item is value
    #     else:
    #         tmpcnt = 0
    #     indentcnt = max([indentcnt, tmpcnt])
    a = CountColReport(report)
    indentcnt = CountColReport(report) * Styles['RowSize']['Indent']
    if CheckArrayExistance(report):
        indentcnt = indentcnt - Styles['RowSize']['Indent']\
            + Array.Styles['Table']['Index']['ColWidth']
    return indentcnt


def InsertReport(worksheet, baseRow, baseCol, valueCol, name, report, noname=False):
    rowCnt = 0
    # if no name row, child column is base column
    childCol = baseCol
    # write name row
    if not noname:
        # write name row
        nameC = worksheet.cell(row=baseRow, column=baseCol, value=name)
        nameC.font = ReportStyles['Font']
        nameC.alignment = ReportStyles['Align']
        DrawUnderLine(worksheet, baseRow,
                      baseCol, valueCol, ReportStyles['Line'])
        # SetColWidth(worksheet, baseCol, Styles['RowSize']['Indent'])
        # add row cnt for name row
        rowCnt += 1

        childCol = baseCol + 1
    # write childs
    for key in report:
        # if item is image
        if Image.IsImage(report[key]):
            childrows = Image.InsertImage(
                worksheet, baseRow+rowCnt, childCol, valueCol, key, report[key])
        # if item is datatable do nothing
        elif Datatable.IsDatatable(report[key]):
            childrows = 0
        # if item is array
        elif Array.IsArray(report[key]):
            childrows = Array.InsertArray(
                worksheet, baseRow+rowCnt, childCol, valueCol, key, report[key])
        # if item is report
        elif IsReport(report[key]):
            childrows = InsertReport(
                worksheet, baseRow+rowCnt, childCol, valueCol, key, report[key])
        else:
            childrows = InsertValue(
                worksheet, baseRow+rowCnt, childCol, valueCol, key, report[key])
        rowCnt += childrows
    return rowCnt
