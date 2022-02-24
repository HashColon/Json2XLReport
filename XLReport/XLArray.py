import collections
import math
from XLReport.XLBase import DrawUnderLine, SetColWidth, SetRowHeight
from XLReport.XLBase import Styles as BaseStyles

Styles = {
    "NameRow": {
        'Font': BaseStyles['Font']['Italic'],
        'Align': BaseStyles['Align']['Left'],
        'Line': BaseStyles['Border']['ThinUnder'],
    },
    "Table": {
        "Index": {
            'Font': BaseStyles['Font']['DataBold'],
            'Align': BaseStyles['Align']['Center'],
            'Fill': BaseStyles['Fill']['LightShade'],
            'Line': BaseStyles['Border']['ThinAllround'],
            'ColWidth': BaseStyles['RowSize']['ArrayIndex']
        },
        "Data": {
            'Font': BaseStyles['Font']['Data'],
            'Align': BaseStyles['Align']['LeftWithWrap'],
            'Line': BaseStyles['Border']['ThinAllround'],
            'MarginWidth': BaseStyles['RowSize']['Indent']

        },
        'GuideLine': BaseStyles['Border']['DoubleLeft'],
    },
}
ColOccupied = 3


def IsArray(obj):
    return isinstance(obj, collections.Sequence) and \
        not isinstance(obj, str)


def CountColArray(obj):
    return ColOccupied


def IndentSizeArray(obj):
    return BaseStyles['RowSize']['Indent'] \
        + Styles['Table']['Index']['ColWidth']


def InsertArray(worksheet, baseRow, baseCol, valueCol, name, data):
    # set value/style for name cell
    nameC = worksheet.cell(row=baseRow, column=baseCol, value=name)
    nameC.font = Styles['NameRow']['Font']
    nameC.alignment = Styles['NameRow']['Align']
    DrawUnderLine(worksheet, baseRow, baseCol,
                  valueCol, Styles['NameRow']['Line'])
    # Set base cell for table data
    arrBaseRow = baseRow+1
    arrIdxCol = valueCol-2
    arrValCol = valueCol-1
    # set indent size
    # SetColWidth(worksheet, baseCol,
    #             BaseStyles['RowSize']['Indent'])
    # SetColWidth(worksheet, baseCol+1,
    #             BaseStyles['RowSize']['Indent'])
    # SetColWidth(worksheet, arrIdxCol,
    #             Styles['Table']['Index']['ColWidth'])

    # set array data row by row
    for i in range(0, len(data)):
        arrRow = arrBaseRow+i
        # draw guide line
        worksheet.cell(row=arrRow, column=baseCol+1).border = \
            Styles['Table']['GuideLine']
        # merges data cells
        worksheet.merge_cells(
            start_row=arrRow, start_column=arrValCol,
            end_row=arrRow, end_column=valueCol)
        # write index cell
        C_idx = worksheet.cell(row=arrRow, column=arrIdxCol, value=i)
        C_idx.font = Styles['Table']['Index']['Font']
        C_idx.alignment = Styles['Table']['Index']['Align']
        C_idx.fill = Styles['Table']['Index']['Fill']
        C_idx.border = Styles['Table']['Index']['Line']
        # write value cell
        C_val = worksheet.cell(
            row=arrRow, column=arrValCol, value=str(data[i]))
        C_val.font = Styles['Table']['Data']['Font']
        C_val.alignment = Styles['Table']['Data']['Align']
        C_val.border = Styles['Table']['Index']['Line']
        C_val2 = worksheet.cell(row=arrRow, column=valueCol)
        C_val2.border = Styles['Table']['Index']['Line']
        # set row height
        SetRowHeight(worksheet, arrRow,
                     _computeRowHeight(C_val.value, baseCol, C_val.font.size))
    # return occupying columns
    return 2 + len(data)


def _computeRowHeight(cellValue, col, fontHeight):
    valLen = len(cellValue)
    cellWidth = \
        BaseStyles['RowSize']['PageWidth'] - \
        (col-2) * BaseStyles['RowSize']['Indent'] - \
        Styles['Table']['Index']['ColWidth'] - \
        Styles['Table']['Data']['MarginWidth']
    return (math.ceil(valLen / cellWidth) + 1) * fontHeight
