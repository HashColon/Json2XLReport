from XLReport.XLBase import DrawUnderLine, SetColWidth, SetPrintArea, SetHeaderFooter
from XLReport.XLBase import Styles
import XLReport.XLReport as Report
import XLReport.XLDatatable as Datatable


TitleStyles = {
    "Font": Styles['Font']['Title'],
    "Align": Styles['Align']['Left'],
    "Line": Styles['Border']['DoubleUnder']
}

SheetStyles = {
    "PageSetup": Styles['PageSetup'],
    "PageMargins": Styles['PageMargins'],
    "PrintOpt": Styles['PrintOpt'],
    "View": Styles['View'],
    'DatatableMargin': Styles['DatatableMargin'],
}


def CountColsInSheet(fullreport):
    colcnt = Report.CountColReport(fullreport, noname=True)
    # if no item found (except datatable), return 0
    if colcnt == 0:
        return 0
    else:
        return colcnt + 2


def BuildSheet(worksheet, report, title=None, headerTitle=None, noPrint=False):
    # count columns in a sheet
    valueCol = CountColsInSheet(report)
    # current row/col
    currRow = 1
    currCol = 1
    # write title & set header/footer
    if title is not None:
        titleStr = title
        worksheet['A1'] = title
        worksheet['A1'].font = TitleStyles["Font"]
        worksheet['A1'].alignment = TitleStyles["Align"]
        DrawUnderLine(worksheet, 1, 1, valueCol, TitleStyles['Line'])
        currRow += 1
    # adjust buffer column size
    totIndent = 0
    for i in range(1, valueCol-1):
        SetColWidth(worksheet, i, Styles["RowSize"]["Indent"])
        totIndent += Styles["RowSize"]["Indent"]
    if Report.CheckArrayExistance(report):
        SetColWidth(worksheet, valueCol-2,
                    Styles["RowSize"]["ArrayIndex"])
        totIndent = totIndent - Styles["RowSize"]["Indent"]\
            + Styles["RowSize"]["ArrayIndex"]
    if valueCol > 1:
        bufferColSize = Styles['RowSize']['PageWidth'] \
            - Styles['RowSize']['ValueCell'] - totIndent
        SetColWidth(worksheet, valueCol-1, bufferColSize)
        SetColWidth(worksheet, valueCol, Styles["RowSize"]["ValueCell"])
    # Insert report
    reportrows = Report.InsertReport(
        worksheet, currRow, currCol, valueCol,
        name=None, report=report, noname=True)
    currRow += reportrows
    # set sheet layout & print options
    worksheet.page_setup = SheetStyles['PageSetup']
    worksheet.page_margins = SheetStyles['PageMargins']
    worksheet.print_options = SheetStyles['PrintOpt']
    worksheet.sheet_view.view = SheetStyles['View']
    # set header/footer
    SetHeaderFooter(worksheet, headerTitle)
    # if noPrint, hide sheet to prevent printing
    if noPrint:
        worksheet.sheet_state = 'hidden'
    else:
        # set print area
        if (currRow >= 1) or (valueCol >= 1):
            SetPrintArea(worksheet, 1, 1, currRow-1, valueCol)
    # Prepare to insert Datatable
    dts = []
    currRow = 1
    currCol += valueCol
    # find all data tables in current sheet
    Datatable.FindDatatable(report, dts)
    # insert datatables
    for tableInfo in dts:
        datatablename = tableInfo[0]
        datatable = tableInfo[1]
        dtCols = Datatable.InsertDataTable(
            worksheet, currRow, currCol, datatablename, datatable)
        currCol += dtCols
        SetColWidth(worksheet, currCol, SheetStyles['DatatableMargin'])
        currCol += 1
