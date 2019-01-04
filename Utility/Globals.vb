Imports Microsoft.Office.Interop

Public Module Globals

    'Paths for both excel files
    Public strExcelPath_1 As String
    Public strExcelPath_2 As String

    'Instances for both excel files
    Public objExcel_1 As Excel.Application
    Public objExcel_2 As Excel.Application

    Public objWorkbook_1 As Excel.Workbook
    Public objWorkbook_2 As Excel.Workbook

    Public objWorksheet_1 As Excel.Worksheet
    Public objWorksheet_2 As Excel.Worksheet

    Public intProcID_1 As Integer
    Public intProcID_2 As Integer

    'Location
    'Key: Page number
    'Value: Location of page in the whole excel sheet
    Public objLocation_1 As Dictionary(Of Integer, Location)
    Public objLocation_2 As Dictionary(Of Integer, Location)

    Public objRemoveCol As Dictionary(Of Integer, List(Of Integer))
    Public objAddCol As Dictionary(Of Integer, List(Of Integer))
    Public objRemoveRow As Dictionary(Of Integer, List(Of Integer))
    Public objAddRow As Dictionary(Of Integer, List(Of Integer))

    'Key: Page number
    'Value: Data change [From - To]
    Public objChangeData As Dictionary(Of Integer, Dictionary(Of String, String))

    'Key: Page number
    'Value: Equivalent columns in each page of both excel sheets
    Public objEquivalentColumns As Dictionary(Of Integer, Dictionary(Of Integer, Integer))
    'Key: Page number
    'Value: Equivalent rows in each page of both excel sheets
    Public objEquivalentRows As Dictionary(Of Integer, Dictionary(Of Integer, Integer))

    Public objValueResult_1 As Dictionary(Of Integer, Dictionary(Of Tuple(Of Integer, Integer), List(Of ValueError)))
    Public objValueResult_2 As Dictionary(Of Integer, Dictionary(Of Tuple(Of Integer, Integer), List(Of ValueError)))
    Public objFormatResult As Dictionary(Of Integer, Dictionary(Of Tuple(Of Integer, Integer), List(Of FormatError)))

    'True: Customization task; False: Conversion task
    Public blnCustFlg As Boolean

    Public intNoOfPages As Integer
    Public dblThreshold As Double

    'Flag for compare by value
    'True: Compare each column/row to column/row most similar to it
    'False: Compare each column/row to immediate similar column/row
    Public blnBestMatchFlg As Boolean

    'Flag if to compare by format
    Public blnCompareMerge As Boolean
    Public blnCompareTextWrap As Boolean
    Public blnCompareTextAlign As Boolean
    Public blnCompareOrientation As Boolean
    Public blnCompareBorder As Boolean
    Public blnCompareBackColor As Boolean
    Public blnCompareFont As Boolean

    Public blnCompareDone As Boolean
    Public blnCancelFlg As Boolean

    'Files for export
    Public strExportExcel_1 As String
    Public strExportExcel_2 As String
    Public strExportReport As String

    'Generated report
    Public objExcel_3 As Excel.Application
    Public objWorkbook_3 As Excel.Workbook
    Public objWorksheet_3 As Excel.Worksheet

    'Key: Page number
    'Value: Differences found
    Public objDifferences As Dictionary(Of Integer, Dictionary(Of Tuple(Of String, String), Tuple(Of String, String)))
End Module
