Imports System.ComponentModel
Imports System.Drawing
Imports System.Threading
Imports Microsoft.Office.Interop.Excel
Imports Utility

Public Class Compare
    'Data for both excel files
    'Key: Page number
    'Value: List of cells[row, column] with their properties
    Private objExcelData_1 As Dictionary(Of Integer, Dictionary(Of Tuple(Of Integer, Integer), Range))
    Private objExcelData_2 As Dictionary(Of Integer, Dictionary(Of Tuple(Of Integer, Integer), Range))

    'Key: Page number
    'Value: Location of each page in excel sheet
    Private objLocation_1 As Dictionary(Of Integer, Location)
    Private objLocation_2 As Dictionary(Of Integer, Location)

    'Object for comparison
    Private objCompareByData As Compare_Data
    Private objCompareByFormat As Compare_Format

    Private dblThreshold As Double

    Private blnBestMatchFlg As Boolean

    Private blnCompareMerge As Boolean
    Private blnCompareTextWrap As Boolean
    Private blnCompareTextAlign As Boolean
    Private blnCompareOrientation As Boolean
    Private blnCompareBorder As Boolean
    Private blnCompareBackColor As Boolean
    Private blnCompareFont As Boolean

    Private objRemoveCol As Dictionary(Of Integer, List(Of Integer))
    Private objAddCol As Dictionary(Of Integer, List(Of Integer))
    Private objRemoveRow As Dictionary(Of Integer, List(Of Integer))
    Private objAddRow As Dictionary(Of Integer, List(Of Integer))

    Private objChangeData As Dictionary(Of Integer, Dictionary(Of String, String))

    Private intNoOfPages As Integer
    'Key: Page number
    'Value: Equivalent columns in each page of both excel sheets
    Private objEquivalentColumns As Dictionary(Of Integer, Dictionary(Of Integer, Integer))
    'Key: Page number
    'Value: Equivalent rows in each page of both excel sheets
    Private objEquivalentRows As Dictionary(Of Integer, Dictionary(Of Integer, Integer))
    'Key: Page number
    'Value: Comparison result in each page of both excel sheets
    Private objValueResult_1 As Dictionary(Of Integer, Dictionary(Of Tuple(Of Integer, Integer), List(Of ValueError)))
    Private objValueResult_2 As Dictionary(Of Integer, Dictionary(Of Tuple(Of Integer, Integer), List(Of ValueError)))
    Private objFormatResult As Dictionary(Of Integer, Dictionary(Of Tuple(Of Integer, Integer), List(Of FormatError)))

    ''' <summary>
    ''' Set threshold in determining similarity of data
    ''' </summary>
    Public WriteOnly Property SetThreshold() As Double
        Set(value As Double)
            dblThreshold = value
        End Set
    End Property

    ''' <summary>
    ''' Set flag if to compare each column/row to column/row most similar to it 
    ''' True: Yes; False: No
    ''' </summary>
    Public WriteOnly Property CompareToBestMatchData() As Boolean
        Set(value As Boolean)
            blnBestMatchFlg = value
        End Set
    End Property

    ''' <summary>
    ''' Set flag if to compare excel sheets on merging
    ''' True: Yes; False: No
    ''' </summary>
    Public WriteOnly Property CompareMerge() As Boolean
        Set(value As Boolean)
            blnCompareMerge = value
        End Set
    End Property

    ''' <summary>
    ''' Set flag if to compare excel sheets on text wrap
    ''' True: Yes; False: No
    ''' </summary>
    Public WriteOnly Property CompareTextWrap() As Boolean
        Set(value As Boolean)
            blnCompareTextWrap = value
        End Set
    End Property

    ''' <summary>
    ''' Set flag if to compare excel sheets on text alignment
    ''' True: Yes; False: No
    ''' </summary>
    Public WriteOnly Property CompareTextAlign() As Boolean
        Set(value As Boolean)
            blnCompareTextAlign = value
        End Set
    End Property

    ''' <summary>
    ''' Set flag if to compare excel sheets on orientation
    ''' True: Yes; False: No
    ''' </summary>
    Public WriteOnly Property CompareOrientation() As Boolean
        Set(value As Boolean)
            blnCompareOrientation = value
        End Set
    End Property

    ''' <summary>
    ''' Set flag if to compare excel sheets on border
    ''' True: Yes; False: No
    ''' </summary>
    Public WriteOnly Property CompareBorder() As Boolean
        Set(value As Boolean)
            blnCompareBorder = value
        End Set
    End Property

    ''' <summary>
    ''' Set flag if to compare excel sheets on back color
    ''' True: Yes; False: No
    ''' </summary>
    Public WriteOnly Property CompareBackColor() As Boolean
        Set(value As Boolean)
            blnCompareBackColor = value
        End Set
    End Property

    ''' <summary>
    ''' Set flag if to compare excel sheets on font
    ''' True: Yes; False: No
    ''' </summary>
    Public WriteOnly Property CompareFont() As Boolean
        Set(value As Boolean)
            blnCompareFont = value
        End Set
    End Property

    ''' <summary>
    ''' Sets number of pages compared
    ''' </summary>
    Public WriteOnly Property NoOfPages() As Integer
        Set(value As Integer)
            intNoOfPages = value
        End Set
    End Property

    ''' <summary>
    ''' Sets the location of compared pages in excel sheet 1
    ''' </summary>
    Public WriteOnly Property Page_Location_1() As Dictionary(Of Integer, Location)
        Set(value As Dictionary(Of Integer, Location))
            objLocation_1 = value
        End Set
    End Property

    ''' <summary>
    ''' Sets the location of compared pages in excel sheet 2
    ''' </summary>
    Public WriteOnly Property Page_Location_2() As Dictionary(Of Integer, Location)
        Set(value As Dictionary(Of Integer, Location))
            objLocation_2 = value
        End Set
    End Property

    ''' <summary>
    ''' Set removed column from excel sheet 1
    ''' </summary>
    Public WriteOnly Property RemovedColumn() As Dictionary(Of Integer, List(Of Integer))
        Set(value As Dictionary(Of Integer, List(Of Integer)))
            objRemoveCol = value
        End Set
    End Property

    ''' <summary>
    ''' Set removed row from excel sheet 1
    ''' </summary>
    Public WriteOnly Property RemovedRow() As Dictionary(Of Integer, List(Of Integer))
        Set(value As Dictionary(Of Integer, List(Of Integer)))
            objRemoveRow = value
        End Set
    End Property

    ''' <summary>
    ''' Set added column from excel sheet 2
    ''' </summary>
    Public WriteOnly Property AddedColumn() As Dictionary(Of Integer, List(Of Integer))
        Set(value As Dictionary(Of Integer, List(Of Integer)))
            objAddCol = value
        End Set
    End Property

    ''' <summary>
    ''' Set added row from excel sheet 2
    ''' </summary>
    Public WriteOnly Property AddedRow() As Dictionary(Of Integer, List(Of Integer))
        Set(value As Dictionary(Of Integer, List(Of Integer)))
            objAddRow = value
        End Set
    End Property

    ''' <summary>
    ''' Set changes in data after customization
    ''' </summary>
    Public WriteOnly Property DataChange() As Dictionary(Of Integer, Dictionary(Of String, String))
        Set(value As Dictionary(Of Integer, Dictionary(Of String, String)))
            objChangeData = value
        End Set
    End Property

    ''' <summary>
    ''' Sets and returns equivalent columns of all pages from both excel sheets
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property EquivalentColumns() As Dictionary(Of Integer, Dictionary(Of Integer, Integer))
        Get
            Return objEquivalentColumns
        End Get
    End Property

    ''' <summary>
    ''' Sets and returns equivalent rows of all pages from both excel sheets
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property EquivalentRows() As Dictionary(Of Integer, Dictionary(Of Integer, Integer))
        Get
            Return objEquivalentRows
        End Get
    End Property

    ''' <summary>
    ''' Returns comparison result of all pages in excel sheet 1
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property ValueResult_1() As Dictionary(Of Integer, Dictionary(Of Tuple(Of Integer, Integer), List(Of ValueError)))
        Get
            Return objValueResult_1
        End Get
    End Property

    ''' <summary>
    ''' Return comparison result of all pages in excel sheet 2
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property ValueResult_2() As Dictionary(Of Integer, Dictionary(Of Tuple(Of Integer, Integer), List(Of ValueError)))
        Get
            Return objValueResult_2
        End Get
    End Property

    ''' <summary>
    ''' Returns comparison result by format of all pages based on equivalent columns and rows
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property FormatResult() As Dictionary(Of Integer, Dictionary(Of Tuple(Of Integer, Integer), List(Of FormatError)))
        Get
            Return objFormatResult
        End Get
    End Property

    ''' <summary>
    ''' Compare two excel sheets by page
    ''' Top-down,left-right comparison approach
    ''' </summary>
    ''' <param name="p_objWorkSheet_1">Excel File 1</param>
    ''' <param name="p_objWorkSheet_2">Excel File 2</param>
    Public Sub Compare(ByRef p_objWorkSheet_1 As Worksheet, ByRef p_objWorkSheet_2 As Worksheet, ByRef p_backgroundWorker As BackgroundWorker, ByRef e As System.ComponentModel.DoWorkEventArgs)

        If p_objWorkSheet_1 Is Nothing OrElse p_objWorkSheet_2 Is Nothing Then
            'Error when no instance on either worksheets was found
            Throw New Exception("No instances of worksheet is found.")
            Exit Sub
        End If

        '********************Start of Comparison*********************
        objExcelData_1 = New Dictionary(Of Integer, Dictionary(Of Tuple(Of Integer, Integer), Range))
        objExcelData_2 = New Dictionary(Of Integer, Dictionary(Of Tuple(Of Integer, Integer), Range))

        objEquivalentColumns = New Dictionary(Of Integer, Dictionary(Of Integer, Integer))
        objEquivalentRows = New Dictionary(Of Integer, Dictionary(Of Integer, Integer))

        objValueResult_1 = New Dictionary(Of Integer, Dictionary(Of Tuple(Of Integer, Integer), List(Of ValueError)))
        objValueResult_2 = New Dictionary(Of Integer, Dictionary(Of Tuple(Of Integer, Integer), List(Of ValueError)))
        objFormatResult = New Dictionary(Of Integer, Dictionary(Of Tuple(Of Integer, Integer), List(Of FormatError)))

        'Loop through all pages
        For w_intCtr_1 As Integer = 1 To intNoOfPages

            If p_backgroundWorker.CancellationPending = True Then
                e.Cancel = True
                Return
            End If

            Dim w_intCurrentStep As Integer = 1
            GetExcelData(p_objWorkSheet_1, objExcelData_1, objLocation_1, w_intCtr_1)
            GetExcelData(p_objWorkSheet_2, objExcelData_2, objLocation_2, w_intCtr_1)

            If objExcelData_1 Is Nothing OrElse objExcelData_2 Is Nothing Then
                'No data to compare
                Exit Sub
            End If

            objCompareByData = New Compare_Data

            'Compare value of excelsheets
            With objCompareByData
                'Set threshold, excel data, and location of page to compare
                .SetThreshold = dblThreshold
                .CompareToBestMatchData = blnBestMatchFlg

                .SetExcelData_1 = objExcelData_1(w_intCtr_1)
                .SetLocation_1 = objLocation_1(w_intCtr_1)

                .SetExcelData_2 = objExcelData_2(w_intCtr_1)
                .SetLocation_2 = objLocation_2(w_intCtr_1)

                If objRemoveCol Is Nothing = False AndAlso objRemoveCol.ContainsKey(w_intCtr_1) Then
                    .SetRemovedColumn = objRemoveCol(w_intCtr_1)
                End If
                If objRemoveRow Is Nothing = False AndAlso objRemoveRow.ContainsKey(w_intCtr_1) Then
                    .SetRemovedRow = objRemoveRow(w_intCtr_1)
                End If
                If objAddCol Is Nothing = False AndAlso objAddCol.ContainsKey(w_intCtr_1) Then
                    .SetAddedColumn = objAddCol(w_intCtr_1)
                End If
                If objAddRow Is Nothing = False AndAlso objAddRow.ContainsKey(w_intCtr_1) Then
                    .SetAddedRow = objAddRow(w_intCtr_1)
                End If

                If objChangeData Is Nothing = False AndAlso objChangeData.ContainsKey(w_intCtr_1) Then
                    .SetDataChange = objChangeData(w_intCtr_1)
                End If

                If p_backgroundWorker.CancellationPending = True Then
                    e.Cancel = True
                    Return
                End If

                'Proceed to compare
                .Compare()

                objEquivalentColumns.Add(w_intCtr_1, .EquivalentColumns)
                objEquivalentRows.Add(w_intCtr_1, .EquivalentRows)

                objValueResult_1.Add(w_intCtr_1, .ExcelData_Result_1)
                objValueResult_2.Add(w_intCtr_1, .ExcelData_Result_2)
            End With

            If blnCompareMerge OrElse blnCompareTextWrap OrElse blnCompareTextAlign OrElse blnCompareOrientation OrElse blnCompareBorder OrElse blnCompareBackColor OrElse blnCompareFont Then
                objCompareByFormat = New Compare_Format

                'Compare format of excelsheets
                With objCompareByFormat
                    'Set excel data to compare
                    .SetExcelFormat_1 = objExcelData_1(w_intCtr_1)
                    .SetExcelFormat_2 = objExcelData_2(w_intCtr_1)
                    'Set equivalent columns of page retrieved from comparing values of both excel sheets
                    'Set equivalent rows of page retrieved from comparing values of both excel sheets
                    If objEquivalentColumns Is Nothing = False AndAlso objEquivalentColumns.ContainsKey(w_intCtr_1) Then
                        .SetEquivalentColumns = objEquivalentColumns(w_intCtr_1)
                    End If
                    If objEquivalentRows Is Nothing = False AndAlso objEquivalentRows.ContainsKey(w_intCtr_1) Then
                        .SetEquivalentRows = objEquivalentRows(w_intCtr_1)
                    End If

                    .CompareMerge = blnCompareMerge
                    .CompareTextWrap = blnCompareTextWrap
                    .CompareTextAlign = blnCompareTextAlign
                    .CompareOrientation = blnCompareOrientation
                    .CompareBorder = blnCompareBorder
                    .CompareBackColor = blnCompareBackColor
                    .CompareFont = blnCompareFont

                    If p_backgroundWorker.CancellationPending = True Then
                        e.Cancel = True
                        Return
                    End If

                    .Compare()

                    'Set comparison result of page to collection
                    objFormatResult.Add(w_intCtr_1, .ExcelFormat_Result)
                End With
            End If

            If p_backgroundWorker.CancellationPending = True Then
                e.Cancel = True
                Return
            End If

            'Set result to excel sheets
            AddValueResultToWorkSheet(p_objWorkSheet_1, p_objWorkSheet_2, w_intCtr_1)

            If p_backgroundWorker.CancellationPending = True Then
                e.Cancel = True
                Return
            End If

            'Set result to excel sheets
            AddFormatResultToWorkSheet(p_objWorkSheet_2, w_intCtr_1)

            p_backgroundWorker.ReportProgress((100 / intNoOfPages) * w_intCtr_1, (100 / intNoOfPages) * w_intCtr_1 & "% Completed " & w_intCtr_1 & " out of " & intNoOfPages & " pages")
            Thread.Sleep(3000)
        Next
    End Sub

    ''' <summary>
    ''' Get excel data of worksheet
    ''' </summary>
    ''' <param name="p_objWorksheet"></param>
    ''' <param name="p_objExcelData"></param>
    ''' <param name="p_objLocation"></param>
    ''' <param name="p_intPage"></param>
    Private Sub GetExcelData(ByRef p_objWorksheet As Worksheet, ByRef p_objExcelData As Dictionary(Of Integer, Dictionary(Of Tuple(Of Integer, Integer), Range)) _
                             , ByRef p_objLocation As Dictionary(Of Integer, Location), ByVal p_intPage As Integer)
        Dim w_objExcelData As Dictionary(Of Tuple(Of Integer, Integer), Range)
        Dim w_objRowColKey As Tuple(Of Integer, Integer)

        With p_objWorksheet
            If p_objLocation.ContainsKey(p_intPage) Then
                w_objExcelData = New Dictionary(Of Tuple(Of Integer, Integer), Range)

                For w_intCol As Integer = p_objLocation(p_intPage).intFromCol To p_objLocation(p_intPage).intToCol

                    For w_intRow As Integer = p_objLocation(p_intPage).intFromRow To p_objLocation(p_intPage).intToRow
                        'Add excel data to collection
                        w_objRowColKey = New Tuple(Of Integer, Integer)(w_intRow, w_intCol)
                        w_objExcelData.Add(w_objRowColKey, .Cells(w_intRow, w_intCol))
                    Next
                Next

                If p_objExcelData.ContainsKey(p_intPage) = False Then
                    p_objExcelData.Add(p_intPage, w_objExcelData)
                End If
            End If
        End With

    End Sub

    ''' <summary>
    ''' Edit worksheet to set comparison result after compare by value
    ''' </summary>
    ''' <param name="p_objWorksheet_1"></param>
    ''' <param name="p_objWorksheet_2"></param>
    ''' <param name="p_intPage">Page to set comparison result</param>
    Private Sub AddValueResultToWorkSheet(ByRef p_objWorksheet_1 As Worksheet, ByRef p_objWorksheet_2 As Worksheet _
                                          , ByVal p_intPage As Integer)
        Dim w_objRowColKey As Tuple(Of Integer, Integer)
        Dim w_objResult As List(Of ValueError)
        Dim w_strComment As String = String.Empty
        Dim w_objExistingComment As Comment

        Try
            If objValueResult_1.Count = 0 Then
            Else
                p_objWorksheet_1.Unprotect()

                For w_intCol As Integer = objLocation_1(p_intPage).intFromCol To objLocation_1(p_intPage).intToCol

                    For w_intRow As Integer = objLocation_1(p_intPage).intFromRow To objLocation_1(p_intPage).intToRow
                        With p_objWorksheet_1
                            w_objRowColKey = New Tuple(Of Integer, Integer)(w_intRow, w_intCol)
                            If objValueResult_1(p_intPage)(w_objRowColKey).Count = 0 Then
                            Else
                                w_objResult = objValueResult_1(p_intPage)(w_objRowColKey)

                                For Each objResult As ValueError In w_objResult
                                    Select Case objResult
                                        Case ValueError.MissingColumn
                                            .Cells(w_intRow, w_intCol).Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightYellow)

                                            If w_intRow = objLocation_1(p_intPage).intFromRow Then
                                                'Add comment to first row of the column

                                                w_objExistingComment = .Cells(w_intRow, w_intCol).Comment
                                                If w_objExistingComment Is Nothing Then
                                                    w_strComment = objResult.ToString
                                                Else
                                                    w_strComment = w_objExistingComment.Shape.TextFrame.Characters().Text & vbNewLine & objResult.ToString
                                                    .Cells(w_intRow, w_intCol).Comment.Delete()
                                                End If
                                                .Cells(w_intRow, w_intCol).AddComment(w_strComment)
                                                .Cells(w_intRow, w_intCol).Select()
                                            End If
                                        Case ValueError.MissingRow
                                            .Cells(w_intRow, w_intCol).Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightGreen)

                                            If w_intCol = objLocation_1(p_intPage).intFromCol Then
                                                w_objExistingComment = .Cells(w_intRow, w_intCol).Comment
                                                If w_objExistingComment Is Nothing Then
                                                    w_strComment = objResult.ToString
                                                Else
                                                    w_strComment = w_objExistingComment.Shape.TextFrame.Characters().Text & vbNewLine & objResult.ToString
                                                    .Cells(w_intRow, w_intCol).Comment.Delete()
                                                End If
                                                'Add comment to first row of the column
                                                .Cells(w_intRow, w_intCol).AddComment(w_strComment)
                                                .Cells(w_intRow, w_intCol).Select()
                                            End If
                                    End Select
                                Next

                            End If

                        End With
                    Next
                Next

                p_objWorksheet_1.Protect()
            End If

            If objValueResult_2.Count = 0 Then
            Else
                p_objWorksheet_2.Unprotect()
                For w_intCol As Integer = objLocation_2(p_intPage).intFromCol To objLocation_2(p_intPage).intToCol
                    For w_intRow As Integer = objLocation_2(p_intPage).intFromRow To objLocation_2(p_intPage).intToRow
                        With p_objWorksheet_2

                            w_objRowColKey = New Tuple(Of Integer, Integer)(w_intRow, w_intCol)
                            If objValueResult_2(p_intPage)(w_objRowColKey).Count = 0 Then
                            Else
                                w_objResult = objValueResult_2(p_intPage)(w_objRowColKey)

                                For Each objResult As ValueError In w_objResult
                                    Select Case objResult
                                        Case ValueError.AddedColumn
                                            .Cells(w_intRow, w_intCol).Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightYellow)

                                            If w_intRow = objLocation_2(p_intPage).intFromRow Then
                                                w_objExistingComment = .Cells(w_intRow, w_intCol).Comment
                                                If w_objExistingComment Is Nothing Then
                                                    w_strComment = objResult.ToString
                                                Else
                                                    w_strComment = w_objExistingComment.Shape.TextFrame.Characters().Text & vbNewLine & objResult.ToString
                                                    .Cells(w_intRow, w_intCol).Comment.Delete()
                                                End If
                                                'Add comment to first row of the column
                                                .Cells(w_intRow, w_intCol).AddComment(w_strComment)
                                                .Cells(w_intRow, w_intCol).Select()
                                            End If
                                        Case ValueError.AddedRow
                                            .Cells(w_intRow, w_intCol).Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightGreen)

                                            If w_intCol = objLocation_2(p_intPage).intFromCol Then
                                                w_objExistingComment = .Cells(w_intRow, w_intCol).Comment
                                                If w_objExistingComment Is Nothing Then
                                                    w_strComment = objResult.ToString
                                                Else
                                                    w_strComment = w_objExistingComment.Shape.TextFrame.Characters().Text & vbNewLine & objResult.ToString
                                                    .Cells(w_intRow, w_intCol).Comment.Delete()
                                                End If
                                                'Add comment to first row of the column
                                                .Cells(w_intRow, w_intCol).AddComment(w_strComment)
                                                .Cells(w_intRow, w_intCol).Select()
                                            End If
                                        Case ValueError.DataError
                                            .Cells(w_intRow, w_intCol).Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.Red)

                                            w_objExistingComment = .Cells(w_intRow, w_intCol).Comment
                                            If w_objExistingComment Is Nothing Then
                                                w_strComment = objResult.ToString
                                            Else
                                                w_strComment = w_objExistingComment.Shape.TextFrame.Characters().Text & vbNewLine & objResult.ToString
                                                .Cells(w_intRow, w_intCol).Comment.Delete()
                                            End If
                                            'Add comment to first row of the column
                                            .Cells(w_intRow, w_intCol).AddComment(w_strComment)
                                            .Cells(w_intRow, w_intCol).Select()
                                    End Select
                                Next

                            End If

                        End With
                    Next
                Next
                p_objWorksheet_2.Protect()
            End If

        Catch ex As Exception
            Throw
        End Try

    End Sub

    ''' <summary>
    ''' Edit worksheet to set comparison result after compare by format
    ''' </summary>
    ''' <param name="p_objWorkSheet"></param>
    ''' <param name="p_intPage">Page to set comparison result</param>
    Private Sub AddFormatResultToWorkSheet(ByRef p_objWorkSheet As Microsoft.Office.Interop.Excel.Worksheet, ByVal p_intPage As Integer)
        Dim w_strComments As String
        Dim w_objExistingComment As Comment

        If objFormatResult.Count = 0 Then
            Exit Sub
        End If

        p_objWorkSheet.Unprotect()

        For w_intCol As Integer = objLocation_2(p_intPage).intFromCol To objLocation_2(p_intPage).intToCol
            For w_intRow As Integer = objLocation_2(p_intPage).intFromRow To objLocation_2(p_intPage).intToRow
                If objFormatResult(p_intPage).ContainsKey(New Tuple(Of Integer, Integer)(w_intRow, w_intCol)) = False _
                    OrElse objFormatResult(p_intPage)(New Tuple(Of Integer, Integer)(w_intRow, w_intCol)) Is Nothing _
                    OrElse objFormatResult(p_intPage)(New Tuple(Of Integer, Integer)(w_intRow, w_intCol)).Count = 0 Then
                    Continue For
                End If

                w_strComments = String.Empty
                For Each objError As FormatError In objFormatResult(p_intPage)(New Tuple(Of Integer, Integer)(w_intRow, w_intCol))
                    w_strComments &= objError.ToString & vbNewLine
                Next

                With p_objWorkSheet
                    w_objExistingComment = .Cells(w_intRow, w_intCol).Comment
                    If w_objExistingComment Is Nothing Then
                    Else
                        w_strComments = w_objExistingComment.Shape.TextFrame.Characters().Text & vbNewLine & w_strComments
                        .Cells(w_intRow, w_intCol).Comment.Delete()
                    End If
                    .Cells(w_intRow, w_intCol).AddComment(w_strComments)
                    .Cells(w_intRow, w_intCol).Select()
                End With
            Next
        Next

        p_objWorkSheet.Protect()
    End Sub
End Class
