Imports Microsoft.Office.Interop.Excel
Imports Utility

Friend Class Compare_Format
    Private objExcel_Format_1 As Dictionary(Of Tuple(Of Integer, Integer), Range)
    Private objExcel_Format_2 As Dictionary(Of Tuple(Of Integer, Integer), Range)

    'Key: From Excel data 1
    'Value: From Excel data 2
    Private objEquivalentColumns As Dictionary(Of Integer, Integer)
    Private objEquivalentRows As Dictionary(Of Integer, Integer)

    Private blnCompareMerge As Boolean
    Private blnCompareTextWrap As Boolean
    Private blnCompareTextAgrn As Boolean
    Private blnCompareOrientation As Boolean
    Private blnCompareBorder As Boolean
    Private blnCompareBackColor As Boolean
    Private blnCompareFont As Boolean

    Private objExcel_Result As Dictionary(Of Tuple(Of Integer, Integer), List(Of FormatError))

    Public WriteOnly Property SetExcelFormat_1() As Dictionary(Of Tuple(Of Integer, Integer), Range)
        Set(value As Dictionary(Of Tuple(Of Integer, Integer), Range))
            objExcel_Format_1 = value
        End Set
    End Property

    Public WriteOnly Property SetExcelFormat_2() As Dictionary(Of Tuple(Of Integer, Integer), Range)
        Set(value As Dictionary(Of Tuple(Of Integer, Integer), Range))
            objExcel_Format_2 = value
        End Set
    End Property

    Public WriteOnly Property SetEquivalentColumns() As Dictionary(Of Integer, Integer)
        Set(value As Dictionary(Of Integer, Integer))
            objEquivalentColumns = value
        End Set
    End Property

    Public WriteOnly Property SetEquivalentRows() As Dictionary(Of Integer, Integer)
        Set(value As Dictionary(Of Integer, Integer))
            objEquivalentRows = value
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

    Public ReadOnly Property ExcelFormat_Result() As Dictionary(Of Tuple(Of Integer, Integer), List(Of FormatError))
        Get
            Return objExcel_Result
        End Get
    End Property

    Public Sub Compare()
        Dim w_intCol_1 As Integer
        Dim w_intCol_2 As Integer

        Dim w_intRow_1 As Integer
        Dim w_intRow_2 As Integer

        Dim w_objRowColKey_1 As Tuple(Of Integer, Integer)
        Dim w_objRowColKey_2 As Tuple(Of Integer, Integer)

        Dim w_objListFormatError As List(Of FormatError)

        Try
            If objExcel_Format_1 Is Nothing OrElse objExcel_Format_2 Is Nothing Then
                'Error when no data or insufficient data to compare
                Throw New Exception("No data / Insufficient data to compare.")
                Exit Sub
            End If

            objExcel_Result = New Dictionary(Of Tuple(Of Integer, Integer), List(Of FormatError))

            'Loop through equivalent columns
            For Each objKeyValPair_Col As KeyValuePair(Of Integer, Integer) In objEquivalentColumns
                For Each objKeyValPair_Row As KeyValuePair(Of Integer, Integer) In objEquivalentRows
                    w_intCol_1 = objKeyValPair_Col.Key
                    w_intRow_1 = objKeyValPair_Row.Key

                    w_intCol_2 = objKeyValPair_Col.Value
                    w_intRow_2 = objKeyValPair_Row.Value

                    w_objRowColKey_1 = New Tuple(Of Integer, Integer)(w_intRow_1, w_intCol_1)
                    w_objRowColKey_2 = New Tuple(Of Integer, Integer)(w_intRow_2, w_intCol_2)

                    w_objListFormatError = New List(Of FormatError)

                    'Merging Error
                    If blnCompareMerge AndAlso objExcel_Format_1(w_objRowColKey_1).MergeCells <>
                        objExcel_Format_2(w_objRowColKey_2).MergeCells Then
                        w_objListFormatError.Add(FormatError.MergingError)
                    End If

                    'Text Alignment Error
                    If blnCompareTextAlign AndAlso objExcel_Format_1(w_objRowColKey_1).HorizontalAlignment <> objExcel_Format_2(w_objRowColKey_2).HorizontalAlignment _
                        OrElse objExcel_Format_1(w_objRowColKey_1).VerticalAlignment <> objExcel_Format_2(w_objRowColKey_2).VerticalAlignment Then
                        w_objListFormatError.Add(FormatError.TextAlignmentError)
                    End If

                    'Font Error
                    If blnCompareFont AndAlso objExcel_Format_1(w_objRowColKey_1).Font.Name <> objExcel_Format_2(w_objRowColKey_2).Font.Name _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Font.FontStyle <> objExcel_Format_2(w_objRowColKey_2).Font.FontStyle _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Font.Color <> objExcel_Format_2(w_objRowColKey_2).Font.Color _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Font.Size <> objExcel_Format_2(w_objRowColKey_2).Font.Size _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Font.Underline <> objExcel_Format_2(w_objRowColKey_2).Font.Underline _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Font.Strikethrough <> objExcel_Format_2(w_objRowColKey_2).Font.Strikethrough _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Font.Subscript <> objExcel_Format_2(w_objRowColKey_2).Font.Subscript _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Font.Superscript <> objExcel_Format_2(w_objRowColKey_2).Font.Superscript Then
                        w_objListFormatError.Add(FormatError.FontError)
                    End If

                    'Border Error
                    If blnCompareBorder AndAlso objExcel_Format_1(w_objRowColKey_1).Borders(XlBordersIndex.xlDiagonalDown).LineStyle <> objExcel_Format_2(w_objRowColKey_2).Borders(XlBordersIndex.xlDiagonalDown).LineStyle _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Borders(XlBordersIndex.xlDiagonalDown).Weight <> objExcel_Format_2(w_objRowColKey_2).Borders(XlBordersIndex.xlDiagonalDown).Weight _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Borders(XlBordersIndex.xlDiagonalDown).Color <> objExcel_Format_2(w_objRowColKey_2).Borders(XlBordersIndex.xlDiagonalDown).Color _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Borders(XlBordersIndex.xlDiagonalUp).LineStyle <> objExcel_Format_2(w_objRowColKey_2).Borders(XlBordersIndex.xlDiagonalUp).LineStyle _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Borders(XlBordersIndex.xlDiagonalUp).Weight <> objExcel_Format_2(w_objRowColKey_2).Borders(XlBordersIndex.xlDiagonalUp).Weight _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Borders(XlBordersIndex.xlDiagonalUp).Color <> objExcel_Format_2(w_objRowColKey_2).Borders(XlBordersIndex.xlDiagonalUp).Color _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Borders(XlBordersIndex.xlEdgeTop).LineStyle <> objExcel_Format_2(w_objRowColKey_2).Borders(XlBordersIndex.xlEdgeTop).LineStyle _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Borders(XlBordersIndex.xlEdgeTop).Weight <> objExcel_Format_2(w_objRowColKey_2).Borders(XlBordersIndex.xlEdgeTop).Weight _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Borders(XlBordersIndex.xlEdgeTop).Color <> objExcel_Format_2(w_objRowColKey_2).Borders(XlBordersIndex.xlEdgeTop).Color _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Borders(XlBordersIndex.xlEdgeBottom).LineStyle <> objExcel_Format_2(w_objRowColKey_2).Borders(XlBordersIndex.xlEdgeBottom).LineStyle _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Borders(XlBordersIndex.xlEdgeBottom).Weight <> objExcel_Format_2(w_objRowColKey_2).Borders(XlBordersIndex.xlEdgeBottom).Weight _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Borders(XlBordersIndex.xlEdgeBottom).Color <> objExcel_Format_2(w_objRowColKey_2).Borders(XlBordersIndex.xlEdgeBottom).Color _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Borders(XlBordersIndex.xlEdgeLeft).LineStyle <> objExcel_Format_2(w_objRowColKey_2).Borders(XlBordersIndex.xlEdgeLeft).LineStyle _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Borders(XlBordersIndex.xlEdgeLeft).Weight <> objExcel_Format_2(w_objRowColKey_2).Borders(XlBordersIndex.xlEdgeLeft).Weight _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Borders(XlBordersIndex.xlEdgeLeft).Color <> objExcel_Format_2(w_objRowColKey_2).Borders(XlBordersIndex.xlEdgeLeft).Color _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Borders(XlBordersIndex.xlEdgeRight).LineStyle <> objExcel_Format_2(w_objRowColKey_2).Borders(XlBordersIndex.xlEdgeRight).LineStyle _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Borders(XlBordersIndex.xlEdgeRight).Weight <> objExcel_Format_2(w_objRowColKey_2).Borders(XlBordersIndex.xlEdgeRight).Weight _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Borders(XlBordersIndex.xlEdgeRight).Color <> objExcel_Format_2(w_objRowColKey_2).Borders(XlBordersIndex.xlEdgeRight).Color _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Borders(XlBordersIndex.xlInsideHorizontal).LineStyle <> objExcel_Format_2(w_objRowColKey_2).Borders(XlBordersIndex.xlInsideHorizontal).LineStyle _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Borders(XlBordersIndex.xlInsideHorizontal).Weight <> objExcel_Format_2(w_objRowColKey_2).Borders(XlBordersIndex.xlInsideHorizontal).Weight _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Borders(XlBordersIndex.xlInsideHorizontal).Color <> objExcel_Format_2(w_objRowColKey_2).Borders(XlBordersIndex.xlInsideHorizontal).Color _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Borders(XlBordersIndex.xlInsideVertical).LineStyle <> objExcel_Format_2(w_objRowColKey_2).Borders(XlBordersIndex.xlInsideVertical).LineStyle _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Borders(XlBordersIndex.xlInsideVertical).Weight <> objExcel_Format_2(w_objRowColKey_2).Borders(XlBordersIndex.xlInsideVertical).Weight _
                        OrElse objExcel_Format_1(w_objRowColKey_1).Borders(XlBordersIndex.xlInsideVertical).Color <> objExcel_Format_2(w_objRowColKey_2).Borders(XlBordersIndex.xlInsideVertical).Color Then
                        w_objListFormatError.Add(FormatError.BorderError)
                    End If

                    'Background Color Error
                    If blnCompareBackColor AndAlso objExcel_Format_1(w_objRowColKey_1).Interior.Color <> objExcel_Format_2(w_objRowColKey_2).Interior.Color Then
                        w_objListFormatError.Add(FormatError.BackColorError)
                    End If

                    'Cell Error
                    If blnCompareTextWrap AndAlso objExcel_Format_1(w_objRowColKey_1).WrapText <> objExcel_Format_2(w_objRowColKey_2).WrapText _
                        OrElse objExcel_Format_1(w_objRowColKey_1).ShrinkToFit <> objExcel_Format_2(w_objRowColKey_2).ShrinkToFit Then
                        w_objListFormatError.Add(FormatError.TextWrapError)
                    End If

                    If blnCompareOrientation AndAlso objExcel_Format_1(w_objRowColKey_1).Orientation <> objExcel_Format_2(w_objRowColKey_2).Orientation Then
                        w_objListFormatError.Add(FormatError.OrientationError)
                    End If

                    objExcel_Result.Add(w_objRowColKey_2, w_objListFormatError)
                Next
            Next
        Catch ex As Exception
            Throw
        End Try

    End Sub
End Class
