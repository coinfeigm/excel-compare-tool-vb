Imports Microsoft.Office.Interop.Excel
Imports Utility

Friend Class Compare_Data
    'Threshold in determining similarity of columns
    Private dblThreshold As Double

    Private blnBestMatchFlg As Boolean

    'Key: Cell[Row, Column]
    'Value: Properties of the cell
    Private objExcel_Data_1 As Dictionary(Of Tuple(Of Integer, Integer), Range)
    Private objExcel_Data_2 As Dictionary(Of Tuple(Of Integer, Integer), Range)

    Private objDataDictionary As Dictionary(Of String, Integer)

    'Key: Column number
    'Value: List of data[key-based] under column
    Private objColumns_1 As Dictionary(Of Integer, List(Of Integer))
    Private objColumns_2 As Dictionary(Of Integer, List(Of Integer))

    'Key: Column number from excel sheet 1
    'Value: Columns from excel sheet 2 and similarity index
    Private objColumn_1 As SortedDictionary(Of Integer, List(Of Tuple(Of Integer, Double)))
    'Key: Column number from excel sheet 2
    'Value: Columns from excel sheet 1 and similarity index
    Private objColumn_2 As SortedDictionary(Of Integer, List(Of Tuple(Of Integer, Double)))

    'Key: Row from excel sheet 1
    'Value: Row from excel sheet 2 and counter for similar data
    Private objRows As SortedDictionary(Of Integer, List(Of Tuple(Of Integer, Integer)))

    'Location of data in both excel sheets
    Private objLocation_1 As Location
    Private objLocation_2 As Location

    Private objRemoveCol As List(Of Integer)
    Private objAddCol As List(Of Integer)
    Private objRemoveRow As List(Of Integer)
    Private objAddRow As List(Of Integer)

    'Key: Data from excel file 1
    'Value: Changed data from excel file 2
    Private objChangeData As Dictionary(Of String, String)
    Private objChangeData_Keys As Dictionary(Of Integer, Integer)

    Private intNoOfRows_1 As Integer
    Private intNoOfCols_1 As Integer

    Private intNoOfRows_2 As Integer
    Private intNoOfCols_2 As Integer

    'Item 1: Column/Row from excel sheet 1
    'Item 2: Column/Row from excel sheet 2
    Private objEquivalentColumns As Dictionary(Of Integer, Integer)
    Private objEquivalentRows As Dictionary(Of Integer, Integer)

    'Key: Cell[Row, Column]
    'Value: List of Error
    Private objExcel_Result_1 As Dictionary(Of Tuple(Of Integer, Integer), List(Of ValueError))
    Private objExcel_Result_2 As Dictionary(Of Tuple(Of Integer, Integer), List(Of ValueError))

    'Distance cost and direction
    Private Structure Distance
        Dim intDist As Short
        Dim enmDirection As Direction
    End Structure

    'Enum for backtracking of matrix
    Private Enum Direction
        Up
        Left
        Diagonal
    End Enum

    ''' <summary>
    ''' Set threshold for comparison
    ''' </summary>
    Public WriteOnly Property SetThreshold() As Double
        Set(value As Double)
            dblThreshold = value
        End Set
    End Property

    ''' <summary>
    ''' Set flag if to compare each column/row to column/row most similar to it
    ''' </summary>
    Public WriteOnly Property CompareToBestMatchData() As Boolean
        Set(value As Boolean)
            blnBestMatchFlg = value
        End Set
    End Property

    ''' <summary>
    ''' Set excel data from excel sheet 1
    ''' </summary>
    Public WriteOnly Property SetExcelData_1() As Dictionary(Of Tuple(Of Integer, Integer), Range)
        Set(ByVal value As Dictionary(Of Tuple(Of Integer, Integer), Range))
            objExcel_Data_1 = value
        End Set
    End Property

    ''' <summary>
    ''' Set location of data in excel sheet 1
    ''' </summary>
    Public WriteOnly Property SetLocation_1() As Location
        Set(value As Location)
            objLocation_1 = value
        End Set
    End Property

    ''' <summary>
    ''' Set excel data from excel sheet 2
    ''' </summary>
    Public WriteOnly Property SetExcelData_2() As Dictionary(Of Tuple(Of Integer, Integer), Range)
        Set(ByVal value As Dictionary(Of Tuple(Of Integer, Integer), Range))
            objExcel_Data_2 = value
        End Set
    End Property

    ''' <summary>
    ''' Set location of data in excel sheet 2
    ''' </summary>
    Public WriteOnly Property SetLocation_2() As Location
        Set(value As Location)
            objLocation_2 = value
        End Set
    End Property

    ''' <summary>
    ''' Set removed column from excel sheet 1
    ''' </summary>
    Public WriteOnly Property SetRemovedColumn() As List(Of Integer)
        Set(value As List(Of Integer))
            objRemoveCol = value
            objRemoveCol.Sort()
        End Set
    End Property

    ''' <summary>
    ''' Set removed row from excel sheet 1
    ''' </summary>
    Public WriteOnly Property SetRemovedRow() As List(Of Integer)
        Set(value As List(Of Integer))
            objRemoveRow = value
            objRemoveRow.Sort()
        End Set
    End Property

    ''' <summary>
    ''' Set added column from excel sheet 2
    ''' </summary>
    Public WriteOnly Property SetAddedColumn() As List(Of Integer)
        Set(value As List(Of Integer))
            objAddCol = value
            objAddCol.Sort()
        End Set
    End Property

    ''' <summary>
    ''' Set added row from excel sheet 2
    ''' </summary>
    Public WriteOnly Property SetAddedRow() As List(Of Integer)
        Set(value As List(Of Integer))
            objAddRow = value
            objAddRow.Sort()
        End Set
    End Property

    Public WriteOnly Property SetDataChange() As Dictionary(Of String, String)
        Set(value As Dictionary(Of String, String))
            objChangeData = value
        End Set
    End Property

    ''' <summary>
    ''' Returns equivalent columns in excel data compared
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property EquivalentColumns() As Dictionary(Of Integer, Integer)
        Get
            Return objEquivalentColumns
        End Get
    End Property

    ''' <summary>
    ''' Returns equivalent rows in excel data compared
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property EquivalentRows() As Dictionary(Of Integer, Integer)
        Get
            Return objEquivalentRows
        End Get
    End Property

    ''' <summary>
    ''' Returns result after comparison for excel data 1
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property ExcelData_Result_1() As Dictionary(Of Tuple(Of Integer, Integer), List(Of ValueError))
        Get
            Return objExcel_Result_1
        End Get
    End Property

    ''' <summary>
    ''' Returns result after comparison for excel data 2
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property ExcelData_Result_2() As Dictionary(Of Tuple(Of Integer, Integer), List(Of ValueError))
        Get
            Return objExcel_Result_2
        End Get
    End Property

    ''' <summary>
    ''' Compare excel data set
    ''' </summary>
    Public Sub Compare()

        If objExcel_Data_1 Is Nothing OrElse objExcel_Data_2 Is Nothing Then
            'Error when no data or insufficient data to compare
            Throw New Exception("No data / Insufficient data to compare.")
            Exit Sub
        End If

        intNoOfRows_1 = objLocation_1.intToRow - objLocation_1.intFromRow + 1
        intNoOfCols_1 = objLocation_1.intToCol - objLocation_1.intFromCol + 1

        intNoOfRows_2 = objLocation_2.intToRow - objLocation_2.intFromRow + 1
        intNoOfCols_2 = objLocation_2.intToCol - objLocation_2.intFromCol + 1

        CreateDataDictionary()
        Set_DataChange()
        SetColumns()

        If blnBestMatchFlg Then
            'Match column/row to most similar column/row
            GetEquivalent_Columns()
            GetEquivalent_Rows()
        Else
            'Match column/row to immediate similar column/row
            GetColumns()
            GetRows()
        End If

        InitializeResult()

        SetError_Column(ValueError.MissingColumn)
        SetError_Column(ValueError.AddedColumn)

        SetError_Row(ValueError.MissingRow)
        SetError_Row(ValueError.AddedRow)
        SetError_Row(ValueError.DataError)

    End Sub

    ''' <summary>
    ''' List of data present in both excel sheet
    ''' </summary>
    Private Sub CreateDataDictionary()
        Dim w_intKey As Integer

        objDataDictionary = New Dictionary(Of String, Integer)

        'Loop through columns of excel data 1
        For w_intCol_1 As Integer = objLocation_1.intFromCol To objLocation_1.intToCol
            If objRemoveCol Is Nothing = False AndAlso objRemoveCol.Contains(w_intCol_1) Then
                Continue For
            End If

            For w_intRow_1 As Integer = objLocation_1.intFromRow To objLocation_1.intToRow

                If objRemoveRow Is Nothing = False AndAlso objRemoveRow.Contains(w_intRow_1) Then
                    Continue For
                End If

                If String.IsNullOrEmpty(objExcel_Data_1(New Tuple(Of Integer, Integer)(w_intRow_1, w_intCol_1)).Value) Then
                ElseIf objDataDictionary.ContainsKey(objExcel_Data_1(New Tuple(Of Integer, Integer)(w_intRow_1, w_intCol_1)).Value) = False Then
                    w_intKey += 1
                    objDataDictionary.Add(objExcel_Data_1(New Tuple(Of Integer, Integer)(w_intRow_1, w_intCol_1)).Value, w_intKey)
                End If
            Next
        Next

        'Loop through columns of excel data 2
        For w_intCol_2 As Integer = objLocation_2.intFromCol To objLocation_2.intToCol
            If objAddCol Is Nothing = False AndAlso objAddCol.Contains(w_intCol_2) Then
                Continue For
            End If

            For w_intRow_2 As Integer = objLocation_2.intFromRow To objLocation_2.intToRow

                If objAddRow Is Nothing = False AndAlso objAddRow.Contains(w_intRow_2) Then
                    Continue For
                End If

                If String.IsNullOrEmpty(objExcel_Data_2(New Tuple(Of Integer, Integer)(w_intRow_2, w_intCol_2)).Value) Then
                ElseIf objDataDictionary.ContainsKey(objExcel_Data_2(New Tuple(Of Integer, Integer)(w_intRow_2, w_intCol_2)).Value) = False Then
                    w_intKey += 1
                    objDataDictionary.Add(objExcel_Data_2(New Tuple(Of Integer, Integer)(w_intRow_2, w_intCol_2)).Value, w_intKey)
                End If
            Next
        Next

    End Sub

    ''' <summary>
    ''' Get keys from data dictionary of data changes for customization
    ''' </summary>
    Private Sub Set_DataChange()
        Dim w_intKey1 As Integer
        Dim w_intKey2 As Integer

        If objChangeData Is Nothing Then
            Exit Sub
        End If

        objChangeData_Keys = New Dictionary(Of Integer, Integer)

        For Each objData As KeyValuePair(Of String, String) In objChangeData
            If String.IsNullOrEmpty(objData.Key) Then
                w_intKey1 = -1
            ElseIf objDataDictionary Is Nothing = False AndAlso objDataDictionary.ContainsKey(objData.Key) Then
                w_intKey1 = objDataDictionary(objData.Key)
            End If

            If String.IsNullOrEmpty(objData.Value) Then
                w_intKey2 = -1
            ElseIf objDataDictionary Is Nothing = False AndAlso objDataDictionary.ContainsKey(objData.Value) Then
                w_intKey2 = objDataDictionary(objData.Value)
            End If

            If w_intKey1 <> 0 AndAlso w_intKey2 <> 0 Then
                objChangeData_Keys.Add(w_intKey1, w_intKey2)
            End If
        Next

    End Sub

    ''' <summary>
    ''' Get rows of columns to compare
    ''' </summary>
    Private Sub SetColumns()
        Dim w_objData As List(Of Integer)

        objColumns_1 = New Dictionary(Of Integer, List(Of Integer))

        'Loop through columns of excel data 1
        For w_intCol_1 As Integer = objLocation_1.intFromCol To objLocation_1.intToCol

            If objRemoveCol Is Nothing = False AndAlso objRemoveCol.Contains(w_intCol_1) Then
                Continue For
            End If

            w_objData = New List(Of Integer)
            For w_intRow_1 As Integer = objLocation_1.intFromRow To objLocation_1.intToRow

                If objRemoveRow Is Nothing = False AndAlso objRemoveRow.Contains(w_intRow_1) Then
                    Continue For
                End If

                If String.IsNullOrEmpty(objExcel_Data_1(New Tuple(Of Integer, Integer)(w_intRow_1, w_intCol_1)).Value) Then
                    w_objData.Add(-1)
                ElseIf objDataDictionary Is Nothing = False AndAlso objDataDictionary.ContainsKey(objExcel_Data_1(New Tuple(Of Integer, Integer)(w_intRow_1, w_intCol_1)).Value) Then
                    w_objData.Add(objDataDictionary(objExcel_Data_1(New Tuple(Of Integer, Integer)(w_intRow_1, w_intCol_1)).Value))
                End If
            Next

            objColumns_1.Add(w_intCol_1, w_objData)
        Next

        objColumns_2 = New Dictionary(Of Integer, List(Of Integer))

        'Loop through columns of excel data 2
        For w_intCol_2 As Integer = objLocation_2.intFromCol To objLocation_2.intToCol

            If objAddCol Is Nothing = False AndAlso objAddCol.Contains(w_intCol_2) Then
                Continue For
            End If

            w_objData = New List(Of Integer)
            For w_intRow_2 As Integer = objLocation_2.intFromRow To objLocation_2.intToRow

                If objAddRow Is Nothing = False AndAlso objAddRow.Contains(w_intRow_2) Then
                    Continue For
                End If

                If String.IsNullOrEmpty(objExcel_Data_2(New Tuple(Of Integer, Integer)(w_intRow_2, w_intCol_2)).Value) Then
                    w_objData.Add(-1)
                ElseIf objDataDictionary Is Nothing = False AndAlso objDataDictionary.ContainsKey(objExcel_Data_2(New Tuple(Of Integer, Integer)(w_intRow_2, w_intCol_2)).Value) Then
                    w_objData.Add(objDataDictionary(objExcel_Data_2(New Tuple(Of Integer, Integer)(w_intRow_2, w_intCol_2)).Value))
                End If
            Next

            objColumns_2.Add(w_intCol_2, w_objData)
        Next
    End Sub

    ''' <summary>
    ''' Get equivalent columns of both excel data
    ''' Using Jaccard similarity index 
    ''' </summary>
    Private Sub GetEquivalent_Columns()
        Dim w_intProgressCtr As Integer 'Progress
        Dim w_dblSimilarityIndex As Double

        Dim w_objCols_1 As List(Of Tuple(Of Integer, Double))
        Dim w_objCols_2 As List(Of Tuple(Of Integer, Double))

        objColumn_1 = New SortedDictionary(Of Integer, List(Of Tuple(Of Integer, Double)))
        objColumn_2 = New SortedDictionary(Of Integer, List(Of Tuple(Of Integer, Double)))

        w_intProgressCtr = 0

        'Loop through columns of excel data 1
        For w_intCol_1 As Integer = objLocation_1.intFromCol To objLocation_1.intToCol
            If objRemoveCol Is Nothing = False AndAlso objRemoveCol.Contains(w_intCol_1) Then
                Continue For
            End If

            'Find equivalent column of excel data 1 to columns of excel data 2
            For w_intCol_2 As Integer = objLocation_2.intFromCol To objLocation_2.intToCol

                If objAddCol Is Nothing = False AndAlso objAddCol.Contains(w_intCol_2) Then
                    Continue For
                End If

                If objChangeData_Keys Is Nothing = False Then
                    For Each objData As KeyValuePair(Of Integer, Integer) In objChangeData_Keys
                        If objColumns_1(w_intCol_1).Contains(objData.Key) Then
                            For w_intCnt As Integer = 0 To objColumns_2(w_intCol_2).Count - 1
                                If objData.Value = objColumns_2(w_intCol_2)(w_intCnt) Then
                                    objColumns_2(w_intCol_2)(w_intCnt) = objData.Key
                                End If
                            Next
                        End If
                    Next
                End If

                'Get similarity index of both columns
                w_dblSimilarityIndex = CalcSimilarity(objColumns_1(w_intCol_1), objColumns_2(w_intCol_2))

                If w_dblSimilarityIndex >= dblThreshold Then
                    'If similarity index is equal or more than the set threshold

                    'Set as equivalent column of excel file 1
                    If objColumn_1.ContainsKey(w_intCol_1) Then
                        w_objCols_1 = objColumn_1(w_intCol_1)
                        w_objCols_1.Add(New Tuple(Of Integer, Double)(w_intCol_2, w_dblSimilarityIndex))
                        objColumn_1(w_intCol_1) = w_objCols_1
                    Else
                        w_objCols_1 = New List(Of Tuple(Of Integer, Double))
                        w_objCols_1.Add(New Tuple(Of Integer, Double)(w_intCol_2, w_dblSimilarityIndex))
                        objColumn_1.Add(w_intCol_1, w_objCols_1)
                    End If

                    'Set as equivalent column of excel file 2
                    If objColumn_2.ContainsKey(w_intCol_2) Then
                        w_objCols_2 = objColumn_2(w_intCol_2)
                        w_objCols_2.Add(New Tuple(Of Integer, Double)(w_intCol_1, w_dblSimilarityIndex))
                        objColumn_2(w_intCol_2) = w_objCols_2
                    Else
                        w_objCols_2 = New List(Of Tuple(Of Integer, Double))
                        w_objCols_2.Add(New Tuple(Of Integer, Double)(w_intCol_1, w_dblSimilarityIndex))
                        objColumn_2.Add(w_intCol_2, w_objCols_2)
                    End If
                End If
            Next
        Next

        objEquivalentColumns = New Dictionary(Of Integer, Integer)

        'Assign each column with its most similar column counterpart based on the similarity index
        For Each objKeyValuePair As KeyValuePair(Of Integer, List(Of Tuple(Of Integer, Double))) In objColumn_1

            If objEquivalentColumns.ContainsKey(objKeyValuePair.Key) Then
                Continue For
            End If

            w_objCols_1 = objKeyValuePair.Value
            w_objCols_1 = w_objCols_1.OrderByDescending(Function(Col) Col.Item2).ThenBy(Function(Col) Col.Item1).ToList

            For Each objTuple_1 As Tuple(Of Integer, Double) In w_objCols_1

                If objEquivalentColumns.ContainsKey(objKeyValuePair.Key) Then
                    Exit For
                End If

                If objEquivalentColumns.ContainsValue(objTuple_1.Item1) Then
                    Continue For
                End If

                w_objCols_2 = objColumn_2(objTuple_1.Item1)
                w_objCols_2 = w_objCols_2.OrderByDescending(Function(Col) Col.Item2).ThenBy(Function(Col) Col.Item1).ToList

                For Each objTuple_2 As Tuple(Of Integer, Double) In w_objCols_2

                    If objEquivalentColumns.ContainsValue(objTuple_1.Item1) OrElse objEquivalentColumns.ContainsKey(objTuple_2.Item1) Then
                        Continue For
                    End If

                    If objTuple_1.Item2 >= objColumn_1(objTuple_2.Item1)(0).Item2 Then
                        objEquivalentColumns.Add(objKeyValuePair.Key, objTuple_1.Item1)
                    End If

                    Exit For
                Next
            Next
        Next

    End Sub

    ''' <summary>
    ''' Get all columns
    ''' </summary>
    Private Sub GetColumns()
        Dim w_objCols_1 As List(Of Integer)
        Dim w_objCols_2 As List(Of Integer)

        objEquivalentColumns = New Dictionary(Of Integer, Integer)

        w_objCols_1 = New List(Of Integer)
        For w_intCol_1 As Integer = objLocation_1.intFromCol To objLocation_1.intToCol
            If objRemoveCol Is Nothing = False AndAlso objRemoveCol.Contains(w_intCol_1) Then
                Continue For
            End If

            w_objCols_1.Add(w_intCol_1)
        Next

        w_objCols_2 = New List(Of Integer)
        For w_intCol_2 As Integer = objLocation_2.intFromCol To objLocation_2.intToCol
            If objAddCol Is Nothing = False AndAlso objAddCol.Contains(w_intCol_2) Then
                Continue For
            End If

            w_objCols_2.Add(w_intCol_2)
        Next

        If w_objCols_1.Count = 0 OrElse w_objCols_2.Count = 0 Then
            Exit Sub
        End If

        For w_intCnt As Integer = 0 To If(w_objCols_1.Count < w_objCols_2.Count, w_objCols_1.Count - 1, w_objCols_2.Count - 1)
            objEquivalentColumns.Add(w_objCols_1(w_intCnt), w_objCols_2(w_intCnt))
        Next

    End Sub

    ''' <summary>
    ''' Get equivalent rows of both excel data
    ''' </summary>
    Private Sub GetEquivalent_Rows()
        Dim w_intProgressCtr As Integer
        Dim w_intCtr_1 As Integer
        Dim w_intCtr_2 As Integer
        Dim w_objDistMatrix As Dictionary(Of Tuple(Of Integer, Integer), Distance)
        Dim w_objRows As List(Of Tuple(Of Integer, Integer))
        Dim w_intEquivalentRow As Integer
        Dim w_intMax As Integer
        Dim w_blnExistFlg As Boolean
        Dim w_strVal_1 As String
        Dim w_strVal_2 As String

        Try

            objRows = New SortedDictionary(Of Integer, List(Of Tuple(Of Integer, Integer)))

            w_intProgressCtr = 0
            'Loop through equivalent columns
            For Each objKeyValPair As KeyValuePair(Of Integer, Integer) In objEquivalentColumns

                'Get distance matrix
                w_objDistMatrix = GetDistanceMatrix(objKeyValPair.Key, objKeyValPair.Value)

                If w_objDistMatrix Is Nothing Then
                    Continue For
                End If

                w_intCtr_1 = intNoOfRows_1
                w_intCtr_2 = intNoOfRows_2

                'Backtracking of Distance Matrix
                While w_intCtr_1 > 0 AndAlso w_intCtr_2 > 0
                    Select Case w_objDistMatrix(New Tuple(Of Integer, Integer)(w_intCtr_1, w_intCtr_2)).enmDirection
                        Case Direction.Up
                            w_intCtr_2 -= 1
                        Case Direction.Left
                            w_intCtr_1 -= 1
                        Case Direction.Diagonal

                            If objRemoveRow Is Nothing = False AndAlso objRemoveRow.Contains((objLocation_1.intFromRow + w_intCtr_1) - 1) Then
                                w_intCtr_1 -= 1
                                w_intCtr_2 -= 1

                                Exit Select
                            End If

                            'Set which rows from excel data 2 are equivalent to excel data 1
                            If objRows.ContainsKey((objLocation_1.intFromRow + w_intCtr_1) - 1) Then
                                w_objRows = objRows((objLocation_1.intFromRow + w_intCtr_1) - 1)

                                w_blnExistFlg = False
                                For w_intCtr As Integer = 0 To w_objRows.Count - 1
                                    If w_objRows(w_intCtr).Item1 = (objLocation_2.intFromRow + w_intCtr_2) - 1 Then
                                        w_objRows(w_intCtr) = New Tuple(Of Integer, Integer)(w_objRows(w_intCtr).Item1, w_objRows(w_intCtr).Item2 + 1)
                                        w_blnExistFlg = True
                                    Else
                                        w_strVal_1 = If(objExcel_Data_1(New Tuple(Of Integer, Integer)((objLocation_1.intFromRow + w_intCtr_1) - 1, objKeyValPair.Key)).Value Is Nothing, String.Empty _
                                   , objExcel_Data_1(New Tuple(Of Integer, Integer)((objLocation_1.intFromRow + w_intCtr_1) - 1, objKeyValPair.Key)).Value)
                                        w_strVal_2 = If(objExcel_Data_2(New Tuple(Of Integer, Integer)((objLocation_2.intFromRow + w_intCtr_2) - 1, objKeyValPair.Value)).Value Is Nothing, String.Empty _
                                                   , objExcel_Data_2(New Tuple(Of Integer, Integer)((objLocation_2.intFromRow + w_intCtr_2) - 1, objKeyValPair.Value)).Value)

                                        If w_strVal_1.Equals(w_strVal_2) Then
                                            w_objRows(w_intCtr) = New Tuple(Of Integer, Integer)(w_objRows(w_intCtr).Item1, w_objRows(w_intCtr).Item2 + 1)
                                        End If

                                    End If
                                Next

                                If objAddRow Is Nothing = False AndAlso objAddRow.Contains((objLocation_2.intFromRow + w_intCtr_2) - 1) Then
                                    w_intCtr_1 -= 1
                                    w_intCtr_2 -= 1

                                    Exit Select
                                End If

                                If w_blnExistFlg = False Then
                                    w_objRows.Add(New Tuple(Of Integer, Integer)((objLocation_2.intFromRow + w_intCtr_2) - 1, 1))
                                End If

                                objRows((objLocation_1.intFromRow + w_intCtr_1) - 1) = w_objRows

                            Else
                                If objAddRow Is Nothing = False AndAlso objAddRow.Contains((objLocation_2.intFromRow + w_intCtr_2) - 1) Then
                                    w_intCtr_1 -= 1
                                    w_intCtr_2 -= 1

                                    Exit Select
                                End If

                                w_objRows = New List(Of Tuple(Of Integer, Integer))
                                w_objRows.Add(New Tuple(Of Integer, Integer)((objLocation_2.intFromRow + w_intCtr_2) - 1, 1))
                                objRows.Add((objLocation_1.intFromRow + w_intCtr_1) - 1, w_objRows)

                            End If

                            w_intCtr_1 -= 1
                            w_intCtr_2 -= 1
                    End Select
                End While
            Next

            objEquivalentRows = New Dictionary(Of Integer, Integer)

            w_intProgressCtr = 0

            'Loop to get most similar rows
            For Each objKeyValuePair As KeyValuePair(Of Integer, List(Of Tuple(Of Integer, Integer))) In objRows
                If objEquivalentRows.ContainsKey(objKeyValuePair.Key) Then
                    Continue For
                End If

                If objKeyValuePair.Value.Count > 0 Then
                    w_objRows = objKeyValuePair.Value

                    w_intEquivalentRow = w_objRows(0).Item1
                    w_intMax = w_objRows(0).Item2
                    For w_intCtr As Integer = 1 To w_objRows.Count - 1
                        If w_intMax <= w_objRows(w_intCtr).Item2 Then
                            w_intEquivalentRow = w_objRows(w_intCtr).Item1
                            w_intMax = w_objRows(w_intCtr).Item2
                        End If
                    Next

                    If objEquivalentRows.ContainsValue(w_intEquivalentRow) Then
                        Continue For
                    End If

                    If objEquivalentRows.ContainsKey(objKeyValuePair.Key) Then
                        objEquivalentRows(objKeyValuePair.Key) = w_intEquivalentRow
                    Else
                        objEquivalentRows.Add(objKeyValuePair.Key, w_intEquivalentRow)
                    End If
                End If
            Next

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Get rows
    ''' </summary>
    Private Sub GetRows()

        Dim w_objRows_1 As List(Of Integer)
        Dim w_objRows_2 As List(Of Integer)

        objEquivalentRows = New Dictionary(Of Integer, Integer)

        w_objRows_1 = New List(Of Integer)
        For w_intRow_1 As Integer = objLocation_1.intFromRow To objLocation_1.intToRow
            If objRemoveRow Is Nothing = False AndAlso objRemoveRow.Contains(w_intRow_1) Then
                Continue For
            End If

            w_objRows_1.Add(w_intRow_1)
        Next

        w_objRows_2 = New List(Of Integer)
        For w_intRow_2 As Integer = objLocation_2.intFromRow To objLocation_2.intToRow
            If objAddRow Is Nothing = False AndAlso objAddRow.Contains(w_intRow_2) Then
                Continue For
            End If

            w_objRows_2.Add(w_intRow_2)
        Next

        If w_objRows_1.Count = 0 OrElse w_objRows_2.Count = 0 Then
            Exit Sub
        End If

        For w_intCnt As Integer = 0 To If(w_objRows_1.Count < w_objRows_2.Count, w_objRows_1.Count - 1, w_objRows_2.Count - 1)
            objEquivalentRows.Add(w_objRows_1(w_intCnt), w_objRows_2(w_intCnt))
        Next

    End Sub

    ''' <summary>
    ''' Calculate similarity of columns using Jaccard similarity index
    ''' </summary>
    ''' <param name="p_objSet_1"></param>
    ''' <param name="p_objSet_2"></param>
    ''' <returns></returns>
    Private Function CalcSimilarity(ByVal p_objSet_1 As List(Of Integer), ByVal p_objSet_2 As List(Of Integer)) As Double
        Dim w_dblSimilarityIdx As Double
        Dim w_objSet_1 As HashSet(Of Integer)
        Dim w_objSet_2 As HashSet(Of Integer)

        If p_objSet_1 Is Nothing OrElse p_objSet_2 Is Nothing Then
            Return w_dblSimilarityIdx
            Exit Function
        End If

        w_objSet_1 = New HashSet(Of Integer)(p_objSet_1)
        w_objSet_2 = New HashSet(Of Integer)(p_objSet_2)

        w_dblSimilarityIdx = (Double.Parse(w_objSet_1.Intersect(w_objSet_2).Count()) / Double.Parse(w_objSet_1.Union(w_objSet_2).Count()))

        Return w_dblSimilarityIdx

    End Function

    ''' <summary>
    ''' Get distance matrix
    ''' Using Levenshtein Distance Algorithm
    ''' </summary>
    ''' <param name="p_intCol1_Idx">Column index of excel data 1</param>
    ''' <param name="p_intCol2_Idx">Column index of excel data 2</param>
    ''' <returns></returns>
    Private Function GetDistanceMatrix(ByVal p_intCol1_Idx As Integer, ByVal p_intCol2_Idx As Integer) As Dictionary(Of Tuple(Of Integer, Integer), Distance)
        Dim w_objDistMatrix As Dictionary(Of Tuple(Of Integer, Integer), Distance)
        Dim w_objRowKey As Tuple(Of Integer, Integer)
        Dim w_objDist As Distance

        Dim w_strValue_1 As String
        Dim w_strValue_2 As String
        Dim w_intCost As Integer

        'Initialize distance matrix
        w_objDistMatrix = New Dictionary(Of Tuple(Of Integer, Integer), Distance)

        'Create distance matrix
        w_objDist = New Distance() With {.intDist = 0, .enmDirection = Direction.Up}
        w_objRowKey = New Tuple(Of Integer, Integer)(0, 0)
        w_objDistMatrix.Add(w_objRowKey, w_objDist)

        For w_intRow_1 As Integer = 1 To intNoOfRows_1
            w_objDist = New Distance() With {.intDist = w_intRow_1, .enmDirection = Direction.Up}
            w_objRowKey = New Tuple(Of Integer, Integer)(w_intRow_1, 0)

            w_objDistMatrix.Add(w_objRowKey, w_objDist)
        Next

        For w_intRow_2 As Integer = 1 To intNoOfRows_2
            w_objDist = New Distance() With {.intDist = 0, .enmDirection = Direction.Left}
            w_objRowKey = New Tuple(Of Integer, Integer)(0, w_intRow_2)

            w_objDistMatrix.Add(w_objRowKey, w_objDist)
        Next

        For w_intRow_2 As Integer = 1 To intNoOfRows_2
            For w_intRow_1 As Integer = 1 To intNoOfRows_1
                w_strValue_1 = If(objExcel_Data_1(New Tuple(Of Integer, Integer)(objLocation_1.intFromRow + w_intRow_1 - 1, p_intCol1_Idx)).Value Is Nothing, String.Empty _
                                   , objExcel_Data_1(New Tuple(Of Integer, Integer)(objLocation_1.intFromRow + w_intRow_1 - 1, p_intCol1_Idx)).Value)
                w_strValue_2 = If(objExcel_Data_2(New Tuple(Of Integer, Integer)(objLocation_2.intFromRow + w_intRow_2 - 1, p_intCol2_Idx)).Value Is Nothing, String.Empty _
                                   , objExcel_Data_2(New Tuple(Of Integer, Integer)(objLocation_2.intFromRow + w_intRow_2 - 1, p_intCol2_Idx)).Value)

                If objChangeData Is Nothing = False AndAlso objChangeData.ContainsKey(w_strValue_1) Then
                    If w_strValue_2.Equals(objChangeData(w_strValue_1)) Then
                        w_intCost = 0
                    Else
                        w_intCost = 1
                    End If
                ElseIf w_strValue_1.Equals(w_strValue_2) Then
                    w_intCost = 0
                Else
                    w_intCost = 1
                End If

                w_objDist = GetMinDist_Direction(w_objDistMatrix(New Tuple(Of Integer, Integer)(w_intRow_1, w_intRow_2 - 1)).intDist _
                                                 , w_objDistMatrix(New Tuple(Of Integer, Integer)(w_intRow_1 - 1, w_intRow_2)).intDist _
                                                 , w_objDistMatrix(New Tuple(Of Integer, Integer)(w_intRow_1 - 1, w_intRow_2 - 1)).intDist, w_intCost)

                w_objRowKey = New Tuple(Of Integer, Integer)(w_intRow_1, w_intRow_2)
                w_objDistMatrix.Add(w_objRowKey, w_objDist)
            Next
        Next

        Return w_objDistMatrix
    End Function

    ''' <summary>
    ''' Get minimum distance and direction
    ''' </summary>
    ''' <param name="p_intUpVal">Value of index above the current index in the matrix</param>
    ''' <param name="p_intLeftVal">Value of index on the left side of the current index in the matrix</param>
    ''' <param name="p_intDiagVal">Value of index on the upper left side of the current index in the matrix</param>
    ''' <param name="p_intCost">Equivalent cost of the current index</param>
    ''' <returns></returns>
    Private Function GetMinDist_Direction(ByVal p_intUpVal As Short, ByVal p_intLeftVal As Short, ByVal p_intDiagVal As Short, ByVal p_intCost As Short) As Distance
        Dim w_objDist As Distance

        'Criteria to get minimum distance
        p_intUpVal += 1
        p_intLeftVal += 1
        p_intDiagVal += p_intCost

        w_objDist = New Distance

        With w_objDist
            .intDist = p_intUpVal
            .enmDirection = Direction.Up
        End With

        If p_intLeftVal < w_objDist.intDist Then
            With w_objDist
                .intDist = p_intLeftVal
                .enmDirection = Direction.Left
            End With
        End If

        If p_intDiagVal < w_objDist.intDist Then
            With w_objDist
                .intDist = p_intDiagVal
                .enmDirection = Direction.Diagonal
            End With
        End If

        Return w_objDist
    End Function

    ''' <summary>
    ''' Initialize comparison results to -1[No Error]
    ''' </summary>
    Private Sub InitializeResult()

        objExcel_Result_1 = New Dictionary(Of Tuple(Of Integer, Integer), List(Of ValueError))
        objExcel_Result_2 = New Dictionary(Of Tuple(Of Integer, Integer), List(Of ValueError))

        For w_intCol As Integer = objLocation_1.intFromCol To objLocation_1.intToCol
            For w_intRow As Integer = objLocation_1.intFromRow To objLocation_1.intToRow
                objExcel_Result_1(New Tuple(Of Integer, Integer)(w_intRow, w_intCol)) = New List(Of ValueError)
            Next
        Next

        For w_intCol As Integer = objLocation_2.intFromCol To objLocation_2.intToCol
            For w_intRow As Integer = objLocation_2.intFromRow To objLocation_2.intToRow
                objExcel_Result_2(New Tuple(Of Integer, Integer)(w_intRow, w_intCol)) = New List(Of ValueError)
            Next
        Next
    End Sub

    ''' <summary>
    ''' Set error to columns
    ''' </summary>
    ''' <param name="p_objError">Value to set in case of error</param>
    Private Sub SetError_Column(ByVal p_objError As ValueError)
        Dim w_objError As List(Of ValueError)

        If p_objError = ValueError.MissingColumn Then
            'For missing column errors, set to excel data 1
            For w_intCol_1 As Integer = objLocation_1.intFromCol To objLocation_1.intToCol
                If objEquivalentColumns.ContainsKey(w_intCol_1) Then
                Else
                    If objRemoveCol Is Nothing = False AndAlso objRemoveCol.Contains(w_intCol_1) Then
                        Continue For
                    End If

                    For w_intRow_1 As Integer = objLocation_1.intFromRow To objLocation_1.intToRow
                        w_objError = objExcel_Result_1(New Tuple(Of Integer, Integer)(w_intRow_1, w_intCol_1))
                        w_objError.Add(p_objError)
                        objExcel_Result_1(New Tuple(Of Integer, Integer)(w_intRow_1, w_intCol_1)) = w_objError
                    Next
                End If
            Next
        ElseIf p_objError = ValueError.AddedColumn Then
            'For added column errors, set to excel data 2
            For w_intCol_2 As Integer = objLocation_2.intFromCol To objLocation_2.intToCol
                If objEquivalentColumns.ContainsValue(w_intCol_2) Then
                Else
                    If objAddCol Is Nothing = False AndAlso objAddCol.Contains(w_intCol_2) Then
                        Continue For
                    End If

                    For w_intRow_2 As Integer = objLocation_2.intFromRow To objLocation_2.intToRow
                        w_objError = objExcel_Result_2(New Tuple(Of Integer, Integer)(w_intRow_2, w_intCol_2))
                        w_objError.Add(p_objError)
                        objExcel_Result_2(New Tuple(Of Integer, Integer)(w_intRow_2, w_intCol_2)) = w_objError
                    Next
                End If
            Next
        End If
    End Sub

    ''' <summary>
    ''' Set error to rows of equivalent columns
    ''' </summary>
    Private Sub SetError_Row(ByVal p_objError As ValueError)
        Dim w_objError As List(Of ValueError)

        Dim w_strValue_1 As String
        Dim w_strValue_2 As String

        If p_objError = ValueError.MissingRow Then
            'For missing row errors, set to excel data 1
            For w_intRow_1 As Integer = objLocation_1.intFromRow To objLocation_1.intToRow
                If objEquivalentRows.ContainsKey(w_intRow_1) Then
                Else
                    If objRemoveRow Is Nothing = False AndAlso objRemoveRow.Contains(w_intRow_1) Then
                        Continue For
                    End If

                    For w_intCol_1 As Integer = objLocation_1.intFromCol To objLocation_1.intToCol
                        w_objError = objExcel_Result_1(New Tuple(Of Integer, Integer)(w_intRow_1, w_intCol_1))
                        w_objError.Add(p_objError)
                        objExcel_Result_1(New Tuple(Of Integer, Integer)(w_intRow_1, w_intCol_1)) = w_objError
                    Next
                End If
            Next
        ElseIf p_objError = ValueError.AddedRow Then
            'For added row errors, set to excel data 2
            For w_intRow_2 As Integer = objLocation_2.intFromRow To objLocation_2.intToRow
                If objEquivalentRows.ContainsValue(w_intRow_2) Then
                Else
                    If objAddRow Is Nothing = False AndAlso objAddRow.Contains(w_intRow_2) Then
                        Continue For
                    End If

                    For w_intCol_2 As Integer = objLocation_2.intFromCol To objLocation_2.intToCol
                        w_objError = objExcel_Result_2(New Tuple(Of Integer, Integer)(w_intRow_2, w_intCol_2))
                        w_objError.Add(p_objError)
                        objExcel_Result_2(New Tuple(Of Integer, Integer)(w_intRow_2, w_intCol_2)) = w_objError
                    Next
                End If
            Next
        ElseIf p_objError = ValueError.DataError Then
            'For data errors, set to excel data 2
            For Each objCol As KeyValuePair(Of Integer, Integer) In objEquivalentColumns
                For Each objRow As KeyValuePair(Of Integer, Integer) In objEquivalentRows
                    w_strValue_1 = If(objExcel_Data_1(New Tuple(Of Integer, Integer)(objRow.Key, objCol.Key)).Value Is Nothing, String.Empty _
                                       , objExcel_Data_1(New Tuple(Of Integer, Integer)(objRow.Key, objCol.Key)).Value)
                    w_strValue_2 = If(objExcel_Data_2(New Tuple(Of Integer, Integer)(objRow.Value, objCol.Value)).Value Is Nothing, String.Empty _
                                       , objExcel_Data_2(New Tuple(Of Integer, Integer)(objRow.Value, objCol.Value)).Value)

                    If objChangeData Is Nothing = False AndAlso objChangeData.ContainsKey(w_strValue_1) Then
                        If w_strValue_2.Equals(objChangeData(w_strValue_1)) = False Then
                            w_objError = objExcel_Result_2(New Tuple(Of Integer, Integer)(objRow.Value, objCol.Value))
                            w_objError.Add(p_objError)
                            objExcel_Result_2(New Tuple(Of Integer, Integer)(objRow.Value, objCol.Value)) = w_objError
                        End If
                    ElseIf w_strValue_1.Equals(w_strValue_2) = False Then
                        w_objError = objExcel_Result_2(New Tuple(Of Integer, Integer)(objRow.Value, objCol.Value))
                        w_objError.Add(p_objError)
                        objExcel_Result_2(New Tuple(Of Integer, Integer)(objRow.Value, objCol.Value)) = w_objError
                    End If
                Next
            Next
        End If
    End Sub
End Class
