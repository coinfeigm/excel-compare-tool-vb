Public Module Utility
    ''' <summary>
    ''' Determine if number is a valid integer
    ''' Optional: Validate if number is within the specified range
    ''' </summary>
    ''' <param name="p_strInt">String to validate</param>
    ''' <param name="p_intMin">Minimum valid integer</param>
    ''' <param name="p_intMax">Maximum valid integer</param>
    ''' <returns></returns>
    Public Function IsValidInteger(ByVal p_strInt As String, Optional ByVal p_intMin As Integer = 0, Optional ByVal p_intMax As Integer = 0) As Boolean
        Dim w_Int As Integer

        Return Integer.TryParse(p_strInt, w_Int) AndAlso If(p_intMin <> 0 OrElse p_intMax <> 0, w_Int >= p_intMin AndAlso w_Int <= p_intMax, True)
    End Function

    ''' <summary>
    ''' Determine if path is valid
    ''' </summary>
    ''' <param name="p_strPath"></param>
    ''' <returns></returns>
    Public Function IsValidPath(ByVal p_strPath As String) As Boolean
        Try
            Return IO.File.Exists(p_strPath)
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Determine if directory is valid
    ''' </summary>
    ''' <param name="p_strDirectory"></param>
    ''' <returns></returns>
    Public Function IsValidDirectory(ByVal p_strDirectory As String) As Boolean
        Try
            If String.IsNullOrEmpty(p_strDirectory) Then
                Return False
            Else
                Return IO.Directory.Exists(IO.Path.GetDirectoryName(p_strDirectory))
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Determine if filename is valid
    ''' </summary>
    ''' <param name="p_strPath"></param>
    ''' <returns></returns>
    Public Function IsValidFileName(ByVal p_strPath As String) As Boolean
        Dim w_strFilename As String

        Try
            w_strFilename = IO.Path.GetFileName(p_strPath)
            Return Not String.IsNullOrEmpty(w_strFilename)
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Determine if input only contains letters
    ''' </summary>
    ''' <param name="p_String"></param>
    ''' <returns></returns>
    Public Function IsAlpha(ByVal p_String As String) As Boolean
        If p_String.Count <> p_String.Count(Function(C) Char.IsLetter(C)) Then
            Return False
        Else
            'Only contains letters.
            Return True
        End If
    End Function

    ''' <summary>
    ''' Return column number of given column name in excel
    ''' </summary>
    ''' <param name="p_strColName"></param>
    ''' <returns></returns>
    Public Function ConvColNameToColNum(ByVal p_strColName As String) As Long
        Dim w_strColName As String
        Dim w_lngCol As Long

        If String.IsNullOrEmpty(p_strColName) OrElse IsAlpha(p_strColName) = False Then
            Return 0
            Exit Function
        End If

        w_strColName = p_strColName.ToUpperInvariant

        For w_intCnt As Integer = 0 To w_strColName.Length - 1
            w_lngCol *= 26
            w_lngCol += (Asc(w_strColName.Substring(w_intCnt, 1)) - Asc("A") + 1)
        Next

        Return w_lngCol
    End Function

    ''' <summary>
    ''' Return column name of given column number in excel
    ''' </summary>
    ''' <param name="p_ColNum"></param>
    ''' <returns></returns>
    Public Function ConvColNumToColName(ByVal p_ColNum As Integer) As String
        Dim w_strColName As String = String.Empty
        Dim w_intCnt As Integer

        Try
            If IsValidInteger(p_ColNum) = False Then
                Return String.Empty
                Exit Function
            End If

            While p_ColNum > 0
                w_intCnt = (p_ColNum - 1) Mod 26
                w_strColName = Convert.ToChar(65 + w_intCnt).ToString() & w_strColName
                p_ColNum = (p_ColNum - w_intCnt) / 26
            End While

            Return w_strColName
        Catch ex As Exception
            Throw
        End Try
    End Function
End Module
