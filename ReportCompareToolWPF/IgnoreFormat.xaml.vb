Public Class IgnoreFormat
    Private m_objValueComments As String
    Private m_objFormatComments As New List(Of String)

    Public Property ValueComments As String
        Get
            Return m_objValueComments
        End Get
        Set(ByVal value As String)
            m_objValueComments = value
        End Set
    End Property

    Public Property FormatComments As List(Of String)
        Get
            Return m_objFormatComments
        End Get
        Set(ByVal value As List(Of String))
            m_objFormatComments = value
        End Set
    End Property

    Private Sub btnIgnoreFormat_Click(sender As Object, e As EventArgs) Handles btnIgnoreFormat.Click
        m_objValueComments = ""
        m_objFormatComments.Clear()

        If lstChkValErr.Items.Count = 1 Then
            If lstChkValErr.Items(0).IsChecked Then
                m_objValueComments = lstChkValErr.Items(0).Name
            End If
        End If

        For index = 0 To lstChkFormatErr.Items.Count - 1
            If lstChkFormatErr.Items(index).IsChecked Then
                m_objFormatComments.Add(lstChkFormatErr.Items(index).Name)
            End If
        Next

        If m_objValueComments <> "" Or m_objFormatComments.Count > 0 Then
            DialogResult = True
        End If
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        Topmost = True
    End Sub
End Class
