Public Class ProgressBar
    Public Event CancelOperation()

    Private Sub btnCancel_Click(sender As Object, e As RoutedEventArgs) Handles btnCancel.Click
        RaiseEvent CancelOperation()

        Me.Close()
    End Sub
End Class
