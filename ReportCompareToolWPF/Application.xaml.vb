Class Application

    ' Application-level events, such as Startup, Exit, and DispatcherUnhandledException
    ' can be handled in this file.

    Public Shared mutex As Threading.Mutex = Nothing

    Protected Overrides Sub OnStartup(e As StartupEventArgs)
        Dim w_blnCreatedNew As Boolean

        mutex = New Threading.Mutex(False, My.Application.Info.AssemblyName, w_blnCreatedNew)

        If Not w_blnCreatedNew Then
            Me.Shutdown()
        End If

        MyBase.OnStartup(e)

    End Sub
End Class
