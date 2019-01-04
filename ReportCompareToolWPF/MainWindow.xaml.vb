Imports System.ComponentModel
Imports System.Data
Imports System.Threading
Imports System.Windows.Forms
Imports Excel_Compare
Imports MahApps.Metro.Controls
Imports MahApps.Metro.Controls.Dialogs
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports Utility

Class MainWindow
    Inherits MetroWindow

    Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Integer, ByVal nCmdShow As Integer) As Integer
    Private Declare Function MoveWindow Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal X As Int32, ByVal Y As Int32, ByVal nWidth As Int32, ByVal nHeight As Int32, ByVal bRepaint As Boolean) As Boolean
    Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As IntPtr) As Long

    Private Const SW_MAXIMIZE As Integer = 3

    Private objExcelCompare As Compare
    Private frmSettings As Settings

    Private objFocusedCtrl As Controls.TextBox

    Private objDataTable_1 As Data.DataTable
    Private objDataTable_2 As Data.DataTable

    Private objValFilter() As String
    Private objFormatFilter() As String

    'Dim WithEvents BackgroundWorker As BackgroundWorker
    Dim m_ProgressBar As ProgressBar

    Dim WithEvents backgroundWorker As BackgroundWorker
    Dim temp As Boolean

    Public Sub New()

        InitializeComponent()

        'Set settings to its initial values
        Initialize_Settings()
        'Set import to its initial values
        Initialize_Import()
        'Set compare to its initial values
        Initialize_Compare()
        'Set export to its initial values
        InitializeExport()

    End Sub

    Private Sub ReportComparisonTool_Loaded(sender As Object, e As EventArgs) Handles Me.Loaded
        ShowMenu(Structures.Menu.Import)
    End Sub

    Private Sub HamburgerMenuControl_OnItemClick(ByVal sender As Object, ByVal e As ItemClickEventArgs)
        ShowMenu(HamburgerMenuControl.SelectedIndex)
    End Sub

    ''' <summary>
    ''' Show selected menu
    ''' </summary>
    ''' <param name="p_enmMenu"></param>
    Private Async Sub ShowMenu(ByVal p_enmMenu As Structures.Menu)

        With HamburgerMenuControl

            If [Enum].IsDefined(GetType(Structures.Menu), p_enmMenu) = False Then
                Exit Sub
            End If

            Select Case p_enmMenu

                Case Structures.Menu.Import

                    'Show import
                    Import_OnLoad()

                Case Structures.Menu.Compare

                    'Check for excel instance
                    If ExcelProcessRunning() = False Then
                        Await Me.ShowMessageAsync("Missing Excel file(s)", "Please import the Excel files to compare", MessageDialogStyle.Affirmative)
                        ShowMenu(Structures.Menu.Import)
                        Exit Sub
                    End If

                    'Reimport excel file to compare again
                    If blnCompareDone Then
                        'Compare is cancelled
                        If blnCancelFlg Then
                            Await Me.ShowMessageAsync("Operation Cancelled", "Please import the Excel files to compare", MessageDialogStyle.Affirmative)
                            ShowMenu(Structures.Menu.Import)
                            Exit Sub
                        End If

                        BringToFront(grpBoxAfterCompare)
                        Exit Select
                    End If

                    blnCancelFlg = False
                    'Show compare
                    Compare_OnLoad()

                Case Structures.Menu.Export

                    'Check for excel instance
                    If ExcelProcessRunning() = False Then
                        Await Me.ShowMessageAsync("Missing an Excel file(s)", "Please import the Excel files to compare.", MessageDialogStyle.Affirmative)
                        ShowMenu(Structures.Menu.Import)
                        Exit Sub
                    End If

                    'Compare first before export
                    If blnCompareDone = False Then
                        Await Me.ShowMessageAsync("No Comparison Yet", "Compare Excel files first before proceeding to export.", MessageDialogStyle.Affirmative)
                        ShowMenu(Structures.Menu.Compare)
                        Exit Sub
                    End If

                    'Compare is cancelled
                    If blnCancelFlg Then
                        Await Me.ShowMessageAsync("Operation Cancelled", "Please import the Excel files to compare.", MessageDialogStyle.Affirmative)
                        ShowMenu(Structures.Menu.Import)
                        Exit Sub
                    End If

                    'Show export
                    Export_OnLoad()
            End Select

            If p_enmMenu <> Structures.Menu.Setting Then
                .SelectedIndex = p_enmMenu
                .Content = HamburgerMenuControl.Items(p_enmMenu)
            End If

        End With

    End Sub

    Private Sub btnSettings_Click(sender As Object, e As RoutedEventArgs) Handles btnSettings.Click

        frmSettings = New Settings

        Me.Opacity = 0.5
        frmSettings.ShowDialog()
        Me.Opacity = 100

    End Sub

    Private Sub ReportComparisonTool_Closing(sender As Object, e As CancelEventArgs)
        CloseExcel(objExcel_1, objWorkbook_1, objWorksheet_1)
        CloseExcel(objExcel_2, objWorkbook_2, objWorksheet_2)
        CloseExcel(objExcel_3, objWorkbook_3, objWorksheet_3)
    End Sub

#Region "Settings"
    ''' <summary>
    ''' Set settings to its initial values
    ''' </summary>
    Private Sub Initialize_Settings()

        dblThreshold = 0.5

        blnBestMatchFlg = True

        blnCompareMerge = False
        blnCompareTextWrap = False
        blnCompareTextAlign = False
        blnCompareOrientation = False
        blnCompareBorder = False
        blnCompareBackColor = False
        blnCompareFont = False

    End Sub
#End Region

#Region "Import"
    ''' <summary>
    ''' Set import to its initial values
    ''' </summary>
    Private Sub Initialize_Import()

        strExcelPath_1 = String.Empty
        strExcelPath_2 = String.Empty

        blnCustFlg = False

    End Sub

    ''' <summary>
    ''' Set values to controls of import
    ''' </summary>
    Private Sub Import_OnLoad()

        'Set excel paths to control
        txtExcelFile_1.Text = strExcelPath_1
        txtExcelFile_2.Text = strExcelPath_2

        'Set type of task [Conversion/Customization]
        If blnCustFlg Then
            rbCustFlg.IsChecked = True
        Else
            rbConvFlg.IsChecked = True
        End If

    End Sub

    ''' <summary>
    ''' Browse for excel files
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Browse_Click(sender As Object, e As EventArgs) Handles btnBrowse_1.Click, btnBrowse_2.Click

        Dim dlgBrowse As New Microsoft.Win32.OpenFileDialog()

        With dlgBrowse
            .Title = "Browse for excel files"
            .Filter = "Excel Files|*.xls;*.xlsx"
            .RestoreDirectory = True
            If sender Is btnBrowse_1 AndAlso String.IsNullOrEmpty(txtExcelFile_1.Text) = False Then
                .InitialDirectory = IO.Path.GetDirectoryName(txtExcelFile_1.Text)
                .FileName = IO.Path.GetFileName(txtExcelFile_1.Text)
            ElseIf sender Is btnBrowse_2 AndAlso String.IsNullOrEmpty(txtExcelFile_2.Text) = False Then
                .InitialDirectory = IO.Path.GetDirectoryName(txtExcelFile_2.Text)
                .FileName = IO.Path.GetFileName(txtExcelFile_2.Text)
            Else
                .FileName = String.Empty
            End If

            If .ShowDialog() Then
                If sender Is btnBrowse_1 Then
                    txtExcelFile_1.Text = .FileName
                ElseIf sender Is btnBrowse_2 Then
                    txtExcelFile_2.Text = .FileName
                End If
            End If
        End With

    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click

        txtExcelFile_1.Text = String.Empty
        txtExcelFile_2.Text = String.Empty

    End Sub

    Private Async Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click

        If IsValidPath(txtExcelFile_1.Text) = False OrElse IsValidPath(txtExcelFile_2.Text) = False Then
            Await Me.ShowMessageAsync("Invalid File Path(s)", "Please input valid file paths.", MessageDialogStyle.Affirmative)
            Exit Sub
        End If

        'Check for existing instance
        If ExcelProcessRunning() = False Then
        Else
            If Await Me.ShowMessageAsync("Import New Excel", "An existing instance(s) of Excel is already running. Reimporting closes these instances. Do you still want to proceed?", MessageDialogStyle.AffirmativeAndNegative) = MessageDialogResult.Affirmative Then
                CloseExcel(objExcel_1, objWorkbook_1, objWorksheet_1)
                CloseExcel(objExcel_2, objWorkbook_2, objWorksheet_2)
                CloseExcel(objExcel_3, objWorkbook_3, objWorksheet_3)
            Else
                Exit Sub
            End If
        End If

        strExcelPath_1 = txtExcelFile_1.Text
        strExcelPath_2 = txtExcelFile_2.Text
        blnCustFlg = rbCustFlg.IsChecked

        'Open excel file
        OpenExcel(objExcel_1, objWorkbook_1, intProcID_1, strExcelPath_1, Structures.Position.Left)
        OpenExcel(objExcel_2, objWorkbook_2, intProcID_2, strExcelPath_2, Structures.Position.Right)

        Me.Topmost = True

        blnCompareDone = False

        'Proceed to compare after import
        ShowMenu(Structures.Menu.Compare)

    End Sub

    ''' <summary>
    ''' Close excel file
    ''' </summary>
    Private Sub CloseExcel(ByRef p_objExcel As Excel.Application, ByRef p_objWorkbook As Excel.Workbook, ByRef p_objWorksheet As Excel.Worksheet)

        If p_objWorksheet Is Nothing Then
        Else
            Runtime.InteropServices.Marshal.ReleaseComObject(p_objWorksheet)
            p_objWorksheet = Nothing
        End If

        If p_objWorkbook Is Nothing Then
        Else
            Try
                p_objWorkbook.Close()
            Catch ex As Exception
            Finally
                RemoveHandler p_objWorkbook.BeforeClose, AddressOf BeforeClose
                Runtime.InteropServices.Marshal.ReleaseComObject(p_objWorkbook)
                p_objWorkbook = Nothing
            End Try
        End If

        If p_objExcel Is Nothing Then
        Else
            Try
                p_objExcel.Quit()
            Catch ex As Exception
            Finally
                RemoveHandler p_objExcel.SheetBeforeRightClick, AddressOf SheetBeforeRightClick
                Runtime.InteropServices.Marshal.ReleaseComObject(p_objExcel)
                p_objExcel = Nothing
            End Try
        End If

    End Sub

    ''' <summary>
    ''' Open excel file and position to screen
    ''' </summary>
    ''' <param name="p_objExcel"></param>
    ''' <param name="p_objWorkbook"></param>
    ''' <param name="p_strPath"></param>
    ''' <param name="p_enmPos"></param>
    Private Sub OpenExcel(ByRef p_objExcel As Excel.Application, ByRef p_objWorkbook As Excel.Workbook, ByRef p_objProcID As Integer, ByVal p_strPath As String, Optional ByVal p_enmPos As Structures.Position = Structures.Position.FullScreen)

        If IsValidPath(p_strPath) = False Then
            Exit Sub
        End If

        p_objExcel = New Excel.Application

        With p_objExcel
            p_objWorkbook = .Workbooks.Add(p_strPath)
            .DisplayAlerts = False
            .DisplayFormulaBar = False
            .ActiveWindow.Zoom = 100
            .ActiveWindow.DisplayRuler = False
            ShowWindow(.Hwnd, SW_MAXIMIZE)
            If p_enmPos = Structures.Position.Left Then
                MoveWindow(.Hwnd, 0, 0, Screen.PrimaryScreen.WorkingArea.Width / 2, Screen.PrimaryScreen.WorkingArea.Height, True)
            ElseIf p_enmPos = Structures.Position.Right Then
                MoveWindow(.Hwnd, Screen.PrimaryScreen.WorkingArea.Width / 2, 0, Screen.PrimaryScreen.WorkingArea.Width / 2, Screen.PrimaryScreen.WorkingArea.Height, True)
            End If

            GetWindowThreadProcessId(.Hwnd, p_objProcID)
        End With

        AddHandler p_objExcel.SheetBeforeRightClick, AddressOf SheetBeforeRightClick
        AddHandler p_objWorkbook.BeforeClose, AddressOf BeforeClose

        p_objWorkbook.Protect()
    End Sub

    ''' <summary>
    ''' Check if excel instances are still running
    ''' </summary>
    ''' <returns></returns>
    Private Function ExcelProcessRunning() As Boolean

        If intProcID_1 = 0 OrElse intProcID_2 = 0 Then
            Return False
            Exit Function
        End If

        Return Process.GetProcesses().Any(Function(Proc_1) Proc_1.Id = intProcID_1) AndAlso Process.GetProcesses().Any(Function(Proc_2) Proc_2.Id = intProcID_2)

    End Function

    ''' <summary>
    ''' Disable right click function
    ''' </summary>
    ''' <param name="Sh"></param>
    ''' <param name="Target"></param>
    ''' <param name="Cancel"></param>
    Private Sub SheetBeforeRightClick(Sh As Object, Target As Excel.Range, ByRef Cancel As Boolean)
        Try
            Cancel = True
        Catch ex As Exception
            Throw
        Finally
        End Try
    End Sub

    ''' <summary>
    ''' Restrict closing of imported excel file
    ''' </summary>
    ''' <param name="Cancel"></param>
    Private Sub BeforeClose(ByRef Cancel As Boolean)
        Try
            Cancel = True
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "Compare"
    ''' <summary>
    ''' Set compare to its initial values
    ''' </summary>
    Private Sub Initialize_Compare()

        Try
            'No of pages to compare (From page 1 to [intNoOfPages])
            intNoOfPages = 0

            'Location of data per page
            objLocation_1 = New Dictionary(Of Integer, Location)
            objLocation_2 = New Dictionary(Of Integer, Location)

            'Specification changes
            objRemoveCol = New Dictionary(Of Integer, List(Of Integer))
            objRemoveRow = New Dictionary(Of Integer, List(Of Integer))
            objAddCol = New Dictionary(Of Integer, List(Of Integer))
            objAddRow = New Dictionary(Of Integer, List(Of Integer))

            objChangeData = New Dictionary(Of Integer, Dictionary(Of String, String))

            'Matching columns and rows
            objEquivalentColumns = New Dictionary(Of Integer, Dictionary(Of Integer, Integer))
            objEquivalentRows = New Dictionary(Of Integer, Dictionary(Of Integer, Integer))

            'Results after comparison
            objValueResult_1 = New Dictionary(Of Integer, Dictionary(Of Tuple(Of Integer, Integer), List(Of ValueError)))
            objValueResult_2 = New Dictionary(Of Integer, Dictionary(Of Tuple(Of Integer, Integer), List(Of ValueError)))
            objFormatResult = New Dictionary(Of Integer, Dictionary(Of Tuple(Of Integer, Integer), List(Of FormatError)))

        Catch ex As Exception
            Throw
        End Try

    End Sub

    Private Sub Compare_OnLoad()

        'Set sheets to combobox
        SetSheetsToComboBox(cmbSheet_1, objExcel_1)
        SetSheetsToComboBox(cmbSheet_2, objExcel_2)

        'Show specification change on customization task
        If blnCustFlg Then
            grpSpecChange.Visibility = Visibility.Visible
        Else
            grpSpecChange.Visibility = Visibility.Collapsed
        End If

        BringToFront(grpBoxBeforeCompare)

    End Sub

    ''' <summary>
    ''' Bring control to front in Compare menu
    ''' </summary>
    ''' <param name="p_objControl"></param>
    Private Sub BringToFront(ByVal p_objControl As Controls.GroupBox)

        If p_objControl Is grpBoxBeforeCompare Then
            grpBoxBeforeCompare.Visibility = Visibility.Visible
            grpBoxCompare.Visibility = Visibility.Hidden
            grpBoxAfterCompare.Visibility = Visibility.Hidden
        ElseIf p_objControl Is grpBoxAfterCompare Then
            grpBoxBeforeCompare.Visibility = Visibility.Hidden
            grpBoxCompare.Visibility = Visibility.Hidden
            grpBoxAfterCompare.Visibility = Visibility.Visible
        End If

    End Sub

    ''' <summary>
    ''' Set sheets of excel files to combobox
    ''' </summary>
    ''' <param name="p_objComboBox"></param>
    ''' <param name="p_objExcel"></param>
    Private Sub SetSheetsToComboBox(ByRef p_objComboBox As Controls.ComboBox, ByRef p_objExcel As Excel.Application)

        Try
            p_objComboBox.Items.Clear()

            For Each w_objSheet As Excel.Worksheet In p_objExcel.Worksheets
                p_objComboBox.Items.Add(w_objSheet.Name)
            Next

            p_objComboBox.SelectedIndex = 0
        Catch ex As Exception
            Throw
        Finally
        End Try
    End Sub

    Private Sub cmbSheet_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSheet_1.SelectionChanged, cmbSheet_2.SelectionChanged

        If sender Is cmbSheet_1 Then
            RemoveHandler objExcel_1.SheetSelectionChange, AddressOf SheetSelectionChange_1

            If objWorkbook_1 Is Nothing Then
                Exit Sub
            End If

            If String.IsNullOrEmpty(cmbSheet_1.SelectedItem) Then
                Exit Sub
            End If

            objWorkbook_1.Unprotect()
            'Show selected worksheet
            objWorksheet_1 = objWorkbook_1.Sheets(cmbSheet_1.SelectedItem)
            objWorksheet_1.Visible = Excel.XlSheetVisibility.xlSheetVisible
            objWorksheet_1.Activate()

            'Hide other sheets
            For Each objSheet As Excel.Worksheet In objWorkbook_1.Worksheets
                If objSheet.Name = objWorksheet_1.Name Then
                Else
                    objSheet.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
                End If
            Next

            objWorkbook_1.Protect()
            objWorksheet_1.Protect()

            'Get location of data of the selected worksheet
            GetLocation(objWorksheet_1, objLocation_1)

            AddHandler objExcel_1.SheetSelectionChange, AddressOf SheetSelectionChange_1
        ElseIf sender Is cmbSheet_2 Then
            RemoveHandler objExcel_2.SheetSelectionChange, AddressOf SheetSelectionChange_2

            If objWorkbook_2 Is Nothing Then
                Exit Sub
            End If

            If String.IsNullOrEmpty(cmbSheet_2.SelectedItem) Then
                Exit Sub
            End If

            objWorkbook_2.Unprotect()
            'Show selected worksheet
            objWorksheet_2 = objWorkbook_2.Sheets(cmbSheet_2.SelectedItem)
            objWorksheet_2.Visible = Excel.XlSheetVisibility.xlSheetVisible
            objWorksheet_2.Activate()

            'Hide other sheets
            For Each objSheet As Excel.Worksheet In objWorkbook_2.Worksheets
                If objSheet.Name = objWorksheet_2.Name Then
                Else
                    objSheet.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
                End If
            Next

            objWorkbook_2.Protect()
            objWorksheet_2.Protect()

            'Get location of data of the selected worksheet
            GetLocation(objWorksheet_2, objLocation_2)

            AddHandler objExcel_2.SheetSelectionChange, AddressOf SheetSelectionChange_2
        End If

        'Compute number of pages to compare
        intNoOfPages = If(objLocation_1.Count < objLocation_2.Count, objLocation_1.Count, objLocation_2.Count)

        If blnCustFlg Then
            'Initialize value and control of specification change
            InitializeSpecChange()
            SpecChange_OnLoad()
        End If

    End Sub

    ''' <summary>
    ''' Get location of each data of selected worksheet
    ''' </summary>
    ''' <param name="p_objWorksheet"></param>
    ''' <param name="p_objLocation"></param>
    Private Sub GetLocation(ByRef p_objWorksheet As Excel.Worksheet, ByRef p_objLocation As Dictionary(Of Integer, Location))
        Dim w_intTotalCols As Integer
        Dim w_intTotalRows As Integer
        Dim w_intPage As Integer
        Dim w_objLocation As Location
        Dim w_objHPageBreak As Excel.HPageBreak
        Dim w_objVPageBreak As Excel.VPageBreak

        Try
            If p_objWorksheet Is Nothing Then
                Exit Sub
            End If

            p_objLocation.Clear()

            With p_objWorksheet

                w_intTotalCols = If(.Cells.Find("*", Reflection.Missing.Value, Reflection.Missing.Value, Reflection.Missing.Value, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious _
                                           , False, Reflection.Missing.Value, Reflection.Missing.Value) Is Nothing, 0, .Cells.Find("*", Reflection.Missing.Value, Reflection.Missing.Value _
                                           , Reflection.Missing.Value, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious, False, Reflection.Missing.Value _
                                           , Reflection.Missing.Value).Column)
                w_intTotalRows = If(.Cells.Find("*", Reflection.Missing.Value, Reflection.Missing.Value, Reflection.Missing.Value, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious _
                                           , False, Reflection.Missing.Value, Reflection.Missing.Value) Is Nothing, 0, .Cells.Find("*", Reflection.Missing.Value, Reflection.Missing.Value _
                                           , Reflection.Missing.Value, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, False, Reflection.Missing.Value _
                                           , Reflection.Missing.Value).Row)

                If w_intTotalCols = 0 OrElse w_intTotalRows = 0 Then
                    Exit Sub
                End If

                .Cells(w_intTotalRows, w_intTotalCols).Select()

                'Get locations of all pages in the excel sheet
                For w_intVPageBreak_Idx As Integer = 0 To .VPageBreaks.Count
                    With w_objLocation
                        If w_intVPageBreak_Idx = 0 Then
                            .intFromCol = 1
                        Else
                            w_objVPageBreak = p_objWorksheet.VPageBreaks.Item(w_intVPageBreak_Idx)
                            .intFromCol = w_objVPageBreak.Location.Column
                        End If

                        If w_intVPageBreak_Idx = p_objWorksheet.VPageBreaks.Count Then
                            .intToCol = w_intTotalCols
                        Else
                            w_objVPageBreak = p_objWorksheet.VPageBreaks.Item(w_intVPageBreak_Idx + 1)
                            .intToCol = w_objVPageBreak.Location.Column - 1
                        End If
                    End With

                    For w_intHPageBreak_Idx As Integer = 0 To .HPageBreaks.Count
                        With w_objLocation
                            If w_intHPageBreak_Idx = 0 Then
                                .intFromRow = 1
                            Else
                                w_objHPageBreak = p_objWorksheet.HPageBreaks.Item(w_intHPageBreak_Idx)
                                .intFromRow = w_objHPageBreak.Location.Row
                            End If

                            If w_intHPageBreak_Idx = p_objWorksheet.HPageBreaks.Count Then
                                .intToRow = w_intTotalRows
                            Else
                                w_objHPageBreak = p_objWorksheet.HPageBreaks.Item(w_intHPageBreak_Idx + 1)
                                .intToRow = w_objHPageBreak.Location.Row - 1
                            End If
                        End With

                        w_intPage += 1
                        p_objLocation.Add(w_intPage, w_objLocation)
                    Next
                Next
            End With

        Catch ex As Exception
            Throw
        End Try

    End Sub

    ''' <summary>
    ''' Set specification change to its initial value
    ''' </summary>
    Private Sub InitializeSpecChange()

        objRemoveCol.Clear()
        objRemoveRow.Clear()
        objAddCol.Clear()
        objAddRow.Clear()

        objChangeData.Clear()

    End Sub

    ''' <summary>
    ''' Set values to controls of specification change
    ''' </summary>
    Private Sub SpecChange_OnLoad()

        rbDataLoc_Col.IsChecked = True
        rbDataLoc_Row.IsChecked = False

        txtDataLoc_Remove.Text = String.Empty
        txtDataLoc_Add.Text = String.Empty

        dgvDataLoc.ItemsSource = Nothing

        objDataTable_1 = New Data.DataTable
        With objDataTable_1
            .Columns.Add("Sheet Name", GetType(String))
            .Columns.Add("Page", GetType(Integer))
            .Columns.Add("Column or Row", GetType(String))
            .Columns.Add("Action", GetType(String))

            dgvDataLoc.ItemsSource() = .AsDataView
            dgvDataLoc.IsReadOnly = True
        End With

        txtData_From.Text = String.Empty
        txtData_To.Text = String.Empty

        dgvData.ItemsSource = Nothing

        objDataTable_2 = New Data.DataTable
        With objDataTable_2
            .Columns.Add("Page", GetType(Integer))
            .Columns.Add("From", GetType(String))
            .Columns.Add("To", GetType(String))

            dgvData.ItemsSource() = .AsDataView
            dgvData.IsReadOnly = True
        End With

    End Sub

    Private Sub DataLoc_CheckedChanges(sender As Object, e As DependencyPropertyChangedEventArgs) Handles rbDataLoc_Col.IsEnabledChanged, rbDataLoc_Row.IsEnabledChanged

        txtDataLoc_Remove.Text = String.Empty
        txtDataLoc_Add.Text = String.Empty

    End Sub

    ''' <summary>
    ''' Get selected cell in excel sheet 1
    ''' </summary>
    ''' <param name="Sh"></param>
    ''' <param name="Target"></param>
    Private Sub SheetSelectionChange_1(Sh As Object, Target As Excel.Range)

        Try
            If objFocusedCtrl Is Nothing Then
                Exit Sub
            End If

            If Sh.Name = objWorksheet_1.Name Then
                With objWorksheet_1
                    Dispatcher.Invoke(New System.Action(
                    Sub()
                        If objFocusedCtrl.Name = txtDataLoc_Remove.Name Then
                            If rbDataLoc_Col.IsChecked Then
                                txtDataLoc_Remove.Invoke(Sub()
                                                             txtDataLoc_Remove.Text = GetSelectedColumns(Target.Address(,, Excel.XlReferenceStyle.xlR1C1), objLocation_1)
                                                         End Sub)
                            ElseIf rbDataLoc_Row.IsChecked Then
                                txtDataLoc_Remove.Invoke(Sub()
                                                             txtDataLoc_Remove.Text = GetSelectedRows(Target.Address(,, Excel.XlReferenceStyle.xlR1C1), objLocation_1)
                                                         End Sub)
                            End If
                        ElseIf objFocusedCtrl.Name = txtData_From.Name Then

                            txtData_From.Invoke(Sub()
                                                    txtData_From.Text = .Cells(Target.Row, Target.Column).Value
                                                End Sub)

                        End If
                    End Sub))

                End With
            End If
        Catch ex As Exception
            Throw
        Finally
        End Try

    End Sub

    ''' <summary>
    ''' Get selected cell in excel sheet 2
    ''' </summary>
    ''' <param name="Sh"></param>
    ''' <param name="Target"></param>
    Private Sub SheetSelectionChange_2(Sh As Object, Target As Excel.Range)

        Try
            If objFocusedCtrl Is Nothing Then
                Exit Sub
            End If

            If Sh.Name = objWorksheet_2.Name Then
                With objWorksheet_2
                    Dispatcher.Invoke(New System.Action(
                        Sub()
                            If objFocusedCtrl.Name = txtDataLoc_Add.Name Then
                                If rbDataLoc_Col.IsChecked Then
                                    txtDataLoc_Add.Invoke(Sub()
                                                              txtDataLoc_Add.Text = GetSelectedColumns(Target.Address(,, Excel.XlReferenceStyle.xlR1C1), objLocation_2)
                                                          End Sub)
                                ElseIf rbDataLoc_Row.IsChecked Then
                                    txtDataLoc_Add.Invoke(Sub()
                                                              txtDataLoc_Add.Text = GetSelectedRows(Target.Address(,, Excel.XlReferenceStyle.xlR1C1), objLocation_2)
                                                          End Sub)
                                End If
                            ElseIf objFocusedCtrl.Name = txtData_To.Name Then

                                txtData_To.Invoke(Sub()
                                                      txtData_To.Text = .Cells(Target.Row, Target.Column).Value
                                                  End Sub)

                            End If
                        End Sub))
                End With
            End If
        Catch ex As Exception
            Throw
        Finally
        End Try

    End Sub

    Private Sub txtDataLoc_From_Enter(sender As Object, e As EventArgs) Handles txtDataLoc_Remove.GotFocus
        objFocusedCtrl = txtDataLoc_Remove
    End Sub

    Private Sub txtDataLoc_To_Enter(sender As Object, e As EventArgs) Handles txtDataLoc_Add.GotFocus
        objFocusedCtrl = txtDataLoc_Add
    End Sub

    Private Sub txtData_From_Enter(sender As Object, e As EventArgs) Handles txtData_From.GotFocus
        objFocusedCtrl = txtData_From
    End Sub

    Private Sub txtData_To_Enter(sender As Object, e As EventArgs) Handles txtData_To.GotFocus
        objFocusedCtrl = txtData_To
    End Sub

    ''' <summary>
    ''' Get columns selected on the worksheet
    ''' </summary>
    ''' <param name="p_strAddress"></param>
    ''' <returns></returns>
    Private Function GetSelectedColumns(ByVal p_strAddress As String, ByVal p_objLocation As Dictionary(Of Integer, Location)) As String
        Dim w_strSelections() As String = Nothing
        Dim w_strCols() As String = Nothing
        Dim w_intFCol As Integer
        Dim w_intSCol As Integer
        Dim w_intCtr As Integer
        Dim w_intCnt_1 As Integer
        Dim w_objCols As List(Of String)
        Dim w_strTemp As String
        Dim w_strCol As String = String.Empty

        Try
            If String.IsNullOrEmpty(p_strAddress) Then
                Return w_strCol
                Exit Function
            End If

            w_objCols = New List(Of String)
            w_strSelections = p_strAddress.Split(",")

            For Each str As String In w_strSelections
                w_strCols = str.Split(":")

                If w_strCols(0).IndexOf("C") = -1 OrElse String.IsNullOrEmpty(w_strCols(0).Substring(w_strCols(0).IndexOf("C") + 1, w_strCols(0).Length - w_strCols(0).IndexOf("C") - 1)) Then
                    w_intFCol = 1
                Else
                    w_intFCol = Integer.Parse(w_strCols(0).Substring(w_strCols(0).IndexOf("C") + 1, w_strCols(0).Length - w_strCols(0).IndexOf("C") - 1))
                End If

                If UBound(w_strCols) > 0 Then
                    w_intSCol = Integer.Parse(w_strCols(1).Substring(w_strCols(1).IndexOf("C") + 1, w_strCols(1).Length - w_strCols(1).IndexOf("C") - 1))
                    w_intCtr = w_intSCol - w_intFCol + 1
                Else
                    w_intCtr = 1
                End If

                For w_intCnt As Integer = 1 To w_intCtr

                    w_intCnt_1 = w_intCnt
                    If p_objLocation.Where(Function(Loc) Loc.Value.intFromCol <= w_intFCol + w_intCnt_1 - 1 AndAlso Loc.Value.intToCol >= w_intFCol + w_intCnt_1 - 1).Count = 0 Then
                        Continue For
                    End If

                    If String.IsNullOrEmpty(w_strCol) Then
                        w_strCol = ConvColNumToColName(w_intFCol + w_intCnt - 1)

                        w_objCols.Add(w_strCol)
                    Else
                        w_strTemp = ConvColNumToColName(w_intFCol + w_intCnt - 1)

                        If w_objCols.Contains(w_strTemp) = False Then
                            w_strCol &= "," & w_strTemp

                            w_objCols.Add(w_strTemp)
                        End If
                    End If
                Next

            Next

            Return w_strCol

        Catch e As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' Get rows selected on the worksheet
    ''' </summary>
    ''' <param name="p_strAddress"></param>
    ''' <returns></returns>
    Private Function GetSelectedRows(ByVal p_strAddress As String, ByVal p_objLocation As Dictionary(Of Integer, Location)) As String
        Dim w_strSelections() As String = Nothing
        Dim w_strRows() As String = Nothing
        Dim w_intFRow As Integer
        Dim w_intSRow As Integer
        Dim w_intCtr As Integer
        Dim w_intCnt_1 As Integer
        Dim w_objRows As List(Of String)
        Dim w_strTemp As String
        Dim w_strRow As String = String.Empty

        If String.IsNullOrEmpty(p_strAddress) Then
            Return w_strRow
            Exit Function
        End If

        w_objRows = New List(Of String)
        w_strSelections = p_strAddress.Split(",")

        For Each str As String In w_strSelections
            w_strRows = str.Split(":")

            If w_strRows(0).IndexOf("R") = -1 OrElse String.IsNullOrEmpty(w_strRows(0).Substring(w_strRows(0).IndexOf("R") + 1, w_strRows(0).Length - w_strRows(0).IndexOf("R") - 1)) Then
                w_intFRow = 1
            ElseIf w_strRows(0).IndexOf("C") = -1 Then
                w_intFRow = Integer.Parse(w_strRows(0).Substring(w_strRows(0).IndexOf("R") + 1, w_strRows(0).Length - w_strRows(0).IndexOf("R") - 1))
            Else
                w_intFRow = Integer.Parse(w_strRows(0).Substring(w_strRows(0).IndexOf("R") + 1, w_strRows(0).IndexOf("C") - w_strRows(0).IndexOf("R") - 1))
            End If

            If UBound(w_strRows) > 0 Then
                If w_strRows(0).IndexOf("C") = -1 Then
                    w_intSRow = Integer.Parse(w_strRows(1).Substring(w_strRows(1).IndexOf("R") + 1, w_strRows(1).Length - w_strRows(1).IndexOf("R") - 1))
                Else
                    w_intSRow = Integer.Parse(w_strRows(1).Substring(w_strRows(1).IndexOf("R") + 1, w_strRows(1).IndexOf("C") - w_strRows(1).IndexOf("R") - 1))
                End If

                w_intCtr = w_intSRow - w_intFRow + 1
            Else
                w_intCtr = 1
            End If

            For w_intCnt As Integer = 1 To w_intCtr

                w_intCnt_1 = w_intCnt
                If p_objLocation.Where(Function(Loc) Loc.Value.intFromRow <= w_intFRow + w_intCnt_1 - 1 AndAlso Loc.Value.intToRow >= w_intFRow + w_intCnt_1 - 1).Count = 0 Then
                    Continue For
                End If

                If String.IsNullOrEmpty(w_strRow) Then
                    w_strRow = w_intFRow + w_intCnt - 1

                    w_objRows.Add(w_strRow)
                Else
                    w_strTemp = w_intFRow + w_intCnt - 1

                    If w_objRows.Contains(w_strTemp) = False Then
                        w_strRow &= "," & w_strTemp

                        w_objRows.Add(w_strTemp)
                    End If
                End If
            Next

        Next

        Return w_strRow
    End Function

    Private Sub btnDataLoc_SelectAll_Click(sender As Object, e As EventArgs) Handles btnDataLoc_SelectAll.Click

        If dgvDataLoc.Items.Count = 0 Then
            Exit Sub
        End If

        If btnDataLoc_SelectAll.Content = "Select All" Then
            dgvDataLoc.SelectAll()

            btnDataLoc_SelectAll.Content = "Deselect All"
        ElseIf btnDataLoc_SelectAll.Content = "Deselect All" Then
            dgvDataLoc.UnselectAll()

            btnDataLoc_SelectAll.Content = "Select All"
        End If

    End Sub

    Private Sub dgvDataLoc_SelectedCellsChanged(sender As Object, e As SelectedCellsChangedEventArgs) Handles dgvDataLoc.SelectedCellsChanged
        With dgvDataLoc
            If .SelectedItems.Count = 0 Then
                btnDataLoc_SelectAll.Content = "Select All"
            ElseIf .SelectedItems.Count = .Items.Count Then
                btnDataLoc_SelectAll.Content = "Deselect All"
            End If
        End With
    End Sub

    Private Sub btnDataLoc_Add_Click(sender As Object, e As EventArgs) Handles btnDataLoc_Add.Click

        If rbDataLoc_Col.IsChecked Then

            'Set column specification change to datagridview
            SetColSpecChangeToDataGridview(objRemoveCol, txtDataLoc_Remove.Text, objLocation_1, objWorksheet_1.Name, "Excel File 1", "Removed Column")
            SetColSpecChangeToDataGridview(objAddCol, txtDataLoc_Add.Text, objLocation_2, objWorksheet_2.Name, "Excel File 2", "Added Column")

        ElseIf rbDataLoc_Row.IsChecked Then

            'Set row specification change to datagridview
            SetRowSpecChangeToDataGridview(objRemoveRow, txtDataLoc_Remove.Text, objLocation_1, objWorksheet_1.Name, "Excel File 1", "Removed Row")
            SetRowSpecChangeToDataGridview(objAddRow, txtDataLoc_Add.Text, objLocation_2, objWorksheet_2.Name, "Excel File 2", "Added Row")

        End If

        'Set controls to empty
        txtDataLoc_Remove.Text = String.Empty
        txtDataLoc_Add.Text = String.Empty

    End Sub

    ''' <summary>
    ''' Set added/removed column specification change to datagridview
    ''' </summary>
    ''' <param name="p_objCol"></param>
    ''' <param name="p_strCols"></param>
    ''' <param name="p_objLocation"></param>
    ''' <param name="p_strWorksheetName"></param>
    ''' <param name="p_strExcelFile"></param>
    ''' <param name="p_strMsg"></param>
    Private Sub SetColSpecChangeToDataGridview(ByRef p_objCol As Dictionary(Of Integer, List(Of Integer)), ByVal p_strCols As String _
                                               , ByVal p_objLocation As Dictionary(Of Integer, Location), ByVal p_strWorksheetName As String, ByVal p_strExcelFile As String, ByVal p_strMsg As String)
        Dim w_objCol() As String
        Dim w_lngCol As Long
        Dim w_objCols As List(Of Integer)

        If String.IsNullOrEmpty(p_strCols) = False Then
            w_objCol = p_strCols.Split(",")

            'Loop through columns selected
            For w_intCnt As Integer = 0 To UBound(w_objCol)
                w_lngCol = ConvColNameToColNum(w_objCol(w_intCnt))

                For w_intCnt_2 As Integer = 1 To intNoOfPages

                    Dim w_objDataRow As DataRow = objDataTable_1.NewRow

                    If w_lngCol >= p_objLocation(w_intCnt_2).intFromCol AndAlso w_lngCol <= p_objLocation(w_intCnt_2).intToCol Then
                        'Check if column is within a page of worksheet

                        If p_objCol.ContainsKey(w_intCnt_2) Then
                            w_objCols = p_objCol(w_intCnt_2)
                            If w_objCols.Contains(w_lngCol) = False Then
                                w_objCols.Add(w_lngCol)

                                w_objDataRow(0) = p_strWorksheetName & "[" & p_strExcelFile & "]"
                                w_objDataRow(1) = w_intCnt_2
                                w_objDataRow(2) = ConvColNumToColName(w_lngCol)
                                w_objDataRow(3) = p_strMsg

                                objDataTable_1.Rows.Add(w_objDataRow)
                            End If
                            p_objCol(w_intCnt_2) = w_objCols
                        Else
                            w_objCols = New List(Of Integer)
                            w_objCols.Add(w_lngCol)
                            p_objCol.Add(w_intCnt_2, w_objCols)

                            w_objDataRow(0) = p_strWorksheetName & "[" & p_strExcelFile & "]"
                            w_objDataRow(1) = w_intCnt_2
                            w_objDataRow(2) = ConvColNumToColName(w_lngCol)
                            w_objDataRow(3) = p_strMsg

                            objDataTable_1.Rows.Add(w_objDataRow)
                        End If
                    End If
                Next
            Next

            dgvDataLoc.ItemsSource = objDataTable_1.AsDataView

        End If

    End Sub

    ''' <summary>
    ''' Set added/removed row specification change to datagridview
    ''' </summary>
    ''' <param name="p_objRow"></param>
    ''' <param name="p_strRows"></param>
    ''' <param name="p_objLocation"></param>
    ''' <param name="p_strWorksheetName"></param>
    ''' <param name="p_strExcelFile"></param>
    ''' <param name="p_strMsg"></param>
    Private Sub SetRowSpecChangeToDataGridview(ByRef p_objRow As Dictionary(Of Integer, List(Of Integer)), ByVal p_strRows As String _
                                               , ByVal p_objLocation As Dictionary(Of Integer, Location), ByVal p_strWorksheetName As String, ByVal p_strExcelFile As String, ByVal p_strMsg As String)
        Dim w_objRow() As String
        Dim w_strRow As String
        Dim w_objRows As List(Of Integer)

        If String.IsNullOrEmpty(p_strRows) = False Then
            w_objRow = p_strRows.Split(",")

            'Loop through selected rows
            For w_intCnt As Integer = 0 To UBound(w_objRow)
                w_strRow = w_objRow(w_intCnt)

                If IsValidInteger(w_strRow) = False Then
                    Continue For
                End If

                For w_intCnt_2 As Integer = 1 To intNoOfPages
                    Dim w_objDataRow As DataRow = objDataTable_1.NewRow

                    If Integer.Parse(w_strRow) >= p_objLocation(w_intCnt_2).intFromRow AndAlso Integer.Parse(w_strRow) <= p_objLocation(w_intCnt_2).intToRow Then
                        'Check if row is with a page of the worksheet

                        If p_objRow.ContainsKey(w_intCnt_2) Then
                            w_objRows = p_objRow(w_intCnt_2)
                            If w_objRows.Contains(w_strRow) = False Then
                                w_objRows.Add(w_strRow)

                                w_objDataRow(0) = p_strWorksheetName & "[" & p_strExcelFile & "]"
                                w_objDataRow(1) = w_intCnt_2
                                w_objDataRow(2) = w_strRow
                                w_objDataRow(3) = p_strMsg

                                objDataTable_1.Rows.Add(w_objDataRow)
                            End If
                            p_objRow(w_intCnt_2) = w_objRows
                        Else
                            w_objRows = New List(Of Integer)
                            w_objRows.Add(w_strRow)
                            p_objRow.Add(w_intCnt_2, w_objRows)

                            w_objDataRow(0) = p_strWorksheetName & "[" & p_strExcelFile & "]"
                            w_objDataRow(1) = w_intCnt_2
                            w_objDataRow(2) = w_strRow
                            w_objDataRow(3) = p_strMsg

                            objDataTable_1.Rows.Add(w_objDataRow)
                        End If
                    End If
                Next
            Next

            dgvDataLoc.ItemsSource = objDataTable_1.AsDataView

        End If

    End Sub

    Private Sub btnDataLoc_Remove_Click(sender As Object, e As EventArgs) Handles btnDataLoc_Remove.Click
        Dim w_objSelectedItems() As DataRowView

        With dgvDataLoc
            ReDim w_objSelectedItems(.SelectedItems.Count - 1)
            .SelectedItems.CopyTo(w_objSelectedItems, 0)

            For Each w_objRow As DataRowView In w_objSelectedItems
                If w_objRow.Item(3).ToString = "Removed Column" Then
                    RemoveColSpecChange(objRemoveCol, Integer.Parse(w_objRow.Item(1).ToString), w_objRow.Item(2).ToString)
                ElseIf w_objRow.Item(3).ToString = "Added Column" Then
                    RemoveColSpecChange(objAddCol, Integer.Parse(w_objRow.Item(1).ToString), w_objRow.Item(2).ToString)
                End If

                If w_objRow.Item(3).ToString = "Removed Row" Then
                    RemoveRowSpecChange(objRemoveRow, Integer.Parse(w_objRow.Item(1).ToString), Integer.Parse(w_objRow.Item(2).ToString))
                ElseIf w_objRow.Item(3).ToString = "Added Row" Then
                    RemoveRowSpecChange(objAddRow, Integer.Parse(w_objRow.Item(1).ToString), Integer.Parse(w_objRow.Item(2).ToString))
                End If

                TryCast(.ItemsSource, Data.DataView).Delete(.Items.IndexOf(w_objRow))
            Next

            If .Items.Count = 0 Then
                btnDataLoc_SelectAll.Content = "Select All"
            End If
        End With

    End Sub

    ''' <summary>
    ''' Remove selected column specification change
    ''' </summary>
    ''' <param name="p_objCol"></param>
    ''' <param name="p_intPage"></param>
    ''' <param name="p_strCol"></param>
    Private Sub RemoveColSpecChange(ByRef p_objCol As Dictionary(Of Integer, List(Of Integer)), ByVal p_intPage As Integer, ByVal p_strCol As String)
        Dim w_objCols As List(Of Integer)

        If p_objCol.ContainsKey(p_intPage) Then
            w_objCols = p_objCol(p_intPage)

            If w_objCols.Contains(ConvColNameToColNum(p_strCol)) Then
                w_objCols.Remove(ConvColNameToColNum(p_strCol))
            End If

            If w_objCols.Count = 0 Then
                p_objCol.Remove(p_intPage)
            Else
                p_objCol(p_intPage) = w_objCols
            End If
        End If

    End Sub

    ''' <summary>
    ''' Remove selected row specification change
    ''' </summary>
    ''' <param name="p_objRow"></param>
    ''' <param name="p_intPage"></param>
    ''' <param name="p_intRow"></param>
    Private Sub RemoveRowSpecChange(ByRef p_objRow As Dictionary(Of Integer, List(Of Integer)), ByVal p_intPage As Integer, ByVal p_intRow As Integer)
        Dim w_objRows As List(Of Integer)

        If p_objRow.ContainsKey(p_intPage) Then
            w_objRows = p_objRow(p_intPage)

            If w_objRows.Contains(p_intRow) Then
                w_objRows.Remove(p_intRow)
            End If

            If w_objRows.Count = 0 Then
                p_objRow.Remove(p_intPage)
            Else
                p_objRow(p_intPage) = w_objRows
            End If
        End If

    End Sub

    Private Sub btnData_Add_Click(sender As Object, e As EventArgs) Handles btnData_Add.Click
        Dim w_objData As Dictionary(Of String, String)

        If String.IsNullOrEmpty(txtData_From.Text) AndAlso String.IsNullOrEmpty(txtData_To.Text) Then
            Exit Sub
        End If

        If String.IsNullOrEmpty(txtData_From.Text) = False OrElse String.IsNullOrEmpty(txtData_To.Text) = False Then
            If String.IsNullOrEmpty(txtData_From.Text) OrElse String.IsNullOrEmpty(txtData_To.Text) Then
                MsgBox("Insufficient data entered.")
                Exit Sub
            End If
        End If

        For w_intCnt_2 As Integer = 1 To intNoOfPages
            Dim w_objDataRow As DataRow = objDataTable_2.NewRow

            If objChangeData.ContainsKey(w_intCnt_2) Then
                w_objData = objChangeData(w_intCnt_2)

                If w_objData.ContainsKey(Trim(txtData_From.Text)) Then
                Else
                    w_objData.Add(Trim(txtData_From.Text), Trim(txtData_To.Text))
                    objChangeData(w_intCnt_2) = w_objData

                    w_objDataRow(0) = w_intCnt_2
                    w_objDataRow(1) = Trim(txtData_From.Text)
                    w_objDataRow(2) = Trim(txtData_To.Text)

                    objDataTable_2.Rows.Add(w_objDataRow)
                End If
            Else
                w_objData = New Dictionary(Of String, String)
                w_objData.Add(Trim(txtData_From.Text), Trim(txtData_To.Text))
                objChangeData.Add(w_intCnt_2, w_objData)

                w_objDataRow(0) = w_intCnt_2
                w_objDataRow(1) = Trim(txtData_From.Text)
                w_objDataRow(2) = Trim(txtData_To.Text)

                objDataTable_2.Rows.Add(w_objDataRow)
            End If
        Next

        dgvData.ItemsSource = objDataTable_2.AsDataView

        txtData_From.Text = String.Empty
        txtData_To.Text = String.Empty

    End Sub

    Private Sub btnData_SelectAll_Click(sender As Object, e As EventArgs) Handles btnData_SelectAll.Click

        If dgvData.Items.Count = 0 Then
            Exit Sub
        End If

        If btnData_SelectAll.Content = "Select All" Then
            dgvData.SelectAll()

            btnData_SelectAll.Content = "Deselect All"
        ElseIf btnData_SelectAll.Content = "Deselect All" Then
            dgvData.UnselectAll()

            btnData_SelectAll.Content = "Select All"
        End If

    End Sub

    Private Sub dgvData_SelectedCellsChanged(sender As Object, e As SelectedCellsChangedEventArgs) Handles dgvData.SelectedCellsChanged
        With dgvData
            If .SelectedItems.Count = 0 Then
                btnData_SelectAll.Content = "Select All"
            ElseIf .SelectedItems.Count = .Items.Count Then
                btnData_SelectAll.Content = "Deselect All"
            End If
        End With
    End Sub

    Private Sub btnData_Remove_Click(sender As Object, e As EventArgs) Handles btnData_Remove.Click
        Dim w_objData As Dictionary(Of String, String)
        Dim w_objSelectedItems() As DataRowView

        With dgvData
            ReDim w_objSelectedItems(.SelectedItems.Count - 1)
            .SelectedItems.CopyTo(w_objSelectedItems, 0)

            For Each w_objRow As DataRowView In w_objSelectedItems
                If objChangeData.ContainsKey(w_objRow.Item(0).ToString) Then
                    w_objData = objChangeData(w_objRow.Item(0).ToString)

                    If w_objData.ContainsKey(w_objRow.Item(1).ToString) Then
                        w_objData.Remove(w_objRow.Item(1).ToString)
                    End If

                    If w_objData.Count = 0 Then
                        objChangeData.Remove(w_objRow.Item(0).ToString)
                    Else
                        objChangeData(w_objRow.Item(0).ToString) = w_objData
                    End If

                End If

                TryCast(.ItemsSource, Data.DataView).Delete(.Items.IndexOf(w_objRow))
            Next

            If .Items.Count = 0 Then
                btnDataLoc_SelectAll.Content = "Select All"
            End If
        End With
    End Sub

    Private Sub btnCompare_Click(sender As Object, e As EventArgs) Handles btnCompare.Click

        RemoveHandler objExcel_1.SheetSelectionChange, AddressOf SheetSelectionChange_1
        RemoveHandler objExcel_2.SheetSelectionChange, AddressOf SheetSelectionChange_2

        Try
            If objWorksheet_1.Name = "" OrElse objWorksheet_2.Name = "" Then
            End If
        Catch ex As Exception
            System.Windows.MessageBox.Show(Application.Current.MainWindow, "Please import the Excel files to compare" _
                                    , "Missing Excel Files", MessageBoxButton.OK)
            ShowMenu(Structures.Menu.Import)
            Exit Sub
        End Try

        m_ProgressBar = New ProgressBar
        AddHandler m_ProgressBar.CancelOperation, AddressOf CancelOperation
        m_ProgressBar.Show()

        Me.Opacity = 0.5
        Me.backgroundWorker = New BackgroundWorker
        Me.backgroundWorker.WorkerReportsProgress = True
        Me.backgroundWorker.WorkerSupportsCancellation = True
        AddHandler Me.backgroundWorker.DoWork, AddressOf worker_DoWork
        AddHandler Me.backgroundWorker.ProgressChanged, AddressOf worker_ProgressChanged
        AddHandler Me.backgroundWorker.RunWorkerCompleted, AddressOf worker_RunWorkerCompleted
        Me.backgroundWorker.RunWorkerAsync()
        TaskbarItemInfo.ProgressState = Shell.TaskbarItemProgressState.Normal

    End Sub

    Private Sub CancelOperation()
        m_ProgressBar.txtBlockMainProgress.Dispatcher.BeginInvoke(Sub()
                                                                      m_ProgressBar.txtBlockMainProgress.Text = "Cancelling operation ..."
                                                                  End Sub)
        Me.backgroundWorker.CancelAsync()

        blnCancelFlg = True

        Me.Opacity = 100
    End Sub

    Private Sub worker_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs)

        blnCompareDone = True

        objExcelCompare = New Compare

        With objExcelCompare
            .SetThreshold = dblThreshold
            .CompareToBestMatchData = blnBestMatchFlg

            .CompareMerge = blnCompareMerge
            .CompareTextWrap = blnCompareTextWrap
            .CompareTextAlign = blnCompareTextAlign
            .CompareOrientation = blnCompareOrientation
            .CompareBorder = blnCompareBorder
            .CompareBackColor = blnCompareBackColor
            .CompareFont = blnCompareFont

            .NoOfPages = intNoOfPages
            .Page_Location_1 = objLocation_1
            .Page_Location_2 = objLocation_2
            .RemovedColumn = objRemoveCol
            .RemovedRow = objRemoveRow
            .AddedColumn = objAddCol
            .AddedRow = objAddRow
            .DataChange = objChangeData

            .Compare(objWorksheet_1, objWorksheet_2, Me.backgroundWorker, e)

            objEquivalentColumns = .EquivalentColumns
            objEquivalentRows = .EquivalentRows
            objValueResult_1 = .ValueResult_1
            objValueResult_2 = .ValueResult_2
            objFormatResult = .FormatResult
        End With

    End Sub

    Private Sub worker_ProgressChanged(sender As Object, e As ProgressChangedEventArgs)

        m_ProgressBar.txtBlockMainProgress.Dispatcher.BeginInvoke(Sub()
                                                                      m_ProgressBar.txtBlockMainProgress.Text = e.UserState
                                                                  End Sub)

        TaskbarItemInfo.ProgressValue = e.ProgressPercentage / 100

    End Sub

    Private Sub worker_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs)
        If e.Error Is Nothing Then
            If e.Cancelled Then
                ShowMenu(Structures.Menu.Import)
            Else
                TaskbarItemInfo.ProgressValue = 100
                TaskbarItemInfo.ProgressState = 0
                m_ProgressBar.Close()
                InitializeIgnoreErrors()
                IgnoreErrors_OnLoad()
            End If

            Me.Opacity = 100
        Else
            MsgBox(e.Error.Message)
        End If

    End Sub

    ''' <summary>
    ''' Set ignore to its initial values
    ''' </summary>
    Private Sub InitializeIgnoreErrors()

        Dim w_lstValueError As List(Of ValueErrorClass) = New List(Of ValueErrorClass)
        Dim w_lstFormatError As List(Of FormatErrorClass) = New List(Of FormatErrorClass)

        For Each w_Error As ValueError In System.Enum.GetValues(GetType(ValueError))
            w_lstValueError.Add(New ValueErrorClass With {.Id = CType(w_Error, Integer), .Name = "Show " & w_Error.ToString, .IsChecked = False})
        Next
        For Each w_Error As FormatError In System.Enum.GetValues(GetType(FormatError))
            w_lstFormatError.Add(New FormatErrorClass With {.Id = CType(w_Error, Integer), .Name = "Show cells with " & w_Error.ToString, .IsChecked = False})
        Next

        lstChkValueError.ItemsSource = w_lstValueError
        lstChkFormatError.ItemsSource = w_lstFormatError

        ReDim objValFilter(lstChkValueError.Items.Count - 1)
        ReDim objFormatFilter(lstChkFormatError.Items.Count - 1)

        objDifferences = New Dictionary(Of Integer, Dictionary(Of Tuple(Of String, String), Tuple(Of String, String)))

    End Sub

    ''' <summary>
    ''' Set controls of ignore errors to initial values
    ''' </summary>
    Private Sub IgnoreErrors_OnLoad()

        SetPageToComboBox()

        PopulateDataGridView(dgvExcel_1)
        PopulateDataGridView(dgvExcel_2)

        cmbShowPage.SelectedIndex = 0
        btnFilter.[RaiseEvent](New RoutedEventArgs(Controls.Button.ClickEvent))

        BringToFront(grpBoxAfterCompare)

    End Sub

    ''' <summary>
    ''' Set pages to combobox
    ''' </summary>
    Private Sub SetPageToComboBox()

        'Set pages to combobox
        cmbShowPage.Items.Clear()
        cmbShowPage.Items.Add("All pages")

        For w_intCnt As Integer = 1 To intNoOfPages
            cmbShowPage.Items.Add(w_intCnt.ToString)
        Next

    End Sub

    ''' <summary>
    ''' Set errors to datagridview
    ''' </summary>
    ''' <param name="p_objDataGridView"></param>
    Private Sub PopulateDataGridView(ByRef p_objDataGridView As Controls.DataGrid)
        Dim w_objDataTable As New Data.DataTable
        Dim w_objValueResult As Dictionary(Of Integer, Dictionary(Of Tuple(Of Integer, Integer), List(Of ValueError)))
        Dim w_objCollectedLines As New List(Of Tuple(Of Integer, Integer)) 'Contains all the collected rows/columns for skipping iteration
        Dim w_strFormatError As String = ""

        p_objDataGridView.IsReadOnly = True
        p_objDataGridView.ItemsSource = Nothing

        If p_objDataGridView Is dgvExcel_1 Then
            w_objValueResult = objValueResult_1
        Else
            w_objValueResult = objValueResult_2
        End If

        '========= Add Columns =========
        w_objDataTable.Columns.Add("Address", GetType(String))
        w_objDataTable.Columns.Add("VE", GetType(Integer))
        w_objDataTable.Columns.Add("Format Error", GetType(String))
        w_objDataTable.Columns.Add("Start Range", GetType(Tuple(Of Integer, Integer)))
        w_objDataTable.Columns.Add("End Range", GetType(Tuple(Of Integer, Integer)))
        w_objDataTable.Columns.Add("Page", GetType(Integer))
        'Temporary
        w_objDataTable.Columns.Add("Value Error", GetType(String))
        '===============================

        For Each w_objPageContent In w_objValueResult 'For each pages in the result
            Dim w_intPage As Integer = w_objPageContent.Key 'Page number

            For Each w_objLineContent In w_objPageContent.Value 'For each cells in the page
                Dim w_intError As ValueError = If(w_objLineContent.Value.Count > 0, w_objLineContent.Value(0), -1)
                Dim w_objDataRow As DataRow = w_objDataTable.NewRow

                If w_intError <> -1 And Not w_objCollectedLines.Contains(w_objLineContent.Key) Then 'Check for VEs and collected rows/columns for skipping iteration
                    If w_intError = ValueError.AddedColumn Or w_intError = ValueError.MissingColumn Then 'Error on column
                        Dim w_objDetectedColumn = w_objLineContent.Key.Item2
                        Dim w_objDetectedColumnCells = w_objPageContent.Value.Keys.ToList.FindAll(Function(p) p.Item2 = w_objDetectedColumn)

                        'Add cells to skip iteration
                        For Each w_objDetectedCell In w_objDetectedColumnCells
                            w_objCollectedLines.Add(w_objDetectedCell)
                        Next

                        w_objDataRow(0) = ConvColNumToColName(w_objDetectedColumn)
                        w_objDataRow(1) = w_intError
                        w_objDataRow(3) = w_objDetectedColumnCells.Item(0)
                        w_objDataRow(4) = w_objDetectedColumnCells.Item(w_objDetectedColumnCells.Count - 1)
                        w_objDataRow(5) = w_intPage
                        w_objDataRow(6) = [Enum].GetName(GetType(ValueError), w_intError)
                    ElseIf w_intError = ValueError.AddedRow Or w_intError = ValueError.MissingRow Then 'Error on row
                        Dim w_objDetectedRow = w_objLineContent.Key.Item1
                        Dim w_objDetectedRowCells = w_objPageContent.Value.Keys.ToList.FindAll(Function(p) p.Item1 = w_objDetectedRow)

                        'Add cells to skip iteration
                        For Each w_objDetectedCell In w_objDetectedRowCells
                            w_objCollectedLines.Add(w_objDetectedCell)
                        Next

                        w_objDataRow(0) = w_objDetectedRow
                        w_objDataRow(1) = w_intError
                        w_objDataRow(3) = w_objDetectedRowCells.Item(0)
                        w_objDataRow(4) = w_objDetectedRowCells.Item(w_objDetectedRowCells.Count - 1)
                        w_objDataRow(5) = w_intPage
                        w_objDataRow(6) = [Enum].GetName(GetType(ValueError), w_intError)
                    Else 'Error on cell
                        w_objDataRow(0) = ConvColNumToColName(w_objLineContent.Key.Item2) & Convert.ToString(w_objLineContent.Key.Item1)
                        w_objDataRow(1) = w_intError
                        w_objDataRow(3) = New Tuple(Of Integer, Integer)(w_objLineContent.Key.Item1, w_objLineContent.Key.Item2)
                        w_objDataRow(4) = New Tuple(Of Integer, Integer)(w_objLineContent.Key.Item1, w_objLineContent.Key.Item2)
                        w_objDataRow(5) = w_intPage
                        w_objDataRow(6) = [Enum].GetName(GetType(ValueError), w_intError)

                        If p_objDataGridView Is dgvExcel_2 And objFormatResult.Count <> 0 Then
                            If objFormatResult(w_intPage).ContainsKey(New Tuple(Of Integer, Integer)(w_objLineContent.Key.Item1, w_objLineContent.Key.Item2)) Then
                                If objFormatResult(w_intPage)(New Tuple(Of Integer, Integer)(w_objLineContent.Key.Item1, w_objLineContent.Key.Item2)).Count <> 0 Then
                                    w_objDataRow(2) = String.Join(",", objFormatResult(w_intPage)(New Tuple(Of Integer, Integer)(w_objLineContent.Key.Item1 _
                                                                                                                                 , w_objLineContent.Key.Item2)).ToArray)
                                End If
                            End If
                        End If
                    End If
                Else
                    If p_objDataGridView Is dgvExcel_2 And objFormatResult.Count <> 0 Then
                        If objFormatResult(w_intPage).ContainsKey(New Tuple(Of Integer, Integer)(w_objLineContent.Key.Item1, w_objLineContent.Key.Item2)) Then
                            If objFormatResult(w_intPage)(New Tuple(Of Integer, Integer)(w_objLineContent.Key.Item1, w_objLineContent.Key.Item2)).Count <> 0 Then
                                w_objDataRow(0) = ConvColNumToColName(w_objLineContent.Key.Item2) & Convert.ToString(w_objLineContent.Key.Item1)
                                w_objDataRow(2) = String.Join(",", objFormatResult(w_intPage)(New Tuple(Of Integer, Integer)(w_objLineContent.Key.Item1, w_objLineContent.Key.Item2)).ToArray)
                                w_objDataRow(3) = New Tuple(Of Integer, Integer)(w_objLineContent.Key.Item1, w_objLineContent.Key.Item2)
                                w_objDataRow(4) = New Tuple(Of Integer, Integer)(w_objLineContent.Key.Item1, w_objLineContent.Key.Item2)
                                w_objDataRow(5) = w_intPage
                            End If
                        End If
                    End If
                End If
                If w_objDataRow(0).ToString <> "" Then
                    w_objDataTable.Rows.Add(w_objDataRow)
                End If
            Next
        Next

        p_objDataGridView.ItemsSource = w_objDataTable.AsDataView
        p_objDataGridView.Columns(1).Visibility = Visibility.Hidden
        p_objDataGridView.Columns(3).Visibility = Visibility.Hidden
        p_objDataGridView.Columns(4).Visibility = Visibility.Hidden
        p_objDataGridView.VerticalScrollBarVisibility = ScrollBarVisibility.Visible

        If p_objDataGridView Is dgvExcel_1 Then
            p_objDataGridView.Columns(2).Visibility = Visibility.Hidden
        End If

    End Sub

    Private Sub Remove_Filter()

        For w_intCnt As Integer = 0 To lstChkValueError.Items.Count - 1
            If lstChkValueError.Items(w_intCnt).IsChecked Then
                lstChkValueError.Items(w_intCnt).IsChecked = False
            End If
        Next

        For w_intCnt As Integer = 0 To lstChkFormatError.Items.Count - 1
            If lstChkFormatError.Items(w_intCnt).IsChecked Then
                lstChkFormatError.Items(w_intCnt).IsChecked = False
            End If
        Next

    End Sub

    Private Sub btnFilter_Click(sender As Object, e As EventArgs) Handles btnFilter.Click

        For intIdx As Integer = 0 To lstChkValueError.Items.Count - 1
            If lstChkValueError.Items(intIdx).IsChecked Then
                objValFilter(intIdx) = intIdx.ToString
            Else
                objValFilter(intIdx) = Nothing
            End If
        Next

        For intIdx As Integer = 0 To lstChkFormatError.Items.Count - 1
            If lstChkFormatError.Items(intIdx).IsChecked Then
                objFormatFilter(intIdx) = intIdx.ToString
            Else
                objFormatFilter(intIdx) = Nothing
            End If
        Next

        FilterDataGridview(dgvExcel_1, objValFilter, objFormatFilter, cmbShowPage.SelectedIndex)
        FilterDataGridview(dgvExcel_2, objValFilter, objFormatFilter, cmbShowPage.SelectedIndex)
        dgvExcel_1.CurrentCell = Nothing
        dgvExcel_2.CurrentCell = Nothing

    End Sub

    ''' <summary>
    ''' Filter errors
    ''' </summary>
    ''' <param name="p_objDataGridView"></param>
    ''' <param name="p_blnValueFilter"></param>
    ''' <param name="p_blnFormatFilter"></param>
    ''' <param name="p_intPage"></param>
    Private Sub FilterDataGridview(ByRef p_objDataGridView As Controls.DataGrid, Optional ByRef p_blnValueFilter() As String = Nothing _
            , Optional ByRef p_blnFormatFilter() As String = Nothing, Optional ByVal p_intPage As Integer = Nothing)

        Dim w_strBuilder As New System.Text.StringBuilder
        Dim w_strValueJoin As String = ""
        Dim w_strFormatJoin As String = ""
        Dim w_objDataTable As Data.DataView = p_objDataGridView.ItemsSource

        'Filter by page
        If p_intPage <> Nothing Then
            w_strBuilder.Append("[Page] = " & p_intPage.ToString)
        End If

        'Filter by errors
        If p_blnValueFilter IsNot Nothing Then
            w_strValueJoin = String.Join(",", p_blnValueFilter.Where(Function(s) Not String.IsNullOrEmpty(s)))
        End If
        If w_strValueJoin <> "" Then
            If w_strBuilder.Length <> 0 Then
                w_strBuilder.Append(" And ")
            End If
            w_strBuilder.Append("[VE] IN (" & w_strValueJoin & ")")
        End If

        'Filter by format
        If p_blnFormatFilter IsNot Nothing Then
            Dim w_objFormatArray() As String

            w_objFormatArray = p_blnFormatFilter.Where(Function(s) Not String.IsNullOrEmpty(s)).ToArray

            If w_objFormatArray.Count <> 0 Then
                For index = 0 To w_objFormatArray.Count - 1
                    w_objFormatArray(index) = [Enum].GetName(GetType(FormatError), Convert.ToUInt32(w_objFormatArray(index))).ToString
                Next

                w_strFormatJoin = String.Join(",", w_objFormatArray)

                If w_strBuilder.Length <> 0 Then
                    If w_strValueJoin = "" Then
                        w_strBuilder.Append(" And ")
                    Else
                        w_strBuilder.Append(" Or ")
                    End If
                End If

                For Each w_strFormat In w_objFormatArray
                    w_strBuilder.Append("[Format Error] Like '%" & w_strFormat & "%'")
                    If Not w_strFormat.Equals(w_objFormatArray.Last) Then
                        w_strBuilder.Append(" OR ")
                    End If
                Next
            End If
        End If

        w_objDataTable.RowFilter = w_strBuilder.ToString
        w_strBuilder.Clear()

    End Sub

    Private Sub btnIgnore_1_Click(sender As Object, e As RoutedEventArgs) Handles btnIgnore_1.Click
        If dgvExcel_1.SelectedCells.Count > 0 Then
            ClearDetection(dgvExcel_1, objWorksheet_1)
        End If
    End Sub

    Private Sub btnIgnore_2_Click(sender As Object, e As RoutedEventArgs) Handles btnIgnore_2.Click
        If dgvExcel_2.SelectedCells.Count > 0 Then
            ClearDetection(dgvExcel_2, objWorksheet_2)
        End If
    End Sub

    Private Sub dgvExcel_1_MouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs) Handles dgvExcel_1.MouseLeftButtonUp
        If sender.SelectedIndex <> -1 Then
            Dim w_objSelection() As Excel.Range = GetCellRange(dgvExcel_1)

            Try
                FocusCell(w_objSelection, objWorksheet_1)
            Catch ex As Exception
                Throw
            Finally
            End Try
        End If
    End Sub

    Private Sub dgvExcel_2_MouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs) Handles dgvExcel_2.MouseLeftButtonUp
        If sender.SelectedIndex <> -1 Then
            Dim w_objSelection() As Excel.Range = GetCellRange(dgvExcel_2)

            Try
                FocusCell(w_objSelection, objWorksheet_2)
            Catch ex As Exception
                Throw
            Finally
            End Try
        End If
    End Sub

    ''' <summary>
    ''' Get cell range if selected cell in the datagridview is a column or row
    ''' </summary>
    ''' <param name="p_objDataGridView"></param>
    ''' <returns></returns>
    Private Function GetCellRange(ByRef p_objDataGridView As Controls.DataGrid) As Excel.Range()
        Dim w_objWorksheet As Excel.Worksheet = Nothing
        Dim w_objSelection(1) As Excel.Range
        Dim w_objStartRange As Tuple(Of Integer, Integer)
        Dim w_objEndRange As Tuple(Of Integer, Integer)


        Try
            If p_objDataGridView Is dgvExcel_1 Then
                w_objWorksheet = objWorksheet_1
            Else
                w_objWorksheet = objWorksheet_2
            End If

            w_objStartRange = p_objDataGridView.SelectedItem.Row(3)
            w_objEndRange = p_objDataGridView.SelectedItem.Row(4)
            w_objSelection(0) = w_objWorksheet.Cells(w_objStartRange.Item1, w_objStartRange.Item2)
            w_objSelection(1) = w_objWorksheet.Cells(w_objEndRange.Item1, w_objEndRange.Item2)

            Return w_objSelection
        Catch ex As Exception
            Throw
        Finally
        End Try
    End Function

    Private Sub FocusCell(ByVal p_objCellArray() As Object, ByRef p_objWorkSheet As Excel.Worksheet)
        Dim w_objRange1 As Excel.Range
        Dim w_objRange2 As Excel.Range

        w_objRange1 = p_objCellArray(0)
        w_objRange2 = p_objCellArray(1)

        If p_objWorkSheet.Range(w_objRange1, w_objRange2).MergeCells Is DBNull.Value Then
            Dim w_objTempRange As Range
            If w_objRange1.Column = w_objRange2.Column And w_objRange1.Row <> w_objRange2.Row Then 'Column
                For index = w_objRange1.Row + 1 To w_objRange2.Row
                    w_objTempRange = p_objWorkSheet.Cells(index, w_objRange1.Column)
                    If p_objWorkSheet.Range(w_objTempRange, w_objTempRange).MergeArea.Column <> w_objRange1.Column Then
                        If w_objTempRange.Row > (p_objWorkSheet.Range(w_objRange1, w_objRange2).Rows.Count * 0.9) Then
                            w_objRange2 = p_objWorkSheet.Cells(index - 1, w_objRange1.Column)
                            Exit For
                        End If

                        w_objRange1 = p_objWorkSheet.Cells(index + 1, w_objRange1.Column)
                    End If
                Next
            ElseIf w_objRange1.Row = w_objRange2.Row And w_objRange1.Column <> w_objRange2.Column Then 'Row
                For index = w_objRange1.Column + 1 To w_objRange2.Column
                    w_objTempRange = p_objWorkSheet.Cells(w_objRange1.Row, index)
                    If p_objWorkSheet.Range(w_objRange1, w_objRange2).MergeArea.Row <> w_objRange1.Row Then

                        If w_objTempRange.Column > (p_objWorkSheet.Range(w_objRange1, w_objRange2).Columns.Count * 0.9) Then
                            w_objRange2 = p_objWorkSheet.Cells(w_objRange1.Row, index - 1)
                        End If

                        w_objRange1 = p_objWorkSheet.Cells(w_objRange1.Row, index + 1)
                    End If
                Next
            End If
        End If

        p_objWorkSheet.Range(w_objRange1, w_objRange2).Select()

        SetForegroundWindow(p_objWorkSheet.Application.Hwnd)

    End Sub

    ''' <summary>
    ''' Clear excel comments
    ''' </summary>
    ''' <param name="p_objDataGridView"></param>
    ''' <param name="p_objWorkSheet"></param>
    Private Sub ClearDetection(ByRef p_objDataGridView As Controls.DataGrid, ByRef p_objWorkSheet As Excel.Worksheet)
        Dim w_objStartRange As Tuple(Of Integer, Integer) = Nothing
        Dim w_objEndRange As Tuple(Of Integer, Integer) = Nothing
        Dim w_objRange1 As Excel.Range = Nothing
        Dim w_objRange2 As Excel.Range = Nothing
        Dim w_objSelectedCellsCount As Integer
        Dim w_objValueResult As Dictionary(Of Integer, Dictionary(Of Tuple(Of Integer, Integer), List(Of ValueError)))

        Try
            w_objSelectedCellsCount = p_objDataGridView.SelectedItems.Count - 1

            p_objWorkSheet.Unprotect()

            If p_objWorkSheet Is objExcel_1.ActiveSheet Then
                w_objValueResult = objValueResult_1
            Else
                w_objValueResult = objValueResult_2
            End If

            If w_objSelectedCellsCount = 0 Then
                w_objStartRange = p_objDataGridView.SelectedItem.Row(3)
                w_objEndRange = p_objDataGridView.SelectedItem.Row(4)
                w_objRange1 = p_objWorkSheet.Cells(w_objStartRange.Item1, w_objStartRange.Item2)
                w_objRange2 = p_objWorkSheet.Cells(w_objEndRange.Item1, w_objEndRange.Item2)

                If w_objRange1.Comment IsNot Nothing Then
                    Dim w_strValueComment As String
                    Dim w_strFormatComment As String
                    Dim w_objFormatCommentList As New List(Of String)

                    w_strValueComment = p_objDataGridView.SelectedItem.Row(6).ToString
                    w_strFormatComment = p_objDataGridView.SelectedItem.Row(2).ToString

                    If w_strFormatComment <> "" Then
                        w_objFormatCommentList.AddRange(Split(w_strFormatComment, ",")) 'If there are multiple format comments
                    End If

                    If w_objFormatCommentList.Count > 1 Or (w_strValueComment <> "" And w_objFormatCommentList.Count > 0) Then 'If there are multiple format comments
                        Dim w_blnResult As DialogResult
                        Dim w_objIgnoreFormat As IgnoreFormat

                        w_objIgnoreFormat = New IgnoreFormat

                        With w_objIgnoreFormat
                            .lstChkValErr.Items.Clear()
                            .lstChkFormatErr.Items.Clear()

                            Dim w_lstValueError As List(Of ValueErrorClass) = New List(Of ValueErrorClass)

                            If w_strValueComment <> "" Then
                                w_lstValueError.Add(New ValueErrorClass With {.Name = w_strValueComment, .IsChecked = False})
                            End If

                            Dim w_lstFormatError As List(Of FormatErrorClass) = New List(Of FormatErrorClass)

                            For Each w_strFormat In w_objFormatCommentList
                                w_lstFormatError.Add(New FormatErrorClass With {.Name = w_strFormat, .IsChecked = False})
                            Next

                            .lstChkValErr.ItemsSource = w_lstValueError
                            .lstChkFormatErr.ItemsSource = w_lstFormatError

                            w_blnResult = .ShowDialog()

                            If w_blnResult Then

                                If System.Windows.MessageBox.Show(Application.Current.MainWindow, "Ignoring error(s) cannot be reverted. Do you still want to continue?" _
                                    , "Ignore error(s)", MessageBoxButton.YesNo) = MessageBoxResult.No Then
                                    Exit Sub
                                End If

                                Dim w_strNewExcelComment As String
                                w_objRange1.Comment.Delete()

                                If .ValueComments <> "" Then 'VE of data error is selected to be deleted
                                    w_strValueComment = ""
                                    'p_objDataGridView.SelectedItem.Row(1) = DBNull.Value
                                    p_objDataGridView.SelectedItem.Row(6) = ""
                                    p_objWorkSheet.Cells(w_objStartRange.Item1, w_objStartRange.Item2).Interior.ColorIndex = 0
                                End If

                                For Each w_strFormat In .FormatComments
                                    w_objFormatCommentList.Remove(w_strFormat)
                                Next

                                If w_objFormatCommentList.Count = 0 And w_strValueComment = "" Then
                                    TryCast(p_objDataGridView.ItemsSource, Data.DataView).Delete(p_objDataGridView.SelectedIndex)
                                Else
                                    w_strNewExcelComment = String.Join(",", w_objFormatCommentList.ToArray)
                                    p_objDataGridView.SelectedItem.Row(2) = w_strNewExcelComment 'Add new format comment
                                    w_objFormatCommentList.Add(w_strValueComment) 'Add VE to the format comment list
                                    w_strNewExcelComment = String.Join(",", w_objFormatCommentList.ToArray.Where(Function(s) Not String.IsNullOrEmpty(s)))
                                    w_objRange1.AddComment(w_strNewExcelComment) 'Add comment to excel
                                End If
                            Else
                                Exit Sub 'Exit when form is cancelled
                            End If
                        End With

                    ElseIf w_objFormatCommentList.Count = 1 Then 'single format error

                        If System.Windows.MessageBox.Show(Application.Current.MainWindow, "Ignoring error(s) cannot be reverted. Do you still want to continue?" _
                                    , "Ignore error(s)", MessageBoxButton.YesNo) = MessageBoxResult.No Then
                            Exit Sub
                        End If

                        w_objRange1.Comment.Delete()
                        TryCast(p_objDataGridView.ItemsSource, Data.DataView).Delete(p_objDataGridView.SelectedIndex)

                    Else

                        If System.Windows.MessageBox.Show(Application.Current.MainWindow, "Ignoring error(s) cannot be reverted. Do you still want to continue?" _
                                    , "Ignore error(s)", MessageBoxButton.YesNo) = MessageBoxResult.No Then
                            Exit Sub
                        End If

                        'Column and row range error
                        Dim w_objCellValueComments As List(Of ValueError) = w_objValueResult(p_objDataGridView.SelectedItem.Row(5))(New Tuple(Of Integer, Integer) _
                                (p_objDataGridView.SelectedItem.Row(3).Item1, p_objDataGridView.SelectedItem.Row(3).Item2))
                        Dim w_objNewValueComment As New List(Of String)
                        Dim w_objCellValueErrorList As List(Of ValueError)

                        For Each w_objCellValueComment In w_objCellValueComments
                            w_objNewValueComment.Add([Enum].GetName(GetType(ValueError), w_objCellValueComment))
                        Next

                        w_objNewValueComment.Remove(w_strValueComment)
                        w_objRange1.Comment.Delete()

                        If w_strValueComment.Contains("Column") Then
                            For index = w_objStartRange.Item1 To w_objEndRange.Item1
                                w_objCellValueErrorList = w_objValueResult(p_objDataGridView.SelectedItem.Row(5))(New Tuple(Of Integer, Integer) _
                            (index, w_objStartRange.Item2))
                                If w_objCellValueErrorList.Count > 1 Then 'Cell has column and range error
                                    If w_strValueComment.Contains("Missing") Then
                                        p_objWorkSheet.Cells(index, w_objStartRange.Item2).Interior.Color = System.Drawing.Color.LightGreen
                                        w_objCellValueErrorList.Remove(ValueError.MissingColumn)
                                    Else
                                        p_objWorkSheet.Cells(index, w_objStartRange.Item2).Interior.Color = System.Drawing.Color.LightGreen
                                        w_objCellValueErrorList.Remove(ValueError.AddedColumn)
                                    End If
                                Else
                                    w_objCellValueErrorList.Clear()
                                    p_objWorkSheet.Cells(index, w_objStartRange.Item2).Interior.ColorIndex = 0
                                End If
                            Next

                            If w_objCellValueComments.Count = 1 And w_objStartRange.Item2 = 1 Then
                                w_objRange1.AddComment(String.Join(",", w_objNewValueComment.ToArray))
                            End If
                        Else
                            For index = w_objStartRange.Item2 To w_objEndRange.Item2
                                w_objCellValueErrorList = w_objValueResult(p_objDataGridView.SelectedItem.Row(5))(New Tuple(Of Integer, Integer) _
                            (w_objStartRange.Item1, index))
                                If w_objCellValueErrorList.Count > 1 Then 'Cell has column and range error
                                    If w_strValueComment.Contains("Missing") Then
                                        p_objWorkSheet.Cells(w_objStartRange.Item1, index).Interior.Color = System.Drawing.Color.LightYellow
                                        w_objCellValueErrorList.Remove(ValueError.MissingRow)
                                    Else
                                        p_objWorkSheet.Cells(w_objStartRange.Item1, index).Interior.Color = System.Drawing.Color.LightYellow
                                        w_objCellValueErrorList.Remove(ValueError.AddedRow)
                                    End If
                                Else
                                    w_objCellValueErrorList.Clear()
                                    p_objWorkSheet.Cells(w_objStartRange.Item1, index).Interior.ColorIndex = 0
                                End If
                            Next

                            If w_objCellValueComments.Count = 1 And w_objStartRange.Item1 = 1 Then
                                w_objRange1.AddComment(String.Join(",", w_objNewValueComment.ToArray))
                            End If
                        End If

                        TryCast(p_objDataGridView.ItemsSource, Data.DataView).Delete(p_objDataGridView.SelectedIndex)

                    End If
                End If
            End If

        Catch ex As Exception
            Throw
        Finally
            p_objWorkSheet.Protect()
            Runtime.InteropServices.Marshal.ReleaseComObject(w_objRange1)
            Runtime.InteropServices.Marshal.ReleaseComObject(w_objRange2)
        End Try

    End Sub

    Private Sub btnGoto_Export_Click(sender As Object, e As EventArgs) Handles btnGoto_Export.Click
        ShowMenu(Structures.Menu.Export)
    End Sub

    ''' <summary>
    ''' Retrieve differences found
    ''' </summary>
    Private Sub GetDifferences()
        Dim w_objDiff As Dictionary(Of Tuple(Of String, String), Tuple(Of String, String))
        Dim w_intPage As Integer
        Dim w_strAddress As String
        Dim w_objStartRange As Tuple(Of Integer, Integer)
        Dim w_objEndRange As Tuple(Of Integer, Integer)
        Dim w_strEqCell As String
        Dim w_strValueErr As String
        Dim w_strFormatErr As String

        For Each w_objRow As DataRowView In dgvExcel_1.Items
            w_intPage = w_objRow.Item("Page")

            If objDifferences.ContainsKey(w_intPage) Then
                w_objDiff = objDifferences(w_intPage)
            Else
                w_objDiff = New Dictionary(Of Tuple(Of String, String), Tuple(Of String, String))
                objDifferences.Add(w_intPage, w_objDiff)
            End If

            w_strAddress = w_objRow.Item("Address")
            w_strEqCell = String.Empty
            If w_objDiff.ContainsKey(New Tuple(Of String, String)(w_strAddress, w_strEqCell)) Then
            Else
                w_strValueErr = If(IsDBNull(w_objRow.Item("Value Error")), String.Empty, w_objRow.Item("Value Error"))
                If w_strValueErr.Equals(ValueError.MissingColumn.ToString) Then
                    w_objDiff.Add(New Tuple(Of String, String)(w_strAddress, w_strEqCell), New Tuple(Of String, String)("Column", "Missing"))
                ElseIf w_strValueErr.Equals(ValueError.MissingRow.ToString) Then
                    w_objDiff.Add(New Tuple(Of String, String)(w_strAddress, w_strEqCell), New Tuple(Of String, String)("Row", "Missing"))
                ElseIf w_strValueErr.Equals(ValueError.AddedColumn.ToString) Then
                    w_objDiff.Add(New Tuple(Of String, String)(w_strAddress, w_strEqCell), New Tuple(Of String, String)("Column", "Added"))
                ElseIf w_strValueErr.Equals(ValueError.AddedRow.ToString) Then
                    w_objDiff.Add(New Tuple(Of String, String)(w_strAddress, w_strEqCell), New Tuple(Of String, String)("Row", "Added"))
                End If
            End If

            objDifferences(w_intPage) = w_objDiff
        Next

        For Each w_objRow As DataRowView In dgvExcel_2.Items
            w_intPage = w_objRow.Item("Page")

            If objDifferences.ContainsKey(w_intPage) Then
                w_objDiff = objDifferences(w_intPage)
            Else
                w_objDiff = New Dictionary(Of Tuple(Of String, String), Tuple(Of String, String))
                objDifferences.Add(w_intPage, w_objDiff)
            End If

            w_strAddress = w_objRow.Item("Address")
            w_objStartRange = w_objRow.Item("Start Range")
            w_objEndRange = w_objRow.Item("End Range")

            w_strEqCell = String.Empty
            If w_objStartRange.Equals(w_objEndRange) Then
                w_strEqCell = FindEquivalentCell(w_intPage, w_objStartRange)
            End If

            If w_objDiff.ContainsKey(New Tuple(Of String, String)(w_strEqCell, w_strAddress)) Then
            Else
                w_strValueErr = If(IsDBNull(w_objRow.Item("Value Error")), String.Empty, w_objRow.Item("Value Error"))
                If w_strValueErr.Equals(ValueError.MissingColumn.ToString) Then
                    w_objDiff.Add(New Tuple(Of String, String)(w_strEqCell, w_strAddress), New Tuple(Of String, String)("Column", "Missing"))
                ElseIf w_strValueErr.Equals(ValueError.MissingRow.ToString) Then
                    w_objDiff.Add(New Tuple(Of String, String)(w_strEqCell, w_strAddress), New Tuple(Of String, String)("Row", "Missing"))
                ElseIf w_strValueErr.Equals(ValueError.AddedColumn.ToString) Then
                    w_objDiff.Add(New Tuple(Of String, String)(w_strEqCell, w_strAddress), New Tuple(Of String, String)("Column", "Added"))
                ElseIf w_strValueErr.Equals(ValueError.AddedRow.ToString) Then
                    w_objDiff.Add(New Tuple(Of String, String)(w_strEqCell, w_strAddress), New Tuple(Of String, String)("Row", "Added"))
                Else
                    w_strFormatErr = If(IsDBNull(w_objRow.Item("Format Error")), String.Empty, w_objRow.Item("Format Error"))

                    If String.IsNullOrEmpty(w_strValueErr) Then
                        w_objDiff.Add(New Tuple(Of String, String)(w_strEqCell, w_strAddress), New Tuple(Of String, String)("Cell", w_strFormatErr))
                    ElseIf String.IsNullOrEmpty(w_strFormatErr) Then
                        w_objDiff.Add(New Tuple(Of String, String)(w_strEqCell, w_strAddress), New Tuple(Of String, String)("Cell", w_strValueErr))
                    Else
                        w_objDiff.Add(New Tuple(Of String, String)(w_strEqCell, w_strAddress), New Tuple(Of String, String)("Cell", w_strValueErr & ":" & w_strFormatErr))
                    End If
                End If
            End If

            objDifferences(w_intPage) = w_objDiff
        Next

    End Sub

    ''' <summary>
    ''' Find matching data of Excel File 2 from Excel File 1
    ''' </summary>
    ''' <param name="p_intPage"></param>
    ''' <param name="p_objRowCol"></param>
    Private Function FindEquivalentCell(ByVal p_intPage As Integer, ByVal p_objRowCol As Tuple(Of Integer, Integer)) As String
        Dim w_objMatch As Dictionary(Of Integer, Integer)
        Dim w_strEqCol As String = String.Empty
        Dim w_strEqRow As String = String.Empty

        If objEquivalentColumns.ContainsKey(p_intPage) Then
            w_objMatch = objEquivalentColumns(p_intPage)

            If w_objMatch.ContainsValue(p_objRowCol.Item2) Then
                For Each objCol As KeyValuePair(Of Integer, Integer) In w_objMatch
                    If objCol.Value = p_objRowCol.Item2 Then
                        w_strEqCol = ConvColNumToColName(objCol.Key)
                        Exit For
                    End If
                Next
            End If
        End If

        If objEquivalentRows.ContainsKey(p_intPage) Then
            w_objMatch = objEquivalentRows(p_intPage)

            If w_objMatch.ContainsValue(p_objRowCol.Item1) Then
                For Each objRow As KeyValuePair(Of Integer, Integer) In w_objMatch
                    If objRow.Value = p_objRowCol.Item1 Then
                        w_strEqRow = objRow.Key
                        Exit For
                    End If
                Next
            End If
        End If

        Return w_strEqCol & w_strEqRow

    End Function
#End Region

#Region "Export"
    ''' <summary>
    ''' Set values of export to initial value
    ''' </summary>
    Private Sub InitializeExport()

        strExportExcel_1 = String.Empty
        strExportExcel_2 = String.Empty
        strExportReport = String.Empty

    End Sub

    ''' <summary>
    ''' Set controls of export to initial values
    ''' </summary>
    Private Sub Export_OnLoad()

        chkResult_1.IsChecked = True
        chkResult_2.IsChecked = True
        chkGenReport.IsChecked = True

        txtResult_1.Text = strExportExcel_1
        txtResult_2.Text = strExportExcel_2
        txtGenReport.Text = strExportReport

        objDifferences.Clear()

        cmbShowPage.SelectedIndex = 0
        Remove_Filter()
        btnFilter.[RaiseEvent](New RoutedEventArgs(Controls.Button.ClickEvent))

        GetDifferences()

    End Sub

    Private Sub Result_CheckedChanged(sender As Object, e As DependencyPropertyChangedEventArgs) Handles chkResult_1.IsEnabledChanged, chkResult_2.IsEnabledChanged, chkGenReport.IsEnabledChanged

        If sender.Checked Then
            If sender Is chkResult_1 Then
                btnBrowse_Result1.IsEnabled = True
            ElseIf sender Is chkResult_2 Then
                btnBrowse_Result2.IsEnabled = True
            ElseIf sender Is chkGenReport Then
                btnBrowse_GenReport.IsEnabled = True
            End If
        Else
            If sender Is chkResult_1 Then
                btnBrowse_Result1.IsEnabled = False
            ElseIf sender Is chkResult_2 Then
                btnBrowse_Result2.IsEnabled = False
            ElseIf sender Is chkGenReport Then
                btnBrowse_GenReport.IsEnabled = False
            End If
        End If

    End Sub

    Private Sub Browse_Result_Click(sender As Object, e As EventArgs) Handles btnBrowse_Result1.Click, btnBrowse_Result2.Click, btnBrowse_GenReport.Click
        Dim dlgBrowseSave As New Microsoft.Win32.SaveFileDialog()

        With dlgBrowseSave
            .Title = "Save excel files"
            .Filter = "Excel Files|*.xls;*.xlsx"
            .RestoreDirectory = True
            .DefaultExt = "xlsx"
            .FilterIndex = 1
            .FileName = String.Empty

            If .ShowDialog() Then
                If sender Is btnBrowse_Result1 Then
                    txtResult_1.Text = .FileName
                ElseIf sender Is btnBrowse_Result2 Then
                    txtResult_2.Text = .FileName
                ElseIf sender Is btnBrowse_GenReport Then
                    txtGenReport.Text = .FileName
                End If
            End If
        End With

    End Sub

    Private Sub btnClear_Export_Click(sender As Object, e As EventArgs) Handles btnClear_Export.Click

        txtResult_1.Text = String.Empty
        txtResult_2.Text = String.Empty
        txtGenReport.Text = String.Empty

    End Sub

    Private Async Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click

        If ExcelProcessRunning() = False Then
            Await Me.ShowMessageAsync("Workbook(s) to save not found.", "An excel instance(s) was stopped.", MessageDialogStyle.Affirmative)
            Exit Sub
        End If

        If objWorkbook_1 Is Nothing OrElse objWorkbook_2 Is Nothing Then
            Await Me.ShowMessageAsync("Workbook(s) to save not found.", "An excel instance(s) was stopped.", MessageDialogStyle.Affirmative)
            Exit Sub
        End If

        If ValidFiles() = False Then
            Await Me.ShowMessageAsync("Invalid file path(s) is/are found. Saving failed ...", "Input valid path(s) to save comparison results", MessageDialogStyle.Affirmative)
            Exit Sub
        End If

        If chkResult_1.IsChecked Then
            objWorkbook_1.SaveAs(txtResult_1.Text, Excel.XlFileFormat.xlOpenXMLWorkbook)
        End If

        If chkResult_2.IsChecked Then
            objWorkbook_2.SaveAs(txtResult_2.Text, Excel.XlFileFormat.xlOpenXMLWorkbook)
        End If

        If chkGenReport.IsChecked Then
            Try
                GenerateReport()
                objWorkbook_3.SaveAs(txtGenReport.Text, Excel.XlFileFormat.xlOpenXMLWorkbook)
                objWorkbook_3.Worksheets("Sheet1").Name = "Comparison_Result"
            Catch ex As Exception
                System.Windows.MessageBox.Show(Application.Current.MainWindow, "Error on accessing the file path. Check the location if you are replacing an existing file." _
                                    , "Access Error", MessageBoxButton.OK)
                Exit Sub
            End Try
        End If

    End Sub

    ''' <summary>
    ''' Determine if files to be saved are valid
    ''' </summary>
    ''' <returns></returns>
    Private Function ValidFiles() As Boolean
        Dim w_blnResult As Boolean

        w_blnResult = True

        If chkResult_1.IsChecked Then
            If IsValidDirectory(txtResult_1.Text) AndAlso IsValidFileName(txtResult_1.Text) Then
            Else
                w_blnResult = False

            End If
        End If
        If chkResult_2.IsChecked Then
            If IsValidDirectory(txtResult_2.Text) AndAlso IsValidFileName(txtResult_2.Text) Then
            Else
                w_blnResult = False

            End If
        End If
        If chkGenReport.IsChecked Then
            If IsValidDirectory(txtGenReport.Text) AndAlso IsValidFileName(txtGenReport.Text) Then
            Else
                w_blnResult = False

            End If
        End If

        Return w_blnResult

    End Function

    ''' <summary>
    ''' Create comparison report
    ''' </summary>
    Private Sub GenerateReport()
        Dim w_intSheetRow As Integer
        Dim w_intStartDataRow As Integer
        Dim w_intEndDataRow As Integer
        Dim w_intRow As Integer

        objExcel_3 = New Excel.Application
        objExcel_3.DisplayAlerts = False
        objExcel_3.Visible = True
        objWorkbook_3 = objExcel_3.Workbooks.Add()
        objWorksheet_3 = objWorkbook_3.ActiveSheet

        With objWorksheet_3
            For w_intPage As Integer = 1 To intNoOfPages

                w_intSheetRow += 1

                InsertHeader(objWorksheet_3, w_intSheetRow, w_intPage)

                w_intSheetRow += 1
                w_intStartDataRow = w_intSheetRow

                w_intRow = InsertColsCompared(objWorksheet_3, w_intStartDataRow, w_intPage)
                w_intEndDataRow = If(w_intEndDataRow > w_intRow, w_intEndDataRow, w_intRow)

                w_intRow = InsertRowsCompared(objWorksheet_3, w_intStartDataRow, w_intPage)
                w_intEndDataRow = If(w_intEndDataRow > w_intRow, w_intEndDataRow, w_intRow)

                w_intRow = InsertDifferences(objWorksheet_3, w_intStartDataRow, w_intPage)
                w_intEndDataRow = If(w_intEndDataRow > w_intRow, w_intEndDataRow, w_intRow)

                w_intSheetRow = w_intEndDataRow
            Next
        End With
    End Sub

    ''' <summary>
    ''' Create header for report per page
    ''' </summary>
    ''' <param name="p_objWorkSheet"></param>
    ''' <param name="p_intRow"></param>
    ''' <param name="p_intPage"></param>
    Private Sub InsertHeader(ByRef p_objWorkSheet As Excel.Worksheet, ByRef p_intRow As Integer, ByVal p_intPage As Integer)

        With p_objWorkSheet
            'Header
            .Range("A" & p_intRow).Value = "Page: " & p_intPage
            .Range("A" & p_intRow).Font.Bold = True
            .Range("A" & p_intRow).EntireRow.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow)

            p_intRow += 1
            .Range("A" & p_intRow).Value = "Columns Compared"
            .Range("A" & p_intRow, "C" & p_intRow).Font.Bold = True
            .Range("A" & p_intRow, "C" & p_intRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen)

            .Range("D" & p_intRow).Value = "Rows Compared"
            .Range("D" & p_intRow, "F" & p_intRow).Font.Bold = True
            .Range("D" & p_intRow, "F" & p_intRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue)

            .Range("G" & p_intRow).Value = "Differences Found"
            .Range("G" & p_intRow, "K" & p_intRow).Font.Bold = True
            .Range("G" & p_intRow, "K" & p_intRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightPink)
        End With

    End Sub

    ''' <summary>
    ''' Set columns compared
    ''' </summary>
    ''' <param name="p_objWorksheet"></param>
    ''' <param name="p_intRow"></param>
    ''' <param name="p_intPage"></param>
    ''' <returns></returns>
    Private Function InsertColsCompared(ByRef p_objWorksheet As Excel.Worksheet, ByVal p_intRow As Integer, ByVal p_intPage As Integer) As Integer
        Dim w_intStartRow As Integer
        Dim w_objMatch As Dictionary(Of Integer, Integer)

        w_intStartRow = p_intRow

        With p_objWorksheet
            .Range("B" & p_intRow).Value = "Excel File 1"
            .Range("B" & p_intRow).Font.Bold = True
            .Range("B" & p_intRow).Columns.AutoFit()

            .Range("C" & p_intRow).Value = "Excel File 2"
            .Range("C" & p_intRow).Font.Bold = True
            .Range("C" & p_intRow).Columns.AutoFit()

            If objEquivalentColumns.ContainsKey(p_intPage) Then
                w_objMatch = objEquivalentColumns(p_intPage)

                For Each w_objCols As KeyValuePair(Of Integer, Integer) In w_objMatch
                    p_intRow += 1
                    .Range("B" & p_intRow).Value = ConvColNumToColName(w_objCols.Key)
                    .Range("B" & p_intRow).NumberFormat = "@"
                    .Range("C" & p_intRow).Value = ConvColNumToColName(w_objCols.Value)
                    .Range("C" & p_intRow).NumberFormat = "@"
                Next
            End If

            .ListObjects.AddEx(, .Range("B" & w_intStartRow, "C" & p_intRow), , Excel.XlYesNoGuess.xlYes,, "TableStyleLight1")

        End With

        Return p_intRow
    End Function

    ''' <summary>
    ''' Set rows compared
    ''' </summary>
    ''' <param name="p_objWorksheet"></param>
    ''' <param name="p_intRow"></param>
    ''' <param name="p_intPage"></param>
    ''' <returns></returns>
    Private Function InsertRowsCompared(ByRef p_objWorksheet As Excel.Worksheet, ByVal p_intRow As Integer, ByVal p_intPage As Integer) As Integer
        Dim w_intStartRow As Integer
        Dim w_objMatch As Dictionary(Of Integer, Integer)

        w_intStartRow = p_intRow

        With p_objWorksheet
            .Range("E" & p_intRow).Value = "Excel File 1"
            .Range("E" & p_intRow).Font.Bold = True
            .Range("E" & p_intRow).Columns.AutoFit()

            .Range("F" & p_intRow).Value = "Excel File 2"
            .Range("F" & p_intRow).Font.Bold = True
            .Range("F" & p_intRow).Columns.AutoFit()

            If objEquivalentRows.ContainsKey(p_intPage) Then
                w_objMatch = objEquivalentRows(p_intPage)

                For Each w_objRows As KeyValuePair(Of Integer, Integer) In w_objMatch
                    p_intRow += 1
                    .Range("E" & p_intRow).Value = w_objRows.Key
                    .Range("E" & p_intRow).NumberFormat = "@"
                    .Range("F" & p_intRow).Value = w_objRows.Value
                    .Range("F" & p_intRow).NumberFormat = "@"
                Next
            End If

            .ListObjects.AddEx(, .Range("E" & w_intStartRow, "F" & p_intRow), , Excel.XlYesNoGuess.xlYes,, "TableStyleLight1")

        End With

        Return p_intRow
    End Function

    ''' <summary>
    ''' Set differences found between excel files
    ''' </summary>
    ''' <param name="p_objWorksheet"></param>
    ''' <param name="p_intRow"></param>
    ''' <param name="p_intPage"></param>
    ''' <returns></returns>
    Private Function InsertDifferences(ByRef p_objWorksheet As Excel.Worksheet, ByVal p_intRow As Integer, ByVal p_intPage As Integer) As Integer
        Dim w_intStartRow As Integer
        Dim w_objDiff As Dictionary(Of Tuple(Of String, String), Tuple(Of String, String))

        w_intStartRow = p_intRow

        With p_objWorksheet
            .Range("H" & p_intRow).Value = "Type Of Data"
            .Range("H" & p_intRow).Font.Bold = True
            .Range("H" & p_intRow).Columns.AutoFit()

            .Range("I" & p_intRow).Value = "Error"
            .Range("I" & p_intRow).Font.Bold = True
            .Range("I" & p_intRow).Columns.AutoFit()

            .Range("J" & p_intRow).Value = "Excel File 1"
            .Range("J" & p_intRow).Font.Bold = True
            .Range("J" & p_intRow).Columns.AutoFit()

            .Range("K" & p_intRow).Value = "Excel File 2"
            .Range("K" & p_intRow).Font.Bold = True
            .Range("K" & p_intRow).Columns.AutoFit()

            If objDifferences.ContainsKey(p_intPage) Then
                w_objDiff = objDifferences(p_intPage)

                For Each w_objData As KeyValuePair(Of Tuple(Of String, String), Tuple(Of String, String)) In w_objDiff
                    p_intRow += 1
                    .Range("H" & p_intRow).Value = w_objData.Value.Item1
                    .Range("H" & p_intRow).NumberFormat = "@"
                    .Range("I" & p_intRow).Value = w_objData.Value.Item2
                    .Range("I" & p_intRow).NumberFormat = "@"
                    .Range("J" & p_intRow).Value = w_objData.Key.Item1
                    .Range("J" & p_intRow).NumberFormat = "@"
                    .Range("K" & p_intRow).Value = w_objData.Key.Item2
                    .Range("K" & p_intRow).NumberFormat = "@"
                Next
            End If

            .ListObjects.AddEx(, .Range("H" & w_intStartRow, "K" & p_intRow), , Excel.XlYesNoGuess.xlYes,, "TableStyleLight1")

        End With

        Return p_intRow
    End Function
#End Region

End Class

Public Class ValueErrorClass

    Public Property Id As Integer
    Public Property Name As String
    Public Property IsChecked As Boolean

End Class

Public Class FormatErrorClass
    Public Property Id As Integer
    Public Property Name As String
    Public Property IsChecked As Boolean
End Class

