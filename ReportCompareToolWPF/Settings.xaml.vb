Imports Utility
Imports MahApps.Metro.Controls
Imports MahApps.Metro.Controls.Dialogs

Public Class Settings
    Inherits MetroWindow
    ''' <summary>
    ''' Set values to controls of settings
    ''' </summary>
    Private Sub Settings_OnLoad()

        txtPercent.MaxLength = 3

        barThreshold.Value = Integer.Parse(dblThreshold * 100)
        txtPercent.Text = Integer.Parse(dblThreshold * 100)

        If blnBestMatchFlg Then
            rbBestMatch.IsChecked = True
            rbImmMatch.IsChecked = False
        Else
            rbBestMatch.IsChecked = False
            rbImmMatch.IsChecked = True
        End If

        chkMerge.IsChecked = blnCompareMerge
        chkTextWrap.IsChecked = blnCompareTextWrap
        chkAlign.IsChecked = blnCompareTextAlign
        chkOrientation.IsChecked = blnCompareOrientation
        chkBorder.IsChecked = blnCompareBorder
        chkBackColor.IsChecked = blnCompareBackColor
        chkFont.IsChecked = blnCompareFont
    End Sub

    Private Sub barThreshold_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles barThreshold.ValueChanged
        txtPercent.Text = barThreshold.Value
    End Sub

    Private Sub txtPercent_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtPercent.TextChanged
        If IsValidInteger(txtPercent.Text, 0, 100) Then
            barThreshold.Value = Integer.Parse(txtPercent.Text)
        ElseIf String.IsNullOrEmpty(txtPercent.Text) Then
            barThreshold.Value = 0
        End If
    End Sub

    Private Async Sub btnApply_Click(sender As Object, e As RoutedEventArgs) Handles btnApply.Click
        If IsValidInteger(txtPercent.Text, 0, 100) = False Then
            Await Me.ShowMessageAsync("Invalid threshold value.", "Set threshold [Range: 0-100].", MessageDialogStyle.Affirmative)
            Exit Sub
        End If

        dblThreshold = Integer.Parse(txtPercent.Text) / 100

        If rbImmMatch.IsChecked Then
            blnBestMatchFlg = False
        Else
            blnBestMatchFlg = True
        End If

        blnCompareMerge = chkMerge.IsChecked
        blnCompareTextWrap = chkTextWrap.IsChecked
        blnCompareTextAlign = chkAlign.IsChecked
        blnCompareOrientation = chkOrientation.IsChecked
        blnCompareBorder = chkBorder.IsChecked
        blnCompareBackColor = chkBackColor.IsChecked
        blnCompareFont = chkFont.IsChecked

        Await Me.ShowMessageAsync("Apply", "Changes made will apply on the next compare.", MessageDialogStyle.Affirmative)

        Me.Close()
    End Sub

    Private Sub MetroWindow_Loaded(sender As Object, e As RoutedEventArgs)
        Topmost = True

        Settings_OnLoad()
    End Sub
End Class
