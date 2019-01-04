Imports Microsoft.Office.Interop

Public Module Structures

    Public Enum Menu
        Setting = -1
        Import = 0
        Compare = 1
        Export = 2
    End Enum

    Public Enum Position
        FullScreen
        Left
        Right
    End Enum

    Public Structure Location
        Dim intFromRow As Integer
        Dim intToRow As Integer
        Dim intFromCol As Integer
        Dim intToCol As Integer
    End Structure

    Public Enum ValueError
        DataError
        AddedColumn
        AddedRow
        MissingColumn
        MissingRow
    End Enum

    Public Structure Format
        Dim MergeCells As Boolean
        Dim HorizontalAlignment As Excel.XlHAlign
        Dim VerticalAlignment As Excel.XlVAlign
        Dim Font As Font
        Dim Border_DiagDown As Border
        Dim Border_DiagUp As Border
        Dim Border_EdgeTop As Border
        Dim Border_EdgeBottom As Border
        Dim Border_EdgeLeft As Border
        Dim Border_EdgeRight As Border
        Dim Border_InsideHl As Border
        Dim Border_InsideVl As Border
        Dim BackColor As Double
        Dim WrapText As Boolean
        Dim ShrinkToFit As Boolean
        Dim Orientation As Integer
    End Structure

    Public Structure Font
        Dim Name As String
        Dim Style As String
        Dim Color As Double
        Dim Size As Double
        Dim Underline As Integer
        Dim Strikethrough As Boolean
        Dim Subscript As Boolean
        Dim Superscript As Boolean
    End Structure

    Public Structure Border
        Dim LineStyle As Integer
        Dim Weight As Integer
        Dim Color As Double
    End Structure

    Public Enum FormatError
        MergingError
        TextWrapError
        TextAlignmentError
        OrientationError
        FontError
        BorderError
        BackColorError
    End Enum

End Module
