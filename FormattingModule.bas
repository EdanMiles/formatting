Attribute VB_Name = "FormattingModule"
Option Explicit

Const vFTIBlue As Long = 6502144
Const vMediumBlue As Long = 11626240  ' used to be 12684585
Const vLightGrey As Long = 11776687   'used to be 13488082

Const sAccounting As String = "#,##0_);(#,##0);-_%"
Const sAccounting0 As String = "#,##0.0_);(#,##0.0);-_%"
Const sAccounting00 As String = "#,##0.00_);(#,##0.00);-_%"
Const sAccounting000 As String = "#,##0.000_);(#,##0.000);-_%"

Const sPercentPoints As String = "+0%;-0%;-_%"
Const sPercentPoints0 As String = "+0.0%;-0.0%;-_%"
Const sPercentPoints00 As String = "+0.00%;-0.00%;-_%"
Const sPercentPoints000 As String = "+0.000%;-0.000%;-_%"

Const sPercentage As String = "0%_);(0%); -_%"
Const sPercentage0 As String = "0.0%_);(0.0%); -_%"
Const sPercentage00 As String = "0.00%_);(0.00%); -_%"
Const sPercentage000 As String = "0.000%_);(0.000%); -_%"

Const sMultiple As String = "0×"
Const sMultiple0 As String = "0.0×"
Const sMultiple00 As String = "0.00×"
Const sMultiple000 As String = "0.000×"


Sub Formatting()
Attribute Formatting.VB_ProcData.VB_Invoke_Func = "F\n14"
' Keyboard Shortcut: Ctrl+Shift+F
    With FormattingForm
        .Top = Application.Height / 2 - (FormattingForm.Height / 2)
        .Left = Application.Width / 2 - (FormattingForm.Width / 2)
        .Show
    End With
End Sub

Sub Section()
    Dim sReturnLoc As String
    sReturnLoc = Selection.Address

    With ActiveSheet.Rows(Selection.row())
        .RowHeight = 10
        .Interior.ColorIndex = 0
        .Insert
    End With
    With ActiveSheet.Rows(Selection.row())
        .Interior.Color = vFTIBlue
        .RowHeight = 15
        .Insert
    End With
    With ActiveSheet.Rows(Selection.row())
        .RowHeight = 10
        .Interior.ColorIndex = 0
    End With
    
    ActiveSheet.Range(sReturnLoc).Select
    With ActiveSheet.cells(Selection.row() + 1, 2)
        .Value = "Section"
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Italic = False
    End With
    With ActiveSheet.cells(Selection.row() + 1, 1)
        .Value = "-"
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = False
    End With
    ActiveCell.Offset(1).Select

End Sub

Sub Subsection()
    Dim sReturnLoc As String
    Dim iTextLoc As Integer
    Dim bCompact As Boolean
    sReturnLoc = Selection.Address

    If ActiveSheet.Rows(Selection.row() - 1).Interior.Color = vFTIBlue Then
        bCompact = True
        iTextLoc = Selection.row()
        ActiveSheet.Rows(Selection.row()).Insert
        ActiveSheet.Rows(Selection.row()).Interior.Color = vMediumBlue
        ActiveSheet.Rows(Selection.row()).RowHeight = 15
        
    Else
        bCompact = False
        iTextLoc = Selection.row() + 1
        With ActiveSheet.Rows(Selection.row())
            .RowHeight = 10
            .Interior.ColorIndex = 0
            .Insert
        End With
        With ActiveSheet.Rows(Selection.row())
            .RowHeight = 15
            .Interior.Color = vMediumBlue
            .Insert
        End With
        With ActiveSheet.Rows(Selection.row())
            .RowHeight = 10
            .Interior.ColorIndex = 0
        End With
        With ActiveSheet.cells(Selection.row() + 1, 1)
        .Value = "-"
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = False
        End With
    End If
        
    ActiveSheet.Range(sReturnLoc).Select
    With ActiveSheet.cells(iTextLoc, 2)
        .Value = "Subsection"
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Italic = False
    End With
    If bCompact = False Then
        ActiveCell.Offset(1).Select
    End If

End Sub

Sub Subsubsection()
    Dim sReturnLoc As String
    Dim iTextLoc As Integer
    Dim bCompact As Boolean
    sReturnLoc = Selection.Address
    
    If ActiveSheet.Rows(Selection.row() - 1).Interior.Color = vMediumBlue Then
        bCompact = True
        iTextLoc = Selection.row()
        ActiveSheet.Rows(Selection.row()).Insert
        ActiveSheet.Rows(Selection.row()).Interior.Color = vLightGrey
        ActiveSheet.Rows(Selection.row()).RowHeight = 15
            
    Else
        bCompact = False
        iTextLoc = Selection.row() + 1
        With ActiveSheet.Rows(Selection.row())
            .RowHeight = 10
            .Interior.ColorIndex = 0
            .Insert
        End With
        With ActiveSheet.Rows(Selection.row())
            .Interior.Color = vLightGrey
            .RowHeight = 15
            .Insert
        End With
        With ActiveSheet.Rows(Selection.row())
            .RowHeight = 10
            .Interior.ColorIndex = 0
        End With
    End If
        
    ActiveSheet.Range(sReturnLoc).Select
    With ActiveSheet.cells(iTextLoc, 2)
        .Value = "Subsubsection"
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = True
        .Font.Italic = False
    End With
    If bCompact = False Then
        ActiveCell.Offset(1).Select
    End If

End Sub

Sub SheetEnd()
    Dim sReturnLoc As String
    sReturnLoc = Selection.Address

    With ActiveSheet.Rows(Selection.row())
        .Interior.Color = vFTIBlue
        .RowHeight = 15
        .Insert
    End With
    With ActiveSheet.Rows(Selection.row())
        .RowHeight = 10
        .Interior.ColorIndex = 0
    End With
    With ActiveSheet.cells(Selection.row() + 1, 1)
        .Value = "-"
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = False
    End With

    ActiveSheet.Range(sReturnLoc).Select
    With ActiveSheet.cells(Selection.row() + 1, 2)
        .Value = "End of Sheet"
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Italic = False
    End With

End Sub

Sub TableHeader()
    With Selection.cells
        .Font.Bold = True
        .Font.Color = vFTIBlue
        .Font.Italic = False
    End With
    
End Sub

Sub Accounting(Optional ByVal zeroes As Integer = 0)
    On Error GoTo ErrorHandler
    With Selection
        Select Case zeroes
            Case 0
            .NumberFormat = sAccounting
            Case 1
            .NumberFormat = sAccounting0
            Case 2
            .NumberFormat = sAccounting00
            Case 3
            .NumberFormat = sAccounting000
            Case Else
            .NumberFormat = sAccounting
            
        End Select
    End With
ErrorHandler:
End Sub

Sub Percentage(Optional ByVal zeroes As Integer = 1)
    On Error GoTo ErrorHandler
    With Selection
        Select Case zeroes
            Case 0
            .NumberFormat = sPercentage
            Case 1
            .NumberFormat = sPercentage0
            Case 2
            .NumberFormat = sPercentage00
            Case 3
            .NumberFormat = sPercentage000
            Case Else
            .NumberFormat = sPercentage
            
        End Select
    End With
ErrorHandler:
End Sub

Sub PercentPoints(Optional ByVal zeroes As Integer = 1)
    On Error GoTo ErrorHandler
    With Selection
        Select Case zeroes
            Case 0
            .NumberFormat = sPercentPoints
            Case 1
            .NumberFormat = sPercentPoints0
            Case 2
            .NumberFormat = sPercentPoints00
            Case 3
            .NumberFormat = sPercentPoints000
            Case Else
            .NumberFormat = sPercentPoints
            
        End Select
    End With
ErrorHandler:
End Sub

Sub Multiple(Optional ByVal zeroes As Integer = 1)
    On Error GoTo ErrorHandler
    With Selection
        Select Case zeroes
            Case 0
            .NumberFormat = sMultiple
            Case 1
            .NumberFormat = sMultiple0
            Case 2
            .NumberFormat = sMultiple00
            Case 3
            .NumberFormat = sMultiple000
            Case Else
            .NumberFormat = sMultiple
            
        End Select
    End With
ErrorHandler:
End Sub

