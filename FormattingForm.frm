VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormattingForm 
   Caption         =   "Formatting"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2400
   OleObjectBlob   =   "FormattingForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormattingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Const kEsc = 27, k1 = 49, k2 = 50, k3 = 51, k4 = 52, k5 = 53, kQ = 113, kE = 101, kA = 97, kD = 100

' Declaration of API function -- for 32 bit systems and 64
#If VBA7 Then
    Private Declare PtrSafe Function GetKeyState Lib "USER32" (ByVal nKey As Long) As Integer
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    Private Declare Function GetKeyState Lib "USER32" (ByVal nKey As Long) As Integer
    Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If

Private Const keyMask As Integer = &HFF80
'Const holdTime As String = "00:00:01" ' minimum of 1 second

' converts from ascii keys to virtual keys
Private Function AsciiToVirtual(ByVal keyAscii As Long)

    Select Case keyAscii
        Case 113  ' Q
            AsciiToVirtual = 81
        Case 101  ' E
            AsciiToVirtual = 69
        Case 97  ' A
            AsciiToVirtual = 65
        Case 100  ' D
            AsciiToVirtual = 68
        Case Else
    End Select

End Function

' a wait function that doesn't tie up excel, uses c. 1000 tick / second
Sub tickCounter(duration As Long)
    Dim nowTick, endTick, counter As Long
    
    endTick = GetTickCount + (duration * 1000)
    counter = 0
    
    Do
        nowTick = GetTickCount
        counter = counter + 1
        DoEvents
    Loop Until nowTick >= endTick And counter >= 500
    
End Sub


Function decimalPlaces(ByVal key As Integer)
'    On Error Resume Next
    tickCounter (0.2)
    
    Dim response As Integer
    If Not CBool(GetKeyState(key) And keyMask) Then
        response = 0
        GoTo ReturnDecimals
    End If
    
    With FormattingProgressBar
        .Top = Application.Height / 2 - (FormattingProgressBar.Height / 2)
        .Left = Application.Width / 2 - (FormattingProgressBar.Width / 2)
        .Show vbModeless
    End With
    
    FormattingProgressBar.dpLabel.Caption = 1
    tickCounter (0.75)
    
    If Not CBool(GetKeyState(key) And keyMask) Then
        response = 1
        GoTo ReturnDecimals
    End If
    
    FormattingProgressBar.dpLabel.Caption = 2
    DoEvents
    tickCounter (1)
    
    If Not CBool(GetKeyState(key) And keyMask) Then
        response = 2
        GoTo ReturnDecimals
    End If
    
    FormattingProgressBar.dpLabel.Caption = 3
    DoEvents
    tickCounter (1)
    
    If Not CBool(GetKeyState(key) And keyMask) Then
        response = 3
    Else
        response = 0
    End If
    
ReturnDecimals:
    FormattingProgressBar.Hide
    decimalPlaces = response
End Function



' Unloading sub for use by userform button handler.
Private Sub UnloadHandler()
    Unload Me
End Sub

' Various subs that handle button clicks

Private Sub ButtonCancel_Click()
    Unload Me
End Sub

Private Sub ButtonEnd_Click()
    Call SheetEnd
    
    Unload Me
End Sub

Private Sub ButtonHeader_Click()
    Call TableHeader
    
    Unload Me
End Sub

Private Sub ButtonSection_Click()
    Call Section
    
    Unload Me
End Sub

Private Sub ButtonSubsection_Click()
    Call Subsection
    
    Unload Me
End Sub

Private Sub ButtonSubsubsection_Click()
    Call Subsubsection
    
    Unload Me
End Sub

Private Sub ButtonAccounting_Click()
    Call Accounting
    
    Unload Me
End Sub

Private Sub ButtonMultiple_Click()
    Call Multiple
    
    Unload Me
End Sub

Private Sub ButtonPercentage_Click()
    Call Percentage
    
    Unload Me
End Sub

Private Sub ButtonPercentPoints_Click()
    Call PercentPoints
    
    Unload Me
End Sub


' MASTER KEY PRESS HANDLER

' Master keystroke handler at userform level
Private Sub FormattingForm_KeyPress(ByVal keyAscii As MSForms.ReturnInteger)
    FormattingForm.Hide
    Select Case keyAscii
        Case kEsc
            Call UnloadHandler
        Case k1
            Call ButtonSection_Click
        Case k2
            Call ButtonSubsection_Click
        Case k3
            Call ButtonSubsubsection_Click
        Case k4
            Call ButtonEnd_Click
        Case k5
            Call ButtonHeader_Click
        Case kQ
            Accounting (decimalPlaces(AsciiToVirtual(keyAscii)))
        Case kE
            Multiple (decimalPlaces(AsciiToVirtual(keyAscii)))
        Case kA
            Percentage (decimalPlaces(AsciiToVirtual(keyAscii)))
        Case kD
            PercentPoints (decimalPlaces(AsciiToVirtual(keyAscii)))
        Case Else

    End Select
End Sub

' PER-BUTTON KEYPRESS HANDLERS

' Keypress handlers for each button that can pull focus, passes on to master
Private Sub ButtonSection_KeyPress(ByVal keyAscii As MSForms.ReturnInteger)
    Call FormattingForm_KeyPress(keyAscii)
End Sub

Private Sub ButtonSubsection_KeyPress(ByVal keyAscii As MSForms.ReturnInteger)
    Call FormattingForm_KeyPress(keyAscii)
End Sub

Private Sub ButtonSubsubsection_KeyPress(ByVal keyAscii As MSForms.ReturnInteger)
    Call FormattingForm_KeyPress(keyAscii)
End Sub

Private Sub ButtonEnd_KeyPress(ByVal keyAscii As MSForms.ReturnInteger)
    Call FormattingForm_KeyPress(keyAscii)
End Sub

Private Sub ButtonHeader_KeyPress(ByVal keyAscii As MSForms.ReturnInteger)
    Call FormattingForm_KeyPress(keyAscii)
End Sub

Private Sub ButtonAccounting_KeyPress(ByVal keyAscii As MSForms.ReturnInteger)
    Call FormattingForm_KeyPress(keyAscii)
End Sub

Private Sub ButtonMultiple_KeyPress(ByVal keyAscii As MSForms.ReturnInteger)
    Call FormattingForm_KeyPress(keyAscii)
End Sub

Private Sub ButtonPercentage_KeyPress(ByVal keyAscii As MSForms.ReturnInteger)
    Call FormattingForm_KeyPress(keyAscii)
End Sub

Private Sub ButtonPercentPoints_KeyPress(ByVal keyAscii As MSForms.ReturnInteger)
    Call FormattingForm_KeyPress(keyAscii)
End Sub


