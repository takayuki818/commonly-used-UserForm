VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   1848
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   2820
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'I[i[ tH[Μ
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Function γόζ() As Range
    Set γόζ = ActiveCell
End Function
Private Sub UserForm_Initialize()
    TextBox1.IMEMode = fmIMEModeDisable 'IMEόΝ[hπ³ψ»
End Sub
Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 46, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
        Case Else: KeyAscii = 0 '"."ΖlΜόΝΘOπ³ψ»
    End Select
End Sub
Private Sub TextBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim aοΆρ As String
    If KeyCode = vbKeyReturn Then
        With TextBox1
            If .Text = "." Then
                γόζ = Date
                Unload UserForm1
            End If
            On Error Resume Next
            Select Case Len(.Text)
                Case 7
                    Select Case Left(.Text, 1)
                        Case 1: aοΆρ = "M" & Mid(.Text, 2, 2) & "/" & Mid(.Text, 4, 2) & "/" & Mid(.Text, 6, 2)
                        Case 2: aοΆρ = "T" & Mid(.Text, 2, 2) & "/" & Mid(.Text, 4, 2) & "/" & Mid(.Text, 6, 2)
                        Case 3: aοΆρ = "S" & Mid(.Text, 2, 2) & "/" & Mid(.Text, 4, 2) & "/" & Mid(.Text, 6, 2)
                        Case 4: aοΆρ = "H" & Mid(.Text, 2, 2) & "/" & Mid(.Text, 4, 2) & "/" & Mid(.Text, 6, 2)
                        Case 5: aοΆρ = "R" & Mid(.Text, 2, 2) & "/" & Mid(.Text, 4, 2) & "/" & Mid(.Text, 6, 2)
                    End Select
                    γόζ = DateValue(aοΆρ)
                    Unload UserForm1
                Case 8
                    γόζ = DateSerial(Mid(.Text, 1, 4), Mid(.Text, 5, 2), Mid(.Text, 7, 2))
                    Unload UserForm1
            End Select
        End With
    End If
End Sub
