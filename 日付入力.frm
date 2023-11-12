VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 日付入力 
   Caption         =   "UserForm1"
   ClientHeight    =   1848
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   2820
   OleObjectBlob   =   "日付入力.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "日付入力"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Function 代入先() As Range
    Set 代入先 = ActiveCell
End Function
Private Sub UserForm_Initialize()
    TextBox1.IMEMode = fmIMEModeDisable 'IME入力モードを無効化
End Sub
Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 46, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
        Case Else: KeyAscii = 0 '"."と数値の入力以外を無効化
    End Select
End Sub
Private Sub TextBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim 和暦文字列 As String
    If KeyCode = vbKeyReturn Then
        With TextBox1
            If .Text = "." Then
                代入先 = Date
                Unload 日付入力
            End If
            On Error Resume Next
            Select Case Len(.Text)
                Case 7
                    Select Case Left(.Text, 1)
                        Case 1: 和暦文字列 = "M" & Mid(.Text, 2, 2) & "/" & Mid(.Text, 4, 2) & "/" & Mid(.Text, 6, 2)
                        Case 2: 和暦文字列 = "T" & Mid(.Text, 2, 2) & "/" & Mid(.Text, 4, 2) & "/" & Mid(.Text, 6, 2)
                        Case 3: 和暦文字列 = "S" & Mid(.Text, 2, 2) & "/" & Mid(.Text, 4, 2) & "/" & Mid(.Text, 6, 2)
                        Case 4: 和暦文字列 = "H" & Mid(.Text, 2, 2) & "/" & Mid(.Text, 4, 2) & "/" & Mid(.Text, 6, 2)
                        Case 5: 和暦文字列 = "R" & Mid(.Text, 2, 2) & "/" & Mid(.Text, 4, 2) & "/" & Mid(.Text, 6, 2)
                    End Select
                    代入先 = DateValue(和暦文字列)
                    Unload 日付入力
                Case 8
                    代入先 = DateSerial(Mid(.Text, 1, 4), Mid(.Text, 5, 2), Mid(.Text, 7, 2))
                    Unload 日付入力
            End Select
        End With
    End If
End Sub
