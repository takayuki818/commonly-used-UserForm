VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ���t���� 
   Caption         =   "UserForm1"
   ClientHeight    =   1848
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   2820
   OleObjectBlob   =   "���t����.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "���t����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Function �����() As Range
    Set ����� = ActiveCell
End Function
Private Sub UserForm_Initialize()
    TextBox1.IMEMode = fmIMEModeDisable 'IME���̓��[�h�𖳌���
End Sub
Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 46, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
        Case Else: KeyAscii = 0 '"."�Ɛ��l�̓��͈ȊO�𖳌���
    End Select
End Sub
Private Sub TextBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim �a����� As String
    If KeyCode = vbKeyReturn Then
        With TextBox1
            If .Text = "." Then
                ����� = Date
                Unload ���t����
            End If
            On Error Resume Next
            Select Case Len(.Text)
                Case 7
                    Select Case Left(.Text, 1)
                        Case 1: �a����� = "M" & Mid(.Text, 2, 2) & "/" & Mid(.Text, 4, 2) & "/" & Mid(.Text, 6, 2)
                        Case 2: �a����� = "T" & Mid(.Text, 2, 2) & "/" & Mid(.Text, 4, 2) & "/" & Mid(.Text, 6, 2)
                        Case 3: �a����� = "S" & Mid(.Text, 2, 2) & "/" & Mid(.Text, 4, 2) & "/" & Mid(.Text, 6, 2)
                        Case 4: �a����� = "H" & Mid(.Text, 2, 2) & "/" & Mid(.Text, 4, 2) & "/" & Mid(.Text, 6, 2)
                        Case 5: �a����� = "R" & Mid(.Text, 2, 2) & "/" & Mid(.Text, 4, 2) & "/" & Mid(.Text, 6, 2)
                    End Select
                    ����� = DateValue(�a�����)
                    Unload ���t����
                Case 8
                    ����� = DateSerial(Mid(.Text, 1, 4), Mid(.Text, 5, 2), Mid(.Text, 7, 2))
                    Unload ���t����
            End Select
        End With
    End If
End Sub
