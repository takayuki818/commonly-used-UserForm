VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ���s�� 
   Caption         =   "UserForm1"
   ClientHeight    =   888
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   4284
   OleObjectBlob   =   "���s��.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "���s��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim �n�� As Date, �I�� As Date
Private Sub UserForm_Initialize()
    Label1.Width = 0
    Label2.Caption = "0������"
    �n�� = Now
End Sub
Sub �v���O���X�o�[�X�V(�i�� As Long)
    If �i�� Mod 10 = 0 Then
        Label1.Width = �i�� * 2
        Label2.Caption = �i�� & "������"
        DoEvents
    End If
    If �i�� = 100 Then
        �I�� = Now
        MsgBox "�������������܂���" & vbCrLf & vbCrLf & "�������ԁF" & �I�� - �n��
    End If
End Sub
