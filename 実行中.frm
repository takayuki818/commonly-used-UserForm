VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 実行中 
   Caption         =   "UserForm1"
   ClientHeight    =   888
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   4284
   OleObjectBlob   =   "実行中.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "実行中"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim 始時 As Date, 終時 As Date
Private Sub UserForm_Initialize()
    Label1.Width = 0
    Label2.Caption = "0％完了"
    始時 = Now
End Sub
Sub プログレスバー更新(進捗 As Long)
    If 進捗 Mod 10 = 0 Then
        Label1.Width = 進捗 * 2
        Label2.Caption = 進捗 & "％完了"
        DoEvents
    End If
    If 進捗 = 100 Then
        終時 = Now
        MsgBox "処理が完了しました" & vbCrLf & vbCrLf & "処理時間：" & 終時 - 始時
    End If
End Sub
