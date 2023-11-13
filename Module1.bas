Attribute VB_Name = "Module1"
Option Explicit
Sub 時間のかかるコードサンプル()
    Dim 行 As Long, 進捗 As Long
    実行中.Show vbModeless
    Call 実行中.プログレスバー更新(0)
    For 行 = 1 To 10000
        Sheet1.Cells(1, 1) = 行
        If Int(行 / 100) - 進捗 >= 10 Then
            進捗 = Int(行 / 100)
            Call 実行中.プログレスバー更新(進捗)
        End If
    Next
    Sheet1.Cells(1, 1).ClearContents
    Unload 実行中
End Sub
