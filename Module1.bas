Attribute VB_Name = "Module1"
Option Explicit
Sub 時間のかかるコードサンプル()
    Dim 行 As Long
    実行中.Show vbModeless
    For 行 = 1 To 1000000
        実行中.プログレスバー更新 (Int(行 / 10000))
    Next
    Unload 実行中
End Sub
