Attribute VB_Name = "Module1"
Option Explicit
Sub ���Ԃ̂�����R�[�h�T���v��()
    Dim �s As Long, �i�� As Long
    ���s��.Show vbModeless
    Call ���s��.�v���O���X�o�[�X�V(0)
    For �s = 1 To 10000
        Sheet1.Cells(1, 1) = �s
        If Int(�s / 100) - �i�� >= 10 Then
            �i�� = Int(�s / 100)
            Call ���s��.�v���O���X�o�[�X�V(�i��)
        End If
    Next
    Sheet1.Cells(1, 1).ClearContents
    Unload ���s��
End Sub
