Attribute VB_Name = "Module1"
Option Explicit
Sub ���Ԃ̂�����R�[�h�T���v��()
    Dim �s As Long
    ���s��.Show vbModeless
    For �s = 1 To 1000000
        ���s��.�v���O���X�o�[�X�V (Int(�s / 10000))
    Next
    Unload ���s��
End Sub
