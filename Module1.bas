Attribute VB_Name = "Module1"
Option Explicit

Public EG As New ExcelGrep

Public Sub �����Ώۃt�H���_�̃p�X�����()
    Call EG.PickupFolderPath("�����Ώۃt�H���_��I�����Ă��������B")
End Sub

Public Sub �������s����()
    Call EG.ExecSearch(Express:=True)
End Sub

Public Sub �������s()
    Call EG.ExecSearch(IgnoreCase:=True)
End Sub

Public Sub �������s_�啶�������������()
    Call EG.ExecSearch(IgnoreCase:=False)
End Sub

Public Sub �������~()
    Call EG.Interrupt
End Sub

Public Sub ���ʃ��X�g���N���A()
    Call EG.ClearResultList
End Sub

Public Sub ���ʂ̃u�b�N��SVNLOCK()
    Call EG.LockResultList
End Sub

Public Sub ���ʂ̃u�b�N��SVNCOMMIT()
    Call EG.CommitResultList
End Sub

