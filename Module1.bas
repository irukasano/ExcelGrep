Attribute VB_Name = "Module1"
Option Explicit

Public EG As New ExcelGrep

Public Sub 検索対象フォルダのパスを入力()
    Call EG.PickupFolderPath("検索対象フォルダを選択してください。")
End Sub

Public Sub 検索実行高速()
    Call EG.ExecSearch(Express:=True)
End Sub

Public Sub 検索実行()
    Call EG.ExecSearch(IgnoreCase:=True)
End Sub

Public Sub 検索実行_大文字小文字を区別()
    Call EG.ExecSearch(IgnoreCase:=False)
End Sub

Public Sub 検索中止()
    Call EG.Interrupt
End Sub

Public Sub 結果リストをクリア()
    Call EG.ClearResultList
End Sub

Public Sub 結果のブックをSVNLOCK()
    Call EG.LockResultList
End Sub

Public Sub 結果のブックをSVNCOMMIT()
    Call EG.CommitResultList
End Sub

