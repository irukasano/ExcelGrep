VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelGrep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private MainSheet           As Worksheet
Private TargetPathCell      As Range
Private SearchPatternCell   As Range
Private ResultListFirstCell As Range
Private ShouldSearch        As Boolean
Private FSO                 As New FileSystemObject
Private REG                 As New VBScript_RegExp_55.RegExp
Private InvisibleExcel      As Excel.Application
Private ExpressSearch       As Boolean
Private TortoiseSVNProc     As String

Private Enum ResultColumns
    path = 1
    Book = 2
    Sheet = 3
    Name = 4
    Value = 5
End Enum

Private Sub Class_Initialize()
    Set MainSheet = ThisWorkbook.Sheets("ExcelGrep")
    Set TargetPathCell = MainSheet.Range("C3")
    Set SearchPatternCell = MainSheet.Range("C4")
    Set ResultListFirstCell = MainSheet.Range("C7")
    ShouldSearch = False
    TortoiseSVNProc = MainSheet.Range("C2").Value
End Sub

Public Sub PickupFolderPath(Msg As String)
    Dim FolderPicker As FileDialog
    Dim Result       As String
    Set FolderPicker = Application.FileDialog(msoFileDialogFolderPicker)
    
    FolderPicker.Title = Msg
    If FolderPicker.Show Then Result = FolderPicker.SelectedItems(1)
    If Result = "" Then Exit Sub
    
    TargetPathCell.Value = Result
End Sub

Public Sub ExecSearch(Optional IgnoreCase As Boolean = True, Optional Express As Boolean = False)
    Dim SpecifiedFolder As Folder: Set SpecifiedFolder = GetSpecifiedFolder()
    If SpecifiedFolder Is Nothing Then
        MsgBox "検索対象フォルダが見つかりませんでした。", vbInformation, ThisWorkbook.Name
        Exit Sub
    End If
    If Trim(SearchPatternCell.Value) = "" Then
        MsgBox "検索文字列を入力してください。", vbInformation, ThisWorkbook.Name
        Exit Sub
    End If
    
    If ShouldSearch = True Then
        MsgBox "別の検索処理がすでに実行中です。", vbInformation, ThisWorkbook.Name
        Exit Sub
    End If
    
    REG.Global = True
    REG.IgnoreCase = IgnoreCase
    REG.Pattern = SearchPatternCell.Value
    ExpressSearch = Express
    
    ShouldSearch = True
    Call ClearResultList
    Call SearchFolder(SpecifiedFolder)
    
    If Not InvisibleExcel Is Nothing Then
        InvisibleExcel.Quit
        Set InvisibleExcel = Nothing
    End If
    
    ShouldSearch = False
    Call DisplayStatus("")
    MsgBox "検索が終了しました。", vbInformation, ThisWorkbook.Name
End Sub

Public Sub Interrupt()
    If ShouldSearch = False Then Exit Sub
    If MsgBox("検索を中止してもよろしいですか?", vbYesNo Or vbQuestion, ThisWorkbook.Name) <> vbYes Then Exit Sub
    
    ShouldSearch = False
End Sub

Public Sub ClearResultList()
    ResultListFirstCell.CurrentRegion.Offset(2).Delete
End Sub

Private Function GetSpecifiedFolder() As Folder
    On Error Resume Next
    Dim SpecifiedFolder As Folder
    Set SpecifiedFolder = FSO.GetFolder(TargetPathCell.Value)
    Set GetSpecifiedFolder = SpecifiedFolder
End Function

Private Sub SearchFolder(objFolder As Folder)
    If ShouldSearch = False Then Exit Sub
    
    Dim objFile   As File
    Dim SubFolder As Folder
    
    Call DisplayStatus(objFolder.path)
    
    For Each objFile In objFolder.Files
        Select Case FSO.GetExtensionName(objFile.path)
            Case "xls", "xlsx", "xlsm"
                Call SearchBook(objFile)
        End Select
    Next
    
    For Each SubFolder In objFolder.SubFolders
        Call SearchFolder(SubFolder) 'recursive call
    Next
End Sub

Private Sub SearchBook(objFile As File)
    If ShouldSearch = False Then Exit Sub
    
    On Error Resume Next
    If InvisibleExcel Is Nothing Then
        Set InvisibleExcel = New Excel.Application
        InvisibleExcel.Visible = False
        InvisibleExcel.ScreenUpdating = False
    End If
    
    Dim Book As Workbook: Set Book = InvisibleExcel.Workbooks.Open(Filename:=objFile.path, ReadOnly:=True)
    If Book Is Nothing Then
        MsgBox objFile.path & vbCrLf & " が開けませんでした。"
        Exit Sub
    End If
    On Error GoTo 0
    
    Dim Sheet      As Worksheet
    For Each Sheet In Book.Worksheets
        If ExpressSearch = True Then
            Call SearchSheet2(Sheet)
        Else
            Call SearchSheet(Sheet)
        End If
    Next
    
    Call Book.Close(SaveChanges:=False)
End Sub

Private Sub SearchSheet(Sheet As Worksheet)
    If ShouldSearch = False Then Exit Sub
    
    Dim TargetRange As Range
    Dim Cell        As Range

    'Search Cells
    Set TargetRange = Sheet.UsedRange.Cells
    For Each Cell In TargetRange
        DoEvents
        Call DisplayStatus(Sheet.Parent.FullName + "(" + Cell.Address + ")")
        
        If Cell.Value <> "" Then
            If REG.Test(Cell.Value) Then
                Call ProcessCell(Cell)
            End If
        End If
    Next
    
    'Search Shapes
    Dim objShape As Shape
    For Each objShape In Sheet.Shapes
        DoEvents
        Call DisplayStatus(Sheet.Parent.FullName + "(" + objShape.Name + ")")
        
        If HasTextFrameCharactersText(objShape) Then
            If REG.Test(objShape.TextFrame.Characters.Text) Then
                Call ProcessShape(objShape)
            End If
        End If
    Next
End Sub

Private Sub SearchSheet2(Sheet As Worksheet)
    If ShouldSearch = False Then Exit Sub
    
    Dim TargetRange     As Range
    Dim r               As Range
    Dim firstAddress    As String

    'Search Cells
    Set TargetRange = Sheet.UsedRange.Cells
    Set r = TargetRange.Find(SearchPatternCell.Value)
    If (Not r Is Nothing) Then
        firstAddress = r.Address
        
        Do
            DoEvents
            Call DisplayStatus(Sheet.Parent.FullName + "(" + r.Address + ")")
            Call ProcessCell(r)
            
            r = TargetRange.FindNext(r)
            If (r Is Nothing) Then
                Exit Do
            End If
        Loop Until r.Address = firstAddress
    End If

End Sub

Private Function HasTextFrameCharactersText(objShape As Shape) As Boolean
    On Error Resume Next
    Dim Text As String
    Text = objShape.TextFrame.Characters.Text
    HasTextFrameCharactersText = (Text <> "")
End Function

Private Sub ProcessCell(Cell As Range)
    Call SetNewRowData(FoundSheet:=Cell.Parent, Name:=Cell.Address, Value:=Cell.Value)
End Sub

Private Sub ProcessShape(objShape As Shape)
    Call SetNewRowData(FoundSheet:=objShape.Parent, Name:=objShape.Name, Value:=objShape.TextFrame.Characters.Text)
End Sub

Private Sub SetNewRowData(FoundSheet As Worksheet, Name As String, Value As String)
    Dim row      As Range: Set row = GetNewRow()
    Dim PathCell As Range: Set PathCell = row.Cells(ResultColumns.path)
    Dim BookCell As Range: Set BookCell = row.Cells(ResultColumns.Book)
    
    'パス
    PathCell.Value = FoundSheet.Parent.FullName
    PathCell.WrapText = False
    Call MainSheet.Hyperlinks.Add(Anchor:=PathCell, Address:=PathCell.Value)
    'ブック
    BookCell.Value = FoundSheet.Parent.Name
    Call MainSheet.Hyperlinks.Add(Anchor:=BookCell, Address:=PathCell.Value)
    'シート
    row.Cells(ResultColumns.Sheet).Value = FoundSheet.Name
    '名前
    row.Cells(ResultColumns.Name).Value = Name
    '値
    row.Cells(ResultColumns.Value).Value = Value
    row.Cells(ResultColumns.Value).WrapText = False
    '--- 罫線 ---
    row.Borders.LineStyle = xlContinuous
End Sub

Private Function GetNewRow() As Range
    Dim ListRange   As Range: Set ListRange = GetResultListRange
    Dim NewRowIndex As Long: NewRowIndex = ListRange.Rows.Count + 1
    Set GetNewRow = ListRange.Rows(NewRowIndex)
End Function

Private Function GetResultListRange() As Range
    Dim ListRowsCount As Long: ListRowsCount = ResultListFirstCell.CurrentRegion.Rows.Count - 1
    Set GetResultListRange = ResultListFirstCell.CurrentRegion.Offset(1).Resize(ListRowsCount)
End Function

Private Sub DisplayStatus(Msg As String)
    Const DotsLen  As Integer = 5
    Dim Dots       As String: Dots = String((Math.Rnd * DotsLen), ".")
    Dim Loading    As String: Loading = Left(Dots & String(DotsLen, " "), DotsLen)
    Dim strDisplay As String: strDisplay = "検索中" & Loading & " " & Msg
    Application.StatusBar = IIf(Msg = "", "", strDisplay)
    DoEvents
End Sub

Private Function IsIncludedInListRange(target As Range) As Boolean
    Dim ResultRange As Range
    Set ResultRange = Application.Intersect(GetResultListRange, target)
    IsIncludedInListRange = Not ResultRange Is Nothing
End Function

Public Sub FollowHyperlink(target As Hyperlink)
    If Not IsIncludedInListRange(target.Range) Then Exit Sub
    
    Dim SourceCell  As Range:     Set SourceCell = target.Range
    Dim SourceSheet As Worksheet: Set SourceSheet = SourceCell.Parent
    Dim SourceRow   As Range:     Set SourceRow = SourceSheet.Range(SourceCell.End(xlToRight).End(xlToLeft), SourceCell.End(xlToRight))
    Dim SheetName   As String:        SheetName = SourceRow.Cells(ResultColumns.Sheet)
    Dim ObjectName  As String:        ObjectName = SourceRow.Cells(ResultColumns.Name)
    Dim Book        As Workbook:  Set Book = ActiveSheet.Parent
    
    On Error Resume Next
    Dim IsRangeObject As Boolean: IsRangeObject = (ObjectName Like "$*")
    Dim DistSheet     As Worksheet: Set DistSheet = Book.Sheets(SheetName)
    DistSheet.Activate
    If IsRangeObject Then
        DistSheet.Range(ObjectName).Activate
    Else
        DistSheet.Shapes(ObjectName).Select
    End If
End Sub

Public Sub LockResultList()
    Dim path()      As String
    Dim cmd         As String
    
    path = GetResultFilePaths()

    If UBound(path) = 0 Then
        Call MsgBox("結果リストがありませんので処理しません", vbExclamation)
        Exit Sub
    End If
    

    cmd = Chr(34) & TortoiseSVNProc & Chr(34) & " " & _
        "/command:lock" & " " & _
        "/path:" & Chr(34) & Join(path, "*") & Chr(34) & " " & _
        "/closeonend:1"
    
    Call Shell(cmd, vbNormalFocus)

End Sub

Private Function GetResultFilePaths() As String()
    Dim i           As Long
    Dim r           As Range
    Dim ListRange   As Range: Set ListRange = GetResultListRange
    Dim prevPath    As String
    Dim path()      As String
    Dim c           As Long

    If ListRange.Rows.Count = 1 Then
        GetResultFilePaths = path
        Exit Function
    End If

    prevPath = vbNullString
    
    For i = 2 To ListRange.Rows.Count
        Set r = ListRange.Rows(i)
        Dim PathCell As Range: Set PathCell = r.Cells(ResultColumns.path)

        If (prevPath <> PathCell.Value) Then
            If Not Not path Then
                c = UBound(path) + 1
            Else
                c = 0
            End If
            ReDim Preserve path(0 To c)
        
            path(c) = PathCell.Value
        End If
        
        prevPath = PathCell.Value
    Next

    GetResultFilePaths = path()
End Function

Public Sub CommitResultList()

    Dim path()      As String
    Dim cmd         As String
    
    path = GetResultFilePaths()

    If UBound(path) = 0 Then
        Call MsgBox("結果リストがありませんので処理しません", vbExclamation)
        Exit Sub
    End If

    cmd = Chr(34) & TortoiseSVNProc & Chr(34) & " " & _
        "/command:commit" & " " & _
        "/path:" & Chr(34) & Join(path, "*") & Chr(34) & " " & _
        "/closeonend:1"
    
    Call Shell(cmd, vbNormalFocus)

End Sub
