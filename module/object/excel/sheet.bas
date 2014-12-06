Attribute VB_Name = "object_excel_sheet"
Function SheetGetAllData()
    Dim objSheet As Object
    Dim intLoop As Integer
    
    intLoop = ActiveCell.Row
    For Each objSheet In ActiveWorkbook.Sheets
        Call objSheet.Range("A1").Copy
        ActiveWorkbook.Sheets("Sheet1").Cells(intLoop, 1).PasteSpecial Paste:=xlPasteValues
        Call objSheet.Range("B6:B11").Copy
        ActiveWorkbook.Sheets("Sheet1").Cells(intLoop, ActiveCell.Column).Offset(0, 1).PasteSpecial Paste:=xlPasteValues, Transpose:=True
        intLoop = intLoop + 1
    Next
End Function

Function SheetGetName( _
    pstr As Integer _
    ) As String
    Call Application.Volatile
    
    SheetGetName = Sheets(pstr).Name
End Function

Function SheetCheck( _
    pstrWSName _
    ) As Boolean
    Dim cnt
    Dim cnt2
    
    cnt2 = ActiveWorkbook.Worksheets.count
    SheetCheck = False
    For cnt = 1 To cnt2
        If ActiveWorkbook.Worksheets(cnt).Name = pstrWSName Then
            SheetCheck = True
        End If
    Next
End Function

Sub SheetListName()
    Dim objSheet As Object
    Dim intLoop As Integer
    
    intLoop = ActiveCell.Row
    
    For Each objSheet In ActiveWorkbook.Sheets
        ActiveWorkbook.ActiveSheet.Cells(intLoop, ActiveCell.Column).Value = objSheet.Name
        intLoop = intLoop + 1
    Next
End Sub

Sub SheetDelete()
    On Error Resume Next
        Application.DisplayAlerts = False
            ActiveWindow.SelectedSheets.Delete
        Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

Sub SheetDeleteEmpty()
    Dim intLoop As Integer
    
    Application.DisplayAlerts = False
        For intLoop = Worksheets.count To 1 Step -1
            If WorksheetFunction.CountA(Worksheets(intLoop).UsedRange) = 0 Then
                If Worksheets.count > 1 Then Worksheets(intLoop).Delete
            End If
        Next intLoop
    Application.DisplayAlerts = True
End Sub

Sub SheetDeleteExclusion()
    Dim lp As Long
    
    If Sheets.count < 2 Then Exit Sub
    If MsgBox( _
        "Are you delete other than this sheet?", _
        vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    ActiveSheet.Move Before:=Sheets(1)
    
    Application.DisplayAlerts = False
        For lp = Sheets.count To 2 Step -1
             Sheets(lp).Delete
        Next lp
    Application.DisplayAlerts = True
End Sub

Sub SheetCopy()
    ActiveSheet.Copy Before:=Sheets(1)
End Sub

Sub SheetCopyToNewBook()
    ActiveSheet.Copy
End Sub
    
Sub SheetCopyToOtherBook()
    Call Application.CommandBars.FindControl(ID:=848).Execute
End Sub

Sub SheetAdd()
    Dim x As String
    Application.DisplayAlerts = False
        Sheets.Add
        Call InitializeFont
        Call DisplayNone2
        Call DisplayZoom
    Application.DisplayAlerts = True
End Sub

Sub SheetNext()
    On Error Resume Next
        ActiveSheet.Next.Select
    On Error GoTo 0
End Sub

Sub SheetPrevious()
    On Error Resume Next
        ActiveSheet.Previous.Select
    On Error GoTo 0
End Sub

Sub SheetNext2()
    On Error Resume Next
        Worksheets(Worksheets.count).Activate
    On Error GoTo 0
End Sub

Sub SheetPrevious2()
    On Error Resume Next
        Worksheets(1).Activate
    On Error GoTo 0
End Sub

Sub SheetSelectAll()
    Dim varIndex() As Variant
    Dim i As Integer
    
    For i = 1 To Sheets.count
        ReDim Preserve varIndex(i - 1)
        varIndex(i - 1) = i
    Next
    If varIndex(LBound(varIndex)) = "" Then Exit Sub
    Worksheets(varIndex).Select
End Sub

Sub SheetFreezePanes()
    If ActiveWindow.FreezePanes Then ActiveWindow.FreezePanes = False Else ActiveWindow.FreezePanes = True
End Sub

Sub SheetMoveNext()
    On Error Resume Next
        ActiveSheet.Move After:=Sheets(ActiveSheet.Next.Name)
    On Error GoTo 0
End Sub

Sub SheetMovePrevious()
    On Error Resume Next
        ActiveSheet.Move Before:=Sheets(ActiveSheet.Previous.Name)
    On Error GoTo 0
End Sub

Sub SheetRename( _
    Optional pstrName = "" _
    )
    On Error Resume Next
        If pstrName = "" Then pstrName = InputBox("What do you change sheet's name ?", , ActiveSheet.Name)
        ActiveSheet.Name = pstrName
    On Error GoTo 0
End Sub

