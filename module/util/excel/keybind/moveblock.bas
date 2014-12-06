Attribute VB_Name = "util_excel_keybind_moveblock"
Sub MoveBlock()
    Dim myCnt As Long
    myCnt = 2
    On Error Resume Next
        Do While myCnt > 1
            Call Range(Selection, Selection.End(xlToRight)).Select
            myCnt = Selection.Cells.count - Application.WorksheetFunction.CountBlank(Selection.Cells)
            Call Range(Selection, Selection.End(xlDown)).Select
            Call Selection.Cut
            Call Selection.End(xlToLeft).Select
            Call Selection.End(xlDown).Select
            Call ActiveCell.Offset(1, 0).Select
            Call ActiveSheet.Paste
            Call ActiveCell.Offset(0, 1).Select
        Loop
    On Error GoTo 0
End Sub

Sub MoveBlock2()
    Dim myCnt As Long
    On Error Resume Next
        Call Range(Selection, Selection.End(xlToRight)).Select
        myCnt = Selection.Cells.count - Application.WorksheetFunction.CountBlank(Selection.Cells)
        
        Call Columns("A:A").Select
        Call Selection.Insert(Shift:=xlToRight)
        
        For i = 1 To myCnt - 1
            ActiveCell.Offset(0, 2).Select
            ActiveCell.EntireColumn.Select
            Selection.Insert Shift:=xlToRight
        Next
        
        Call Range(Cells(1, 2), Cells(1, myCnt * 2)).Select
        Call Selection.Cut
        Call Range("A1").Select
        Call ActiveSheet.Paste
        Call ActiveCell.Offset(0, 1).Select
        Call ActiveCell.Offset(1, 0).Select
        Call Selection.End(xlDown).Select
        Call ActiveCell.Offset(0, -1).Select
        Call Range(Selection, Selection.End(xlUp)).Select
        Call Selection.FillDown
        Call Selection.End(xlDown).Select
        
        For i = 1 To myCnt - 1
            Call ActiveCell.Offset(0, 1).Select
            Call Selection.End(xlUp).Select
            Call Selection.End(xlToRight).Select
            Call Selection.End(xlDown).Select
            Call ActiveCell.Offset(0, -1).Select
            Call Range(Selection, Selection.End(xlUp)).Select
            Call Selection.FillDown
            Call Selection.End(xlDown).Select
        Next
    
        Call Range("A1").Select
        Call ActiveCell.EntireRow.Select
        Selection.Delete Shift:=xlToTop
        Call Range("A1").Select
        Call ActiveCell.Offset(0, 2).Select
        
        For i = 1 To myCnt - 1
            Call Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
            Call Selection.Cut
            Call Selection.End(xlToLeft).Select
            Call Selection.End(xlDown).Select
            Call ActiveCell.Offset(1, 0).Select
            Call ActiveSheet.Paste
            Call ActiveCell.Offset(0, 2).Select
        Next
    On Error GoTo 0
End Sub

Sub MoveBlock3()
    Dim myCnt As Long
    Call Range(Selection, Selection.End(xlToRight)).Select
    myCnt = Selection.Cells.count / 2
    
    Call Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    
    For i = 1 To myCnt - 1
        Call ActiveCell.Offset(0, 3).Select
        Call ActiveCell.EntireColumn.Select
        Selection.Insert Shift:=xlToRight
    Next
    
    Call Range(Cells(1, 2), Cells(1, myCnt * 3)).Select
    Call Selection.Cut
    Call Range("A1").Select
    Call ActiveSheet.Paste
    Call ActiveCell.Offset(0, 1).Select
    Call ActiveCell.Offset(1, 0).Select
    Call Selection.End(xlDown).Select
    Call ActiveCell.Offset(0, -1).Select
    Call Range(Selection, Selection.End(xlUp)).Select
    Call Selection.FillDown
    Call Selection.End(xlDown).Select
    
    For i = 1 To myCnt - 1
        Call ActiveCell.Offset(0, 2).Select
        Call Selection.End(xlUp).Select
        Call Selection.End(xlToRight).Select
        Call Selection.End(xlDown).Select
        Call ActiveCell.Offset(0, -1).Select
        Call Range(Selection, Selection.End(xlUp)).Select
        Call Selection.FillDown
        Call Selection.End(xlDown).Select
    Next

    Call Range("A1").Select
    Call ActiveCell.EntireRow.Select
    Selection.Delete Shift:=xlToTop
    Call Range("A1").Select
    Call ActiveCell.Offset(0, 3).Select
    
    For i = 1 To myCnt - 1
        Call Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
        Call Selection.Cut
        Call Selection.End(xlToLeft).Select
        Call Selection.End(xlDown).Select
        Call ActiveCell.Offset(1, 0).Select
        Call ActiveSheet.Paste
        Call ActiveCell.Offset(0, 3).Select
    Next
End Sub

Sub MoveBlock4()
    On Error Resume Next
        For i = 1 To 20
            Call Range("A1").Select
            Call Selection.End(xlDown).Select
            Call Selection.End(xlDown).Select
            Call Range(Selection, Selection.End(xlDown)).Select
            Call Range(Selection, Selection.End(xlToRight)).Select
            Call Selection.Cut
            Call Selection.End(xlToRight).Select
            Call Selection.End(xlToRight).Select
            Call Selection.End(xlUp).Select
            Call Selection.End(xlToLeft).Select
            Call ActiveCell.Offset(0, 2).Select
            Call ActiveSheet.Paste
        Next
    On Error GoTo 0
End Sub

Sub MoveBlock5()
    Call Range("A1").Select
    Call ActiveCell.Offset(0, 1).Select
    On Error Resume Next
        For i = 1 To 20
            Call MoveBlock
            Call Selection.End(xlUp).Select
            Call Selection.End(xlToRight).Select
            Call ActiveCell.Offset(0, 1).Select
        Next
    On Error GoTo 0
End Sub

Sub MoveBlock6()
    Call Range("A1").Select
    Call Selection.End(xlToRight).Select
    On Error Resume Next
        For i = 1 To 20
            Call Range(Selection, Selection.End(xlDown)).Select
            Call Selection.Cut
            Call Range("A1").Select
            Call Selection.End(xlDown).Select
            Call ActiveCell.Offset(1, 0).Select
            Call ActiveSheet.Paste
            Call Selection.End(xlUp).Select
            Call Selection.End(xlToRight).Select
        Next
    On Error GoTo 0
End Sub

