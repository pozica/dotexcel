Attribute VB_Name = "object_excel_cell"
Sub CellDeleteBlankRows()
    Dim lngLstRow As Long
    Dim lngLop As Long

    lngLstRow = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
    Application.ScreenUpdating = False
        For lngLop = lngLstRow To 1 Step -1
            If Application.WorksheetFunction.CountA(Rows(lngLop)) = 0 Then Rows(lngLop).Delete
        Next lngLop
    Application.ScreenUpdating = True
End Sub

Sub CellFromAlignToValign()
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strData As String
    Dim varData As Variant
    Dim lngLop As Long
    Dim rngOutPut As Range
    Dim x As String
    
    If Selection.count = 1 Then
        x = Application.InputBox(prompt:="Please Select." & vbCrLf & vbCrLf & "[1] All" & vbCrLf & "[2] CurrentRegion", Type:=2)
        If x = "" Then Exit Sub
        Select Case x
            Case "1": Range("A1").Select: Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
            Case "2": Selection.CurrentRegion.Select
        End Select
    End If
    For lngCol = Selection(1).Column To Selection(Selection.count).Column
        For lngRow = Selection(1).Row To Selection(Selection.count).Row
            If Not IsEmpty(Cells(lngRow, lngCol).Value) Then
                strData = strData & vbTab & Cells(lngRow, lngCol).Value
            End If
        Next lngRow
    Next lngCol
    Selection.ClearContents
    varData = Split(strData, vbTab)
    On Error Resume Next
        Set rngOutPut = Application.InputBox(prompt:="Which cell do you output?", Type:=8)
        For lngLop = 1 To UBound(varData)
            rngOutPut.Offset(lngLop - 1, 0).Value = varData(lngLop)
        Next
End Sub

Sub CellFromFunctionToString()
    Dim c As Range
    
    For Each c In Selection
        With c
            If .HasFormula Then .Value = "'" & .Formula
        End With
    Next c
End Sub

Sub CellSelectFunction()
    ActiveCell.SpecialCells(xlCellTypeFormulas).Select
End Sub

Sub CellSort()
    Call Application.CommandBars.FindControl(ID:=928).Execute
End Sub

Sub CellColumnWidth( _
    Optional pstrWidth = 15 _
    )
    Selection.ColumnWidth = pstrWidth
End Sub

Sub CellRowHeight( _
    Optional pstrHeight = 13.5 _
    )
    Selection.RowHeight = pstrHeight
End Sub

Sub CellInsert()
    On Error Resume Next
        Application.CommandBars.FindControl(ID:=295).Execute
    On Error GoTo 0
End Sub

Sub CellCount()
    Call MsgBox(Selection.count)
End Sub

Sub CellTest()
    aa = Cells(Rows.count, "A").End(xlUp).Row
    bb = Cells(Rows.count, "B").End(xlUp).Row
    cc = Cells(Rows.count, "C").End(xlUp).Row
    dd = Cells(Rows.count, "D").End(xlUp).Row
    ee = Cells(Rows.count, "E").End(xlUp).Row
    ff = Cells(Rows.count, "F").End(xlUp).Row
    gg = Cells(Rows.count, "G").End(xlUp).Row
    hh = Cells(Rows.count, "H").End(xlUp).Row
    ii = Cells(Rows.count, "I").End(xlUp).Row
    jj = Cells(Rows.count, "J").End(xlUp).Row
    kk = Cells(Rows.count, "K").End(xlUp).Row
    x = 1
    For i01 = 1 To aa
        If Cells(i01, "A") <> "" Then Cells(x, "L") = Cells(x, "L") & Cells(i01, "A") & ";"
        For i02 = 1 To bb
            If Cells(i02, "B") <> "" Then Cells(x, "L") = Cells(x, "L") & Cells(i02, "B") & ";"
            For i03 = 1 To cc
                If Cells(i03, "C") <> "" Then Cells(x, "L") = Cells(x, "L") & Cells(i03, "C") & ";"
                For i04 = 1 To dd
                    If Cells(i04, "D") <> "" Then Cells(x, "L") = Cells(x, "L") & Cells(i04, "D") & ";"
                    For i05 = 1 To ee
                        If Cells(i05, "E") <> "" Then Cells(x, "L") = Cells(x, "L") & Cells(i05, "E") & ";"
                        For i06 = 1 To ff
                            If Cells(i06, "F") <> "" Then Cells(x, "L") = Cells(x, "L") & Cells(i06, "F") & ";"
                            For i07 = 1 To gg
                                If Cells(i07, "G") <> "" Then Cells(x, "L") = Cells(x, "L") & Cells(i07, "G") & ";"
                                For i08 = 1 To hh
                                    If Cells(i08, "H") <> "" Then Cells(x, "L") = Cells(x, "L") & Cells(i08, "H") & ";"
                                    For i09 = 1 To ii
                                        If Cells(i09, "I") <> "" Then Cells(x, "L") = Cells(x, "L") & Cells(i09, "I") & ";"
                                        For i10 = 1 To jj
                                            If Cells(i10, "J") <> "" Then Cells(x, "L") = Cells(x, "L") & Cells(i10, "J") & ";"
                                            For i11 = 1 To kk
                                                If Cells(i11, "K") <> "" Then Cells(x, "L") = Cells(x, "L") & Cells(i11, "K") & ";"
                                                x = x + 1
                                            Next i11
                                        Next i10
                                    Next i09
                                Next i08
                            Next i07
                        Next i06
                    Next i05
                Next i04
            Next i03
        Next i02
    Next i01
End Sub

Function CellSelect( _
    pstrCell As String _
)
    Range(ActiveCell, pstrCell).Select
End Function

Function CellUnique( _
    pstrCell As String _
)
    Range(ActiveCell, pstrCell).AdvancedFilter Action:=xlFilterInPlace, Unique:=True
End Function

Sub CellShowAll()
    ActiveSheet.ShowAllData
End Sub

Sub CellShrink()
    With Selection
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub

Function CellMerge( _
    pstrQueries As String _
)
    Dim rng As Range
    Dim a As Variant
    Dim b As Long
    Dim x As String
    Dim y As Variant
    Dim z As String
    Dim i As Long
    Dim j As Long
    
    a = Split(pstrQueries, "|")
    b = UBound(a)
    For Each rng In Selection
        x = x & "," & rng.Row
    Next
    y = Split(x, ",")
    On Error Resume Next
        For i = 1 To UBound(y)
            z = ""
            For j = 0 To b
                If Range(a(j) & y(i)) <> "" Then z = Range(a(j) & y(i))
            Next j
            ActiveCell.Offset(i - 1, 0) = z
        Next i
End Function

Function CellIsDuplicated( _
    pstrQuery As String _
)
    Dim rng As Range
    Dim x As String
    Dim y As Variant
    Dim z As String
    Dim i As Long
    
    For Each rng In Selection
        x = x & "," & rng.Row
    Next
    y = Split(x, ",")
    On Error Resume Next
        For i = 1 To UBound(Split(x, ","))
            If Range(pstrQuery & y(i)) = Range(pstrQuery & y(i - 1)) Then
                z = True
            Else
                z = False
            End If
            ActiveCell.Offset(i - 1, 0) = z
        Next i
End Function

Function CellPrimaryKey( _
    pstrQuery1 As String, _
    pstrQuery2 As String, _
    Optional pstrQuery3 = "" _
    )
    Dim rng As Range
    Dim x As String
    Dim y As Variant
    Dim z As String
    Dim i As Long
    Dim sc
    Dim js
    Set sc = CreateObject("ScriptControl")
        sc.Language = "Jscript"
        Set js = sc.CodeObject
            For Each rng In Selection
                x = x & "," & rng.Row
            Next
            y = Split(x, ",")
            On Error Resume Next
                For i = 1 To UBound(Split(x, ","))
                    z = Range(pstrQuery1 & y(i)) & Range(pstrQuery2 & y(i)) & Range(pstrQuery3 & y(i))
                    z = FormatString(z)
                    z = replace(z, " ", "")
                    z = js.encodeURIComponent(z)
                    ActiveCell.Offset(i - 1, 0) = z
                Next i
End Function

Function CellChangeVal( _
    pstrVal As Variant _
    ) As String
    Dim al As String
    If IsNumeric(pstrVal) = True Then
        al = Cells(1, pstrVal).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        CellChangeVal = Left(al, Len(al) - 1)
    Else
        CellChangeVal = Range(pstrVal & "1").Column
    End If
End Function

