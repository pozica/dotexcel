Attribute VB_Name = "object_match"
Function MatchVlookup( _
    pstrQuery As String, _
    pstrSearches As String, _
    pstrNum As Long _
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
            If IsError(Application.WorksheetFunction.VLookup(Range(pstrQuery & y(i)), Range(pstrSearches), pstrNum, False)) Then
                z = ""
            Else
                z = Application.WorksheetFunction.VLookup(Range(pstrQuery & y(i)), Range(pstrSearches), pstrNum, False)
            End If
            ActiveCell.Offset(i - 1, 0) = z
        Next i

End Function

Function MatchVlookup1_1( _
    pstrQuery As String, _
    pstrSearchesSheet As String _
    )
    Dim rng As Range
    Dim iCol1 As Long
    Dim iCol2 As Long
    Dim iRow As Long
    Dim colCnt As Long
    Dim colCnt1000 As Long
    Dim colCntMod As Long
    Dim rowCnt As Long
    Dim x As String
    Dim y As Variant
    Dim z As String
    Dim ws As Worksheet
    
    For Each rng In Selection
        x = x & "," & rng.Row
    Next
    y = Split(x, ",")
    
    colCnt = Selection.Columns.count
    colCnt1000 = Selection.Columns.count \ 1000
    colCntMod = colCnt - (colCnt1000 * 1000)
    rowCnt = UBound(y) / colCnt
    
    Set ws = Worksheets(pstrSearchesSheet)
        On Error Resume Next
            For iCol2 = 1 To colCnt1000
                For iCol1 = 1 To colCnt1000 * 1000
                    For iRow = 1 To rowCnt
                        If IsError(Application.WorksheetFunction.VLookup( _
                            Range(pstrQuery & y(iRow * colCnt)), _
                            ws.Range(ws.Cells(1, 1), ws.Cells(65535, 1 + colCnt)), _
                            1 + iCol1, _
                            False) _
                        ) Then
                            z = ""
                        Else
                            z = _
                                Application.WorksheetFunction.VLookup( _
                                Range(pstrQuery & y(iRow * colCnt)), _
                                ws.Range(ws.Cells(1, 1), ws.Cells(65535, 1 + colCnt)), _
                                1 + iCol1, _
                                False)
                        End If
                        ActiveCell.Offset(iRow - 1, iCol1 - 1) = z
                    Next iRow
                Next iCol1
            Next iCol2
            For iCol1 = 1 To colCntMod
                For iRow = 1 To rowCnt
                    If IsError(Application.WorksheetFunction.VLookup( _
                        Range(pstrQuery & y(iRow * colCnt)), _
                        ws.Range(ws.Cells(1, 1), ws.Cells(65535, 1 + colCnt)), _
                        1 + iCol1, _
                        False) _
                    ) Then
                        z = ""
                    Else
                        z = _
                            Application.WorksheetFunction.VLookup( _
                            Range(pstrQuery & y(iRow * colCnt)), _
                            ws.Range(ws.Cells(1, 1), ws.Cells(65535, 1 + colCnt)), _
                            1 + iCol1, _
                            False)
                    End If
                    ActiveCell.Offset(iRow - 1, iCol1 - 1) = z
                Next iRow
            Next iCol1
End Function

Function MatchVlookup2( _
    pstrQuery As String, _
    pstrSearches As String, _
    Optional pstrFlg = False _
    )
    Dim strSetRelayPath As String
    Dim rng As Range
    Dim x As String
    Dim y As Variant
    Dim z As String
    Dim i As Long
    On Error GoTo ErrorHandler
        strSetRelayPath = ThisWorkbook.Path & "\tmpl\bstyle\" & pstrSearches & ".tsv"
        Call FileAddData(strSetRelayPath, cstrTmpDataCell & 1, cstrWSName2)
        
        For Each rng In Selection
            x = x & "," & rng.Row
        Next
        y = Split(x, ",")
        On Error Resume Next
            For i = 1 To UBound(Split(x, ","))
                If IsError(Application.WorksheetFunction.VLookup(Range(pstrQuery & y(i)), Range("relay!A:B"), 2, pstrFlg)) Then
                    z = ""
                Else
                    z = Application.WorksheetFunction.VLookup(Range(pstrQuery & y(i)), Range("relay!A:B"), 2, pstrFlg)
                End If
                ActiveCell.Offset(i - 1, 0) = z
            Next i
        
        Application.DisplayAlerts = False
        Worksheets(cstrWSName2).Delete

ErrorHandler:
    Application.DisplayAlerts = False
    Worksheets(cstrWSName2).Delete
End Function

Function MatchHWC( _
    pstrQuery1 As String, _
    pstrQuery2 As String _
    )
    Dim strSetRelayPath As String
    Dim rng As Range
    Dim x As String
    Dim y As Variant
    Dim z1 As String
    Dim z2 As String
    Dim z As String
    Dim i As Long
    On Error GoTo ErrorHandler
        strSetRelayPath = ThisWorkbook.Path & "\tmpl\bstyle\hwc.tsv"
        Call FileAddData(strSetRelayPath, cstrTmpDataCell & 1, cstrWSName2)
        
        For Each rng In Selection
            x = x & "," & rng.Row
        Next
        y = Split(x, ",")
        On Error Resume Next
            For i = 1 To UBound(y)
                z1 = Date - 90
                If Range(pstrQuery1 & y(i)) <= (Date - 90) Then
                    z1 = "-3mo"
                ElseIf Range(pstrQuery1 & y(i)) <= (Date - 180) Then
                    z1 = "-6mo"
                ElseIf Range(pstrQuery1 & y(i)) <= (Date - 270) Then
                    z1 = "-9mo"
                Else
                    z1 = "9mo-"
                End If
                If Range(pstrQuery2 & y(i)) > 30 Then
                    z2 = "30p-"
                ElseIf Range(pstrQuery2 & y(i)) >= 20 Then
                    z2 = "20-30p"
                ElseIf Range(pstrQuery2 & y(i)) >= 13 Then
                    z2 = "13-19p"
                Else
                    z2 = "-13p"
                End If
                z = Application.WorksheetFunction.VLookup(z1 & z2, Range("relay!A:B"), 2, False)
                ActiveCell.Offset(i - 1, 0) = z
            Next i
        
        Application.DisplayAlerts = False
        Worksheets(cstrWSName2).Delete

ErrorHandler:
    Application.DisplayAlerts = False
    Worksheets(cstrWSName2).Delete
End Function

Function MatchDuplicated( _
    pstrQuery As String, _
    pstrSearches As String, _
    pstrFlg As Boolean _
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
            
            If Range(pstrQuery & y(i)) = False Then
                z = Range(pstrSearches & y(i))
            Else
                If pstrFlg = True Then
                    z = ActiveCell.Offset(i - 2, 0) & "," & Range(pstrSearches & y(i))
                Else
                    z = ActiveCell.Offset(i - 2, 0)
                End If
            End If
            ActiveCell.Offset(i - 1, 0) = z
        Next i
End Function

Function MatchDuplicated2( _
    pstrQuery As String, _
    pstrSearches As String, _
    pstrFlg As Boolean _
    )
    Dim rng As Range
    Dim a As Variant
    Dim b As Long
    Dim x As String
    Dim y As Variant
    Dim z As String
    Dim i As Long
    Dim j As Long
    
    a = Split(pstrSearches, "|")
    b = UBound(a)
    For Each rng In Selection
        x = x & "," & rng.Row
    Next
    y = Split(x, ",")
    
    On Error Resume Next
        For i = 1 To UBound(y)
            If i Mod b = 1 Then
                For j = 0 To b
                    If Range(pstrQuery & y(i)) = False Then
                        z = Range(a(j) & y(i))
                    Else
                        If pstrFlg = True Then
                            z = ActiveCell.Offset((i \ (b + 1)) - 2, j) & "," & Range(a(j) & y(i))
                        Else
                            z = ActiveCell.Offset((i \ (b + 1)) - 2, j) + Range(a(j) & y(i))
                        End If
                    End If
                    If b = 0 Then
                        ActiveCell.Offset((i \ (b + 1)) - 1, j) = z
                    Else
                        ActiveCell.Offset((i \ (b + 1)), j) = z
                    End If
                Next j
            End If
        Next i
End Function

Function MatchVlookup3( _
    pstrQueries As String, _
    pstrSearchesSheet As String _
    )
    Dim rng As Range
    Dim a As Variant
    Dim b As Long
    Dim c As Variant
    Dim iRow As Long
    Dim iCol1 As Long
    Dim iCol2 As Long
    Dim x As String
    Dim y As Variant
    Dim z As String
    Dim ws As Worksheet
    
    a = Split(pstrQueries, "|")
    b = UBound(a)
    
    For Each rng In Selection
        x = x & "," & rng.Row
    Next
    y = Split(x, ",")
    
    Set ws = Worksheets(pstrSearchesSheet)
    
    On Error Resume Next
        For iCol1 = 1 To UBound(y)
            For iRow = 0 To b
                If IsError(Application.WorksheetFunction.VLookup( _
                    Range(a(iRow) & y(iCol1)), _
                    ws.Range(ws.Cells(1, 1 + iRow), ws.Cells(65535, 2 + b)), _
                    2 + b - iRow, _
                    False) _
                ) Then
                    z = ""
                Else
                    z = _
                        Application.WorksheetFunction.VLookup( _
                        Range(a(iRow) & y(iCol1)), _
                        ws.Range(ws.Cells(1, 1 + iRow), ws.Cells(65535, 2 + b)), _
                        2 + b - iRow, _
                        False)
                    ActiveCell.Offset(iCol1 - 1, 0) = z
                    Exit For
                End If
            Next iRow
        Next iCol1
End Function

