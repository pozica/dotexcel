Attribute VB_Name = "qlcb_cell"
Function QLCBCellGet( _
    pstrCell, _
    pstrStartCellNo, _
    pstrEndCellNo _
    ) As String
    Dim rc
    Dim cnt
    
    For cnt = pstrStartCellNo To pstrEndCellNo
        rc = pstrCell & cnt
        rc = Application.ConvertFormula(rc, xlR1C1, xlA1)
        QLCBCellGet = QLCBCellGet & "," & Range(cstrWSName1 & "!" & rc)
    Next
    
    QLCBCellGet = RegexBASPReplace("s/^,(.+?)$/$1/g", QLCBCellGet)
End Function

Function QLCBCellMove( _
    pstrCell, _
    pstrX, _
    pstrY _
    )
    Dim rc1
    Dim rc2
    
    pstrCell = RegexBASPReplace("s/" & cstrWSName1 & "!//g", pstrCell)
    pstrCell = Application.ConvertFormula(pstrCell, xlA1, xlR1C1)
    rc1 = RegexBASPReplace("s/R\d+?C(\d+?)$/$1/g", pstrCell)
    rc2 = RegexBASPReplace("s/R(\d+?)C\d+?$/$1/g", pstrCell)
    rc1 = rc1 + pstrX
    rc2 = rc2 + pstrY
    pstrCell = "R" & rc2 & "C" & rc1
    
    QLCBCellMove = cstrWSName1 & "!" & Application.ConvertFormula(pstrCell, xlR1C1, xlA1)
End Function

Function QLCBconCells( _
    pstrExtent As Range, _
    Optional pstrSV = "" _
    ) As Variant
    Dim rc1
    Dim rc2
    
    If pstrExtent.Rows.count = 1 Or pstrExtent.Columns.count = 1 Then
        For Each rc1 In pstrExtent
            rc2 = rc2 & pstrSV & rc1.Value
        Next rc1
        If pstrSV <> "" Then
            myJoin = Mid$(rc2, 2)
        Else
            myJoin = rc1
        End If
    Else
        QLCBconCells = CVErr(xlErrRef)
    End If
End Function

Sub QLCBCellConvertCSV2Col( _
    pstrData, _
    pstrCells _
    )
    Dim rc1
    Dim rc2
    Dim rc3
    Dim cnt1
    Dim cnt2
    
    rc1 = Split(pstrData, ",")
    cnt1 = UBound(Split(pstrData, ","))
    rc2 = rc1(0)
    Range(pstrCells) = rc2
    rc3 = pstrCells
    For cnt2 = 1 To cnt1
        rc2 = rc1(cnt2)
        rc3 = QLCBCellMove(rc3, 1, 0)
        Range(rc3) = rc2
    Next
End Sub



