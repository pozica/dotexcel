Attribute VB_Name = "object_format"
Option Explicit

Function FormatFigure(pstrCell)
    Dim strReturn
    
    strReturn = format(pstrCell, "#,##0")
    
    FormatFigure = strReturn
End Function

Function FormatTime(pstrCell)
    Dim strReturn
    
    strReturn = format(pstrCell, "hFmm")
    
    FormatTime = strReturn
End Function

Function FormatDate(pstrCell)
    Dim strReturn
    
    strReturn = format(pstrCell, "mŒd“ú;@")
    
    FormatDate = strReturn
End Function

Function FormatNow()
    Dim strNow
    
    strNow = format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    
    FormatNow = strNow
End Function

Function FormatF(pstrCell, pstrOption)
    Dim rc
    
    rc = format(pstrCell, pstrOption)
    
    FormatF = rc
End Function

Function FormatR( _
    pstrData, _
    pstrFormat _
    )
    Dim rc
    
    rc = format(pstrData, pstrFormat)
    rc = format(rc, "@")
    rc = RegexBASPReplace("s/ŒßŒã/PM/g", rc)
    rc = RegexBASPReplace("s/Œß‘O/AM/g", rc)
    
    FormatR = rc
End Function

Function FormatURLEnc( _
    pstrCell _
    )
    Dim sc
    Dim js
    Set sc = CreateObject("ScriptControl")
    sc.Language = "Jscript"
    Set js = sc.CodeObject
    
    FormatURLEnc = js.encodeURIComponent(pstrCell.Text)
End Function

Public Function FormatToCharCode( _
    pstrInput _
    ) As Integer
    FormatToCharCode = pstrInput.charCodeAt(0)
End Function

Function FormatString( _
    •¶š—ñ As String, _
    Optional ”¼Šp As Boolean = True, _
    Optional ‘SŠp As Boolean = False, _
    Optional ’uŠ·ƒV[ƒg As String = "relay", _
    Optional –@lŠi•ªŠ„ As Boolean = False _
    )
    Dim ReplaceList As String
    Dim TargetStr As String, TargetStr2 As String, TargetStr3 As String
    Dim WIDE As String, NARROW As String, ALB As String
    Dim i As Long, j As Long
    
    WIDE = "I“•ij–{FGHobpBA"
    NARROW = "‚O‚P‚Q‚R‚S‚T‚U‚V‚W‚X‚O|D—C,@^Q”ƒ„mnúVO"
    ALB = "‚`‚a‚b‚c‚d‚e‚f‚g‚h‚i‚j‚k‚l‚m‚n‚o‚p‚q‚r‚s‚t‚u‚v‚w‚x‚y"
    ReplaceList = NARROW & ALB & StrConv(ALB, vbLowerCase)
    If ”¼Šp Then ReplaceList = ReplaceList & NARROW
    If ‘SŠp Then ReplaceList = ReplaceList & WIDE
    •¶š—ñ = StrConv(•¶š—ñ, vbWide)
    
    For i = 1 To Len(ReplaceList)
        TargetStr = Mid(ReplaceList, i, 1)
        •¶š—ñ = replace(•¶š—ñ, TargetStr, StrConv(TargetStr, vbNarrow))
    Next i
        
    If SheetCheck(’uŠ·ƒV[ƒg) Then
        If –@lŠi•ªŠ„ Then
            j = 2
            Do Until ActiveWorkbook.Sheets(’uŠ·ƒV[ƒg).Cells(j, 1) = ""
                TargetStr2 = ActiveWorkbook.Sheets(’uŠ·ƒV[ƒg).Cells(j, 1)
                TargetStr3 = ";" & ActiveWorkbook.Sheets(’uŠ·ƒV[ƒg).Cells(j, 2) & ";"
                •¶š—ñ = replace(•¶š—ñ, TargetStr2, TargetStr3)
                j = j + 1
            Loop
        Else
            j = 2
            Do Until ActiveWorkbook.Sheets(’uŠ·ƒV[ƒg).Cells(j, 1) = ""
                TargetStr2 = ActiveWorkbook.Sheets(’uŠ·ƒV[ƒg).Cells(j, 1)
                TargetStr3 = ActiveWorkbook.Sheets(’uŠ·ƒV[ƒg).Cells(j, 2)
                •¶š—ñ = replace(•¶š—ñ, TargetStr2, TargetStr3)
                j = j + 1
            Loop
        End If
    End If
    
    FormatString = •¶š—ñ
End Function

Function FormatString2( _
    •¶š—ñ As String, _
    Optional ”¼Šp As Boolean = True, _
    Optional ‘SŠp As Boolean = False _
    )
    Dim ReplaceList As String
    Dim TargetStr As String
    Dim WIDE As String, NARROW As String, ALB As String
    Dim i As Long
    
    WIDE = "I“•úVij–{FGHOobpBA"
    NARROW = "‚O‚P‚Q‚R‚S‚T‚U‚V‚W‚X‚O|D—C,@^Q”ƒ„mn"
    ALB = "‚`‚a‚b‚c‚d‚e‚f‚g‚h‚i‚j‚k‚l‚m‚n‚o‚p‚q‚r‚s‚t‚u‚v‚w‚x‚y"
    ReplaceList = NARROW & ALB & StrConv(ALB, vbLowerCase)
    If ”¼Šp Then ReplaceList = ReplaceList & NARROW
    If ‘SŠp Then ReplaceList = ReplaceList & WIDE
    
    •¶š—ñ = StrConv(•¶š—ñ, vbWide)
    For i = 1 To Len(ReplaceList)
        TargetStr = Mid(ReplaceList, i, 1)
        •¶š—ñ = replace(•¶š—ñ, TargetStr, StrConv(TargetStr, vbNarrow))
    Next i
    
    FormatString2 = •¶š—ñ
End Function

Sub FormatString3()
    On Error Resume Next
        Dim rngCell As Range
        Application.ScreenUpdating = False
        For Each rngCell In Selection
            If rngCell.Value <> "" Then
            rngCell.Value = FormatString2(rngCell.Value, True, False)
            End If
        Next rngCell
        Application.ScreenUpdating = True
End Sub

Function FormatStringSelection( _
    Optional pstr’uŠ·ƒV[ƒg As String = cstrWSName2, _
    Optional pbln–@lŠi•ªŠ„ As Boolean = False _
    )
    Dim rngCell As Range
    Dim strSetRelayPath As String
    
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
        strSetRelayPath = ThisWorkbook.Path & "\tmpl\bstyle\le.tsv"
        Call FileAddData(strSetRelayPath, cstrTmpDataCell & 1, pstr’uŠ·ƒV[ƒg)
        
        For Each rngCell In Selection
            If rngCell.Value <> "" Then
                rngCell.Value = FormatString(rngCell.Value, True, False, pstr’uŠ·ƒV[ƒg, pbln–@lŠi•ªŠ„)
            End If
        Next rngCell
    
    Application.DisplayAlerts = False
    Worksheets(cstrWSName2).Delete
    Application.ScreenUpdating = True

ErrorHandler:
    Application.DisplayAlerts = False
    Worksheets(cstrWSName2).Delete
End Function
