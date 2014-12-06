Attribute VB_Name = "other_bfile"
Sub BFileReadMain()
    Dim strDir As String
    Dim strData As String
    Dim varData As Variant
    Dim strKey1 As String
    Dim strKey2 As String
    Dim strKey3 As String
    
    strKey1 = "ÅyÉLÅ[ÉèÅ[ÉhÅz"
    strKey2 = "ÅyãLéñÉ^ÉCÉgÉãÅz"
    strKey3 = "ÅyãLéññ{ï∂Åz"
    strDir = ActiveWorkbook.Path & "\..\PS_AppsForASP\"
    strData = FileList2(strDir)
    varData = Split(strData, vbTab)
    For i = 0 To UBound(varData)
        Call BFileRead( _
            varData(i), _
            strKey1, _
            strKey2, _
            strKey3 _
        )
    Next i
End Sub

Function BFileRead( _
    pstrFile, _
    strKey1, _
    strKey2, _
    strKey3 _
    )
    Dim TextLine As String
    Dim strData1 As String
    Dim strData2 As String
    Dim strData3 As String
    Dim flg As Integer
    
    R = ActiveCell.Row
    c = ActiveCell.Column
    flg = 0
    Open pstrFile For Input As #1
        Do While Not EOF(1)
            Line Input #1, TextLine
            If InStr(1, TextLine, strKey1) <> 0 Then
                flg = 1
            ElseIf InStr(1, TextLine, strKey2) <> 0 Then
                flg = 2
            ElseIf InStr(1, TextLine, strKey3) <> 0 Then
                flg = 3
            ElseIf flg = 1 Then
                strData1 = strData1 & TextLine
            ElseIf flg = 2 Then
                strData2 = strData2 & TextLine
            ElseIf flg = 3 Then
                strData3 = strData3 & TextLine
            End If
        Loop
    Close #1
    Cells(R, c) = replace(replace(strData1, vbCr, ""), vbLf, "")
    Cells(R, c + 1) = replace(replace(strData2, vbCr, ""), vbLf, "")
    Cells(R, c + 2) = replace(replace(strData3, vbCr, ""), vbLf, "")
    Call ActiveCell.Offset(1, 0).Select
End Function

