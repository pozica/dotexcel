Attribute VB_Name = "qlcb_mtos_text"
Sub QLCBMTOSText( _
    Optional pstrCom _
    )
    Dim strTmpFilePath As String
    Dim strTmpFileFullName As String
    Dim strTmpData As String
    Dim strMediaName As String
    Dim strMediaNames As Variant
    Dim intMediaNamesTimes As Integer
    Dim strEntryBasename As String
    Dim strSetRelayPath As String
    Dim strNow As String
    
    strNow = TimeGetDate(1)
    strTmpFilePath = ActiveWorkbook.Path & "\_tmp"
    If SheetCheck(cstrWSName1) = True Then strEntryBasename = Range(cstrEntryBasenameCell)
    If WorksheetFunction.IsError(Range(cstrMediaNameCell)) = True Then Call QLCBErr003
    strMediaName = Range(cstrMediaNameCell)
    strMediaNames = Split(strMediaName, ",")
    intMediaNamesTimes = UBound(Split(strMediaName, ","))
    For k = 0 To intMediaNamesTimes
        strMediaName = strMediaNames(k)
        If strMediaName <> "" Then
            strTmpFilePath = strTmpFilePath & "\" & "mt_" & strMediaName & "_" & strEntryBasename & "_" & strNow
            Call FileCheck2(strTmpFilePath)
            strSetRelayPath = ThisWorkbook.Path & "\tmpl\mtos\" & strMediaName & "\relay.tsv"
            
            '''' B-2-1. Get data
            If FileAddData(strSetRelayPath, cstrTmpDataCell & 1, cstrWSName2, pstrCom) = 3 Then Exit Sub
            For i = 1 To Cells(Rows.count, 1).End(xlUp).Row
                strTmpFileFullName = strTmpFilePath & "\" & cstrTmpFileName & i & ".txt"
                strTmpData = Range(cstrTmpDataCell & i)
                If strTmpData <> "" Then
                    If strTmpData <> "" Then Call FileMake("Utf-8", strTmpData, strTmpFileFullName)
                End If
            Next i
            Application.DisplayAlerts = False
            Worksheets(cstrWSName2).Delete
            strTmpFilePath = ActiveWorkbook.Path & "\_tmp"
        End If
    Next
End Sub

Sub QLCBMTOSTextUpload()
    If WorksheetFunction.IsError(Range(cstrMediaNameCell)) = True Then Call QLCBErr003
    Dim strSetImportPath As String
    
    strSetImportPath = ThisWorkbook.Path & "\tmpl\mtos\" & Range(cstrMediaNameCell) & "\import.bat"
    If FileCheck(ActiveWorkbook.Path & "\_tmp\" & cstrTmpFileName & 1 & ".txt") = False Then
        MsgBox _
            "Error: Not exist " & _
            & ActiveWorkbook.Path & "\_tmp\" & cstrTmpFileName & 1 & ".txt", _
            vbExclamation
    Else
        Call TimeWait(1500)
        Call QLCBMTOSEntry(ActiveWorkbook.Path & "\_tmp", strSetImportPath, intActionStat2)
    End If
End Sub
