Attribute VB_Name = "qlcb_remote"
Sub QLCBRemoteUpload( _
    Optional pstrCom _
    )
    Dim strMediaName As String
    Dim strMediaNames As Variant
    Dim i As Integer
    Dim strSetPutimgPath As String
    
    If SheetCheck(cstrWSName1) = False Then Call QLCBErr001
    If WorksheetFunction.IsError(Range(cstrMediaNameCell)) = True Then Call QLCBErr003
    strMediaName = Range(cstrMediaNameCell)
    strMediaNames = Split(strMediaName, ",")
    For i = 0 To UBound(Split(strMediaName, ","))
        strMediaName = strMediaNames(i)
        If strMediaName <> "" Then
            strSetPutimgPath = FileRead(ThisWorkbook.Path & "\tmpl\mtos\" & strMediaName & "\putimg", 1)
            Call RemoteUpload(RegexBASPReplace("s/^u (.+?)$/$1/g", pstrCom), strSetPutimgPath)
        End If
    Next
    MsgBox _
        "*************" & vbCrLf & _
        " Img->Upload" & vbCrLf & _
        "*************" & vbCrLf & _
        vbCrLf & _
        " Setup complete." & _
        vbCrLf, _
        vbInformation + vbMsgBoxSetForeground, _
        cstrMacroName & " " & cstrMacroVer
End Sub
