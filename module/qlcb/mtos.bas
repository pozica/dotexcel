Attribute VB_Name = "qlcb_mtos"
Sub QLCBMTOSEntry( _
    pstrFromDir, _
    pstrMTEntryPath, _
    pstrCond _
    )
    Call RemoteUpload(pstrFromDir & "\*", "entsht\006\import")
    Call Shell(pstrMTEntryPath, vbHide)
    If MsgBox( _
        "*************" & vbCrLf & _
        " Import MTOS" & vbCrLf & _
        "*************" & vbCrLf & _
        vbCrLf & _
        " If you import MTOS data," & vbCrLf & _
        " then you should input [OK button]." & vbCrLf & _
        vbCrLf & _
        "* Deleted temporary files of server" & vbCrLf & _
        "  and process images." & vbCrLf & _
        "* In case of bugs" & vbCrLf & _
        "  then you should check Excel zombies" & vbCrLf & _
        "  in the process manager." & _
        vbCrLf & _
        vbCrLf, _
        vbExclamation + vbMsgBoxSetForeground, _
        cstrMacroName & " " & cstrMacroVer _
        ) = 1 Then
        Call RemoteDeleteFile("entsht\006\import\*")
    End If
End Sub

Function QLCBMTOSGetEntryIDs( _
    ) As Variant
    
    If SheetCheck(cstrWSName1) = True Then
        QLCBMTOSGetEntryIDs = QLCBCellGet(cstrEntryIDCell1, cstrEntryIDCell2, cstrEntryIDCell3)
        If RegexBASPMatch("/invalid/", QLCBMTOSGetEntryIDs) <> 0 Then QLCBMTOSGetEntryIDs = ""
    Else
        QLCBMTOSGetEntryIDs = ""
    End If
End Function

Sub QLCBMTOSFillEntryID()
    If SheetCheck(cstrWSName1) = False Then Call QLCBErr001
    If WorksheetFunction.IsError(Range(cstrMediaNameCell)) = True Then Call QLCBErr003
    Dim i As Integer
    Dim x As String
    Dim y As String
    Dim z As String
    Dim strMediaName As String
    Dim strMediaNames As Variant
    Dim strEntryIDs As Variant
    
    strEntryIDs = QLCBMTOSGetEntryIDs
    strMediaName = Range(cstrMediaNameCell)
    strMediaNames = Split(strMediaName, ",")
    For i = 0 To UBound(Split(strMediaName, ","))
        strMediaName = strMediaNames(i)
        If strMediaName <> "" Then
            x = x & "" & strMediaName & vbCrLf
        Else
            x = x & "" & "N/A" & vbCrLf
        End If
    Next
    y = InputBox( _
        "**********" & vbCrLf & _
        " $ID->Set" & vbCrLf & _
        "**********" & vbCrLf & _
        vbCrLf & _
        "Setup" & vbCrLf & _
        vbCrLf & _
        "* You should input orderly media codes in case of a lot of media" & vbCrLf & _
        "* You may arrange orderly it in case of entrying media." & vbCrLf & _
        vbCrLf & _
        "Media code:" & vbCrLf & _
        "-----------" & vbCrLf & _
        x & vbCrLf & _
        "", _
        cstrMacroName & " " & cstrMacroVer, strEntryIDs _
        )
    If y <> "" Then
        z = Application.ConvertFormula(cstrEntryIDCell1 & cstrEntryIDCell2, xlR1C1, xlA1)
        z = cstrWSName1 & "!" & z
        Call QLCBCellConvertCSV2Col(y, z)
    End If
End Sub

Sub QLCBMTOSFillEntryBasename()
    If SheetCheck(cstrWSName1) = False Then Call QLCBErr001
    Dim x As String
    Dim y As String
    
    x = Range(cstrEntryBasenameCell)
    y = InputBox( _
        "****************" & vbCrLf & _
        " $Basename->Set" & vbCrLf & _
        "****************" & vbCrLf & _
        vbCrLf & _
        "Setup" & vbCrLf & _
        "", _
        cstrMacroName & " " & cstrMacroVer, Range(cstrEntryBasenameCell) _
        )
    If y <> "" Then Range(cstrEntryBasenameCell) = y
End Sub

Sub QLCBMTOSFillEntryDate()
    If SheetCheck(cstrWSName1) = False Then Call QLCBErr001
    Dim x As String
    Dim y As String
    
    x = Range(cstrEntryDateCell)
    y = InputBox( _
        "************" & vbCrLf & _
        " $Date->Set" & vbCrLf & _
        "************" & vbCrLf & _
        vbCrLf & _
        "Setup" & vbCrLf & _
        "", _
        cstrMacroName & " " & cstrMacroVer, x _
        )
    y = "=""" & format(y, "@") & """"
    If y <> "" Then Range(cstrEntryDateCell) = y
End Sub

Sub QLCBMTOSFillEntryDateNow()
    If SheetCheck(cstrWSName1) = False Then Call QLCBErr001
    Dim x As String
    Dim y As String
    
    x = FormatR(Now(), "m/d/yyyy HH:mm:ss AMPM")
    y = InputBox( _
        "**************" & vbCrLf & _
        " $Date->Now()" & vbCrLf & _
        "**************" & vbCrLf & _
        vbCrLf & _
        "Setup" & vbCrLf & _
        "", _
        cstrMacroName & " " & cstrMacroVer, x _
        )
    If y <> "" Then Range(cstrEntryDateCell) = "=""" & y & """"
End Sub

