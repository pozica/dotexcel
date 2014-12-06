Attribute VB_Name = "qlcb_display"
Function QLCBDisplay( _
    Optional pstrCom, _
    Optional pstrCommand _
    )
    Dim x As String
    Dim y As String
    Dim ws As Worksheet
    Dim bln As Boolean
    
    pstrCom = InputBox( _
        pstrCommand & vbCrLf & _
        vbCrLf & y, _
        cstrMacroName & " " & cstrMacroVer, pstrCom _
        )
    QLCBDisplay = pstrCom
End Function

Function QLCBDisplayMTOS()
    Dim i As Integer
    Dim y As String
    Dim strMediaName As String
    Dim strMediaNames As Variant
    Dim strEntryIDs As Variant
    
    strEntryIDs = QLCBMTOSGetEntryIDs

    y = "Settings:" & vbCrLf
    y = y & "---------" & vbCrLf
    If SheetCheck(cstrWSName1) = True Then
        y = y & "Initialize    :Complete" & vbCrLf
        If WorksheetFunction.IsError(Range(cstrMediaNameCell)) = False Then
            y = y & "Media code:" & vbCrLf
            strMediaName = Range(cstrMediaNameCell)
            strMediaNames = Split(strMediaName, ",")
            For i = 0 To UBound(Split(strMediaName, ","))
                strMediaName = strMediaNames(i)
                If strMediaName <> "" Then
                    y = y & "      " & strMediaName & vbCrLf
                Else
                    y = y & "      " & "N/A" & vbCrLf
                End If
            Next
        Else
            y = y & "Media code: Error" & vbCrLf
        End If
        y = y & "$ID  :" & strEntryIDs & vbCrLf
        If Range(cstrEntryBasenameCell) <> "" Then
            y = y & "$Basename :" & Range(cstrEntryBasenameCell) & vbCrLf
        Else
            y = y & "$Basename :NULL" & vbCrLf
        End If
        y = y & "$Date     :" & Range(cstrEntryDateCell) & vbCrLf
    Else
        y = y & "Initialize    :Not complete" & vbCrLf
    End If
    QLCBDisplayMTOS = y
End Function

Function QLCBDisplayMTOSImages()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim x As String
    Dim z As String
    Dim strMediaName As String
    Dim strMediaNames As Variant
    Dim arrImageNameCell As String
    Dim arrImageName As Variant
    Dim arrImageNames As Variant
    Dim strImageName As String
    Dim strImageNames As Variant
    
    If SheetCheck(cstrWSName1) = True Then
        If WorksheetFunction.IsError(Range(cstrMediaNameCell)) = False Then strMediaName = Range(cstrMediaNameCell)
        If strMediaName <> "" Then
            x = ""
            z = ""
            strMediaNames = Split(strMediaName, ",")
            For i = 0 To UBound(Split(strMediaName, ","))
                If strMediaName <> "" Then
                    x = x & "  * " & strMediaNames(i) & "" & vbCrLf
                    k = carrImageNameCell2 + i
                    arrImageNameCell = carrImageNameCell1 & k
                    arrImageNameCell = Application.ConvertFormula(arrImageNameCell, xlR1C1, xlA1)
                    arrImageName = Range(arrImageNameCell)
                    arrImageNames = Split(arrImageName, ";")
                    For j = 0 To UBound(Split(arrImageName, ";"))
                        strImageName = arrImageNames(j)
                        strImageNames = Split(strImageName, ",")
                        If RegexBASPMatch("/!/", strImageNames(1)) + RegexBASPMatch("/3/", strImageNames(0)) = 0 Then z = z & "    - " & strImageNames(1) & vbCrLf
                    Next
                    If UBound(Split(arrImageName, ";")) < 21 Then
                        x = z
                    Else
                        x = "    Over 20 setting images" & vbCrLf
                    End If
                End If
            Next
        End If
    End If
    QLCBDisplayMTOSImages = x
End Function

