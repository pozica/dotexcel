Attribute VB_Name = "qlcb_end"
Sub QLCBEnd( _
    Optional pstrCom _
    )
    Call TimeWait(3000)
    Call FileKill(ActiveWorkbook.Path & "\temp")
    
    Select Case pstrCom
        Case "0111"
            MsgBox _
                "*******************************" & vbCrLf & _
                " Make img + Make file + Import" & vbCrLf & _
                "*******************************" & vbCrLf & _
                vbCrLf & _
                " Setup complete." & _
                vbCrLf, _
                vbInformation + vbMsgBoxSetForeground, _
                cstrMacroName & " " & cstrMacroVer
        Case "0100"
            MsgBox _
                "***********" & vbCrLf & _
                " Make img" & vbCrLf & _
                "***********" & vbCrLf & _
                vbCrLf & _
                " Setup complete." & _
                vbCrLf, _
                vbInformation + vbMsgBoxSetForeground, _
                cstrMacroName & " " & cstrMacroVer
        Case "0010"
            MsgBox _
                "***********" & vbCrLf & _
                " Make file" & vbCrLf & _
                "***********" & vbCrLf & _
                vbCrLf & _
                " Setup complete." & _
                vbCrLf, _
                vbInformation + vbMsgBoxSetForeground, _
                cstrMacroName & " " & cstrMacroVer
        Case "0110"
            MsgBox _
                "**********************" & vbCrLf & _
                " Make img + Make file" & vbCrLf & _
                "**********************" & vbCrLf & _
                vbCrLf & _
                " Setup complete." & _
                vbCrLf, _
                vbInformation + vbMsgBoxSetForeground, _
                cstrMacroName & " " & cstrMacroVer
        Case "0101"
            MsgBox _
                "*******************" & vbCrLf & _
                " Make img + Import" & vbCrLf & _
                "*******************" & vbCrLf & _
                vbCrLf & _
                " Setup complete." & _
                vbCrLf, _
                vbInformation + vbMsgBoxSetForeground, _
                cstrMacroName & " " & cstrMacroVer
        Case "0011"
            MsgBox _
                "********************" & vbCrLf & _
                " Make file + Import" & vbCrLf & _
                "********************" & vbCrLf & _
                vbCrLf & _
                " Setup complete." & _
                vbCrLf, _
                vbInformation + vbMsgBoxSetForeground, _
                cstrMacroName & " " & cstrMacroVer
    End Select
    Exit Sub
End Sub



