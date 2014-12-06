Attribute VB_Name = "qlcb_error"
Sub QLCBErr001()
    MsgBox _
        "************" & vbCrLf & _
        " Error M001" & vbCrLf & _
        "************" & vbCrLf & _
        vbCrLf & _
        "You should initialize it." & vbCrLf & _
        vbCrLf & _
        "* You can initialize it with commands [1000-1999]" & vbCrLf & _
        "", _
        vbCritical + vbMsgBoxSetForeground, _
        cstrMacroName & " " & cstrMacroVer
End Sub

Sub QLCBErr003()
    MsgBox _
        "************" & vbCrLf & _
        " Error M003" & vbCrLf & _
        "************" & vbCrLf & _
        vbCrLf & _
        "You should validate existense of a media code " & vbCrLf & _
        "in cell(" & cstrMediaNameCell & ")." & vbCrLf & _
        vbCrLf & _
        "* Error:" & vbCrLf & _
        "  1. Format error of registration sheet" & vbCrLf & _
        "     => You should fix failures." & vbCrLf & _
        "  2. Failure of WS function in cell(" & cstrMediaNameCell & ")" & vbCrLf & _
        "     => You should fix failures with command [2000-2999]." & vbCrLf & _
        vbCrLf & _
        "", _
        vbCritical + vbMsgBoxSetForeground, _
        cstrMacroName & " " & cstrMacroVer
End Sub



