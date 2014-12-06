Attribute VB_Name = "util_network_remote"
Sub RemoteUpload( _
    pstrFromDir, _
    pstrToDir _
    )
    Call Shell( _
        "cmd /c copy /y" & _
        " " & pstrFromDir & _
        " " & cstrServer & pstrToDir _
        )
End Sub

Sub RemoteDeleteFile( _
    pstrRemoteDir _
    )
    Call Shell( _
        "cmd /c del /q" & _
        " " & cstrServer & pstrRemoteDir _
        )
End Sub

