Attribute VB_Name = "util_image"
Sub ImageResize( _
    pstrInputFile, _
    pstrOutputFile, _
    pstrResize, _
    pstrFormat, _
    pstrPath _
    )
    Call TimeWait(100)
    Call Shell( _
        "cmd /c pushd " & pstrPath & _
        " & convert" & _
        " " & pstrInputFile & _
        " -resize " & pstrResize & "^^" & _
        " -gravity center" & _
        " -format " & pstrFormat & _
        " -extent " & pstrResize & _
        " " & pstrOutputFile _
        )
    Call TimeWait(1000)
End Sub

Sub ImageFromText( _
    pstrLabel, _
    pstrOutputName, _
    pstrFont, _
    pstrOther, _
    pstrPath _
    )
    Call TimeWait(4000)
    Call Shell( _
        "cmd /c pushd " & pstrPath & _
        " & convert" & _
        " -font " & pstrFont & _
        " " & pstrOther & _
        " label:@" & pstrLabel & _
        " " & pstrOutputName, vbNormalFocus _
        )
    Call TimeWait(1000)
End Sub

Sub ImageFromTextSelection()
    Dim i As Integer
    Dim strImageInputName As String
    Dim strImageOutputName As String
    Dim strImageFirst As String
    Dim strImageSecond As String
    
    For i = 0 To Selection.count - 1
        If ActiveCell.Value <> "" Then
            Call FileMake("UTF-8N", ActiveCell.Value, ActiveWorkbook.Path & "\temp")
            strImageInputName = "temp"
            Selection.Offset(0, 1).Select
            strImageOutputName = ActiveWorkbook.Path & "\" & ActiveCell.Value
            Selection.Offset(0, 1).Select
            strImageFirst = ActiveWorkbook.Path & "\" & ActiveCell.Value
            Selection.Offset(0, 1).Select
            strImageSecond = ActiveCell.Value
            Call ImageFromText(strImageInputName, strImageOutputName, strImageFirst, strImageSecond, ActiveWorkbook.Path)
            Selection.Offset(1, -3).Select
        Else
            Selection.Offset(1, 0).Select
        End If
    Next
End Sub

Sub ImageCopy( _
    pstrInputFileName, _
    pstrOutputName, _
    pstrPath _
    )
    Call TimeWait(100)
    Call Shell( _
        "cmd /c pushd " & pstrPath & _
        " & copy" & _
        " " & pstrInputFileName & _
        " " & pstrOutputName, vbNormalFocus _
        )
    Call TimeWait(100)
End Sub


