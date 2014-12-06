Attribute VB_Name = "object_excel_ui_link"
Sub LinkOpenSelection()
    Dim rng As Range
    Dim strPath As String
    Dim strURL As String
    strPath = "C:\Progra~2\Google\Chrome\Application\chrome.exe --no-startup-window "
    
    For Each rng In Selection
        strURL = strURL & " " & rng
    Next
    Call Shell(strPath & strURL)
End Sub

Sub LinkDeleteSelection()
    Dim Hlnks As Hyperlink
    Dim strPath As String
    
    For Each Hlnks In Selection.Hyperlinks
        Call Hlnks.Delete
    Next
End Sub

Sub LinkDelete()
    ActiveSheet.Hyperlinks.Delete
End Sub
