Attribute VB_Name = "object_excel_book"
Function BookIsOpenCheck( _
    strBookName As String _
    ) As Boolean
    Dim objBook As Workbook
    
    IsBookOpen = False
    
    For Each objBook In Workbooks
        If objBook.Name = strBookName Then
            IsBookOpen = True
            Exit For
        End If
    Next
End Function

Sub BookAdd()
    Dim x As String
    Application.DisplayAlerts = False
        Workbooks.Add
        x = ActiveSheet.Name
        SheetSelectAll
        Call InitializeFont
        Call DisplayNone2
        Call DisplayZoom
        Sheets(x).Select
    Application.DisplayAlerts = True
End Sub

Sub BookTiled()
    Windows.Arrange ArrangeStyle:=xlTiled
End Sub

Sub BookClose()
    ActiveWindow.Close
End Sub

Sub BookMaximized()
    On Error Resume Next
        ActiveWindow.WindowState = xlMaximized
    On Error GoTo 0
End Sub

Sub BookShare()
    Application.CommandBars.FindControl(ID:=2040).Execute
End Sub
