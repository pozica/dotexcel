Attribute VB_Name = "object_excel_ui_initialize"
Sub Initialize()
    Call BookMaximized
    Call Keybind
'    Call QLCBHereDoc
    If Application.CommandBars("Worksheet Menu Bar").Visible Then Call MenuHide2 Else Call MenuUnhide
End Sub

Sub InitializeDisplay()
    Dim x As String
    Dim y As String
    Application.DisplayAlerts = False
        x = ActiveSheet.Name
        y = ActiveCell.Address
        SheetSelectAll
            If ActiveWindow.DisplayWorkbookTabs Then Call DisplayNone Else Call DisplayNone2
            Call DisplayZoom
        Sheets(x).Select
        Range(y).Select
    Application.DisplayAlerts = True
End Sub

Sub InitializeDisplay2()
    Dim x As String
    Dim y As String
    Application.DisplayAlerts = False
        x = ActiveSheet.Name
        y = ActiveCell.Address
        SheetSelectAll
            If ActiveWindow.DisplayWorkbookTabs Then Call DisplayNone Else Call DisplayAll
            Call DisplayZoom
        Sheets(x).Select
        Range(y).Select
    Application.DisplayAlerts = True
End Sub

Sub InitializeSheet()
    Application.DisplayAlerts = False
    Call Sheets.Add
    ActiveWindow.Zoom = 75
    With ActiveWindow
        .DisplayHeadings = False
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = False
    End With
    Call Cells.Select
    With Selection.Font
        .Name = "Arial"
        .FontStyle = "ïWèÄ"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Call ActiveSheet.Next.Select
    Call Cells.Select
    Call Selection.Copy
    Call ActiveSheet.Previous.Select
    Call Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Call Range("A1").Select
    Call ActiveSheet.Next.Select
    Call ActiveWindow.SelectedSheets.Delete
    Application.DisplayAlerts = True
End Sub

Sub InitializeFont()
    Dim strR As String
    Dim x As String
    Dim y As String
    
    On Error Resume Next
    Application.DisplayAlerts = False
        
        With Cells.Font
            .Name = "ÇlÇr ÉSÉVÉbÉN"
            .FontStyle = "ïWèÄ"
            .Size = 12
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
        End With
        
        With Cells
            .HorizontalAlignment = xlGeneral
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
    Application.DisplayAlerts = True
    
    
    On Error GoTo 0
End Sub


