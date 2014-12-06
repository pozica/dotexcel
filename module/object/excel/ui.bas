Attribute VB_Name = "object_excel_ui"
Sub DisplayZoom( _
    Optional pintNum _
    )
    If IsNumeric(pintNum) = False Then pintNum = 75
    ActiveWindow.Zoom = pintNum
End Sub

Sub DisplayNone()
    With ActiveWindow
        .DisplayGridlines = False
        .DisplayHeadings = False
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = False
        .DisplayWorkbookTabs = False
    End With
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
End Sub

Sub DisplayNone2()
    With ActiveWindow
        .DisplayGridlines = False
        .DisplayHeadings = False
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = False
        .DisplayWorkbookTabs = True
    End With
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = False
End Sub

Sub DisplayAll()
    With ActiveWindow
        .DisplayGridlines = True
        .DisplayHeadings = True
        .DisplayWorkbookTabs = True
    End With
    Application.DisplayStatusBar = True
End Sub

Sub DisplayRefresh()
    ActiveWorkbook.PrecisionAsDisplayed = False
    ActiveSheet.Calculate
    AnalPivotRefresh
End Sub


