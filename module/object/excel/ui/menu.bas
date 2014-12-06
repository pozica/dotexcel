Attribute VB_Name = "object_excel_ui_menu"
Sub MenuHide()
    With Application
        .DisplayFullScreen = True
        .DisplayStatusBar = False
        .DisplayFormulaBar = False
        .CommandBars("Worksheet Menu Bar").Enabled = False
        .CommandBars("Formatting").Visible = False
        .CommandBars("Picture").Visible = False
        .CommandBars("Drawing").Visible = False
    End With
End Sub

Sub MenuHide2()
    With Application
        .DisplayFullScreen = True
        .DisplayStatusBar = False
        .DisplayFormulaBar = True
        .CommandBars("Worksheet Menu Bar").Enabled = False
        .CommandBars("Formatting").Visible = False
        .CommandBars("Picture").Visible = False
        .CommandBars("Drawing").Visible = False
    End With
End Sub

Sub MenuUnhide()
    With Application
        .DisplayFullScreen = False
        .DisplayStatusBar = True
        .DisplayFormulaBar = True
        .CommandBars("Worksheet Menu Bar").Enabled = True
        .CommandBars("Formatting").Visible = True
        .CommandBars("Picture").Visible = True
        .CommandBars("Drawing").Visible = True
    End With
End Sub

