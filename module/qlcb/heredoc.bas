Attribute VB_Name = "qlcb_heredoc"
Public Sub QLCBHereDoc()
    hdcQLCB = HereDoc("__CommandQLCB__", "mdl10_QLCBAlias")
    hdcQLCBMTOS = HereDoc("__CommandQLCBMTOS__", "mdl10_QLCBAlias")
End Sub
