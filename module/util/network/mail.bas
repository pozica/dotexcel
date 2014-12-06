Attribute VB_Name = "util_network_mail"
Sub MailBASPSelection()
    Dim rc
    Dim i As Integer
    Dim strMailDir As String
    Dim strMailTo As String
    Dim strMailFrom As String
    Dim strMailSubject As String
    Dim strMailBody As String
    Dim strNow As String
    
    strNow = TimeGetDate(1)
    With CreateObject("basp21")
        strMailDir = ActiveWorkbook.Path & "\_tmp\mail_" & strNow
        Call FileCheck2(strMailDir)
        For i = 0 To Selection.count - 1
            If ActiveCell.Value <> "" Then
                strMailTo = ActiveCell.Value
                Call Selection.Offset(0, 1).Select
                strMailFrom = ActiveCell.Value
                strMailFrom = strMailFrom & vbTab & "send:bstyle1qaz"
                Call Selection.Offset(0, 1).Select
                strMailSubject = ActiveCell.Value
                Call Selection.Offset(0, 1).Select
                strMailBody = ActiveCell.Value
                Call SendMail(strMailDir, strMailTo, strMailFrom, strMailSubject, strMailBody, "")
                Call Selection.Offset(1, -3).Select
            Else
                Selection.Offset(1, 0).Select
            End If
        Next
        Call .FlushMail("b-style-part.sakura.ne.jp:587", strMailDir, strMailDir & "/log.txt")
    End With
End Sub
