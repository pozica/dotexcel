Attribute VB_Name = "util_network_browser"
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function SetForegroundWindow Lib "USER32" (ByVal Hwnd As Long) As Long

Sub IEGo()
    Dim ie As Object
    
    Set ie = IENew("d:\test.html", False)
        MsgBox IEDOM(ie, Array("ID", "ires", "AT", "li", "class", "aaa", "TG", "a", 0)).innerText
        
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub IEWait( _
    ie As Object, _
    Optional time As Integer = 100 _
    )
    Do While ie.Busy = True Or ie.ReadyState <> 4
        DoEvents
    Loop
    Call Sleep(time)
End Sub

Function IENew( _
    home_url As String, _
    Optional flg_visible As Boolean = False _
    ) As Object
    Dim ie As Object
    
    Application.Visible = flg_visible
    Set ie = CreateObject("InternetExplorer.Application")
        Call IEGotoURL(ie, home_url)
        ie.Visible = flg_visible
    Set IENew = ie
End Function

Sub IEGotoURL( _
    ie As Object, _
    url As String _
    )
    Call ie.Navigate(url)
    Call IEWait(ie)
End Sub

Function IEgid( _
    ie As Object, _
    dom_id As String _
    ) As Object
    ' ’FIE‚ÌgetElementById‚Íname‚àŽQÆ‚·‚é
    Set IEgid = ie.document.getElementByID(dom_id)
End Function

Function IEgnm( _
    ie As Object, _
    Optional dom_nm As String, _
    Optional index_num As Integer = 0 _
    ) As Object
    Set IEgnm = ie.document.getElementsByName(dom_nm)(index_num)
End Function

Function IEgtg( _
    parent As Object, _
    tag_name As String, _
    Optional index_num As Integer _
    ) As Object
    Set IEgtg = parent.getElementsByTagName(tag_name)(index_num)
    'If index_num = "" Then
        'Set IEgtg = parent.getElementsByTagName(tag_name)
    'Else
        'Set IEgtg = parent.getElementsByTagName(tag_name)(index_num)
    'End If
End Function

Function IEgat( _
    ie As Object, _
    tag_name As String, _
    at_name As String, _
    at_val As String _
    )
    Dim tg As Object
    
    For Each tg In ie.document.getElementsByTagName(tag_name)
        Select Case at_name
            Case "accesskey": If tg.AccessKey = at_val Then Set IEgat = tg: Exit For
            Case "alt": If tg.Alt = at_val Then Set IEgat = tg: Exit For
            Case "class": If tg.ClassName = at_val Then Set IEgat = tg: Exit For
            Case "id": If tg.ID = at_val Then Set IEgat = tg: Exit For
            Case "name": If tg.Name = at_val Then Set IEgat = tg: Exit For
            Case "type": If tg.Type = at_val Then Set IEgat = tg: Exit For
            Case "value": If tg.Value = at_val Then Set IEgat = tg: Exit For
        End Select
    Next
End Function

Sub IEInput( _
    ie As Object, _
    val As String, _
    flg_elm As String, _
    dom_elm As String, _
    Optional index_num As Integer, _
    Optional at_name As String, _
    Optional at_val As String _
    )
    Select Case flg_elm
        Case "ID": IEgid(ie, dom_elm).Value = val
        Case "NM": IEgnm(ie, dom_elm, index_num).Value = val
        Case "TG": IEgtg(ie.document, dom_elm, index_num).Value = val
        Case "AT": IEgat(ie, dom_elm, at_name, at_val).Value = val
    End Select
    Call Sleep(100)
End Sub

Sub IEClick( _
    ie As Object, _
    flg_elm As String, _
    dom_elm As String, _
    Optional index_num As Integer, _
    Optional at_name As String, _
    Optional at_val As String _
    )
    Select Case flg_elm
        Case "ID": Call IEgid(ie, dom_elm).Click
        Case "NM": Call IEgnm(ie, dom_elm, index_num).Click
        Case "TG": Call IEgtg(ie.document, dom_elm, index_num).Click
        Case "AT": Call IEgat(ie, dom_elm, at_name, at_val).Click
    End Select
    Call IEWait(ie)
End Sub

Sub IEFocus( _
    ie As Object, _
    flg_elm As String, _
    dom_elm As String, _
    Optional index_num As Integer, _
    Optional at_name As String, _
    Optional at_val As String _
    )
    Select Case flg_elm
        Case "ID": Call IEgid(ie, dom_elm).Focus
        Case "NM": Call IEgnm(ie, dom_elm, index_num).Focus
        Case "TG": Call IEgtg(ie.document, dom_elm, index_num).Focus
        Case "AT": Call IEgat(ie, dom_elm, at_name, at_val).Focus
    End Select
    Call IEWait(ie)
End Sub

Function IEDOM( _
    ie As Object, _
    Arr As Variant _
    ) As Object
    Dim parent As Object
    Dim child As Object
    Dim tg As Object
    Dim i As Integer
    Dim flg As Boolean
    
    i = 0
    flg = True
    Set parent = ie.document
    Do While flg = True
        Select Case Arr(i)
            Case "ID"
                Set child = parent.getElementByID(Arr(i + 1))
                i = i + 2
            Case "TG"
                Set child = parent.getElementsByTagName(Arr(i + 1))(Arr(i + 2))
                i = i + 3
            Case "AT"
                For Each tg In parent.getElementsByTagName(Arr(i + 1))
                    Select Case Arr(i + 2)
                        Case "accesskey": If tg.AccessKey = Arr(i + 3) Then Set child = tg: Exit For
                        Case "alt": If tg.Alt = Arr(i + 3) Then Set child = tg: Exit For
                        Case "class": If tg.ClassName = Arr(i + 3) Then Set child = tg: Exit For
                        Case "id": If tg.ID = Arr(i + 3) Then Set child = tg: Exit For
                        Case "name": If tg.Name = Arr(i + 3) Then Set child = tg: Exit For
                        Case "type": If tg.Type = Arr(i + 3) Then Set child = tg: Exit For
                        Case "value": If tg.Value = Arr(i + 3) Then Set child = tg: Exit For
                    End Select
                Next
                i = i + 4
        End Select
        Set parent = child
            If i > UBound(Arr) Then
                flg = False
            End If
    Loop
    Set IEDOM = parent
End Function

Function IECaptcha()
    Call SetForegroundWindow(Application.Hwnd)
    IECaptcha = InputBox("Captcha")
End Function

