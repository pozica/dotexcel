Attribute VB_Name = "other_breg"
Sub BReg()
    Dim rng As Range
    Dim R As Integer
    
    If ActiveSheet.Name <> "main" Then Exit Sub
    For Each rng In Selection
        R = rng.Row
        If Cells(R, 4) <> "" Then Exit Sub   ' date-new
        If Cells(R, 28) = "" Then
            Select Case Cells(R, 10)         ' media-reg-id
                Case 1: If Cells(R, 13) = "" Then Call BReg001_1st(rng) Else: Call BReg001_2nd(rng)   ' 60Sec
                Case 2: If Cells(R, 13) = "" Then Call BReg002_1st(rng) Else: Call BReg002_2nd(rng)   ' Cyber Mode
                Case 3: If Cells(R, 13) = "" Then Call BReg003_1st(rng) Else: Call BReg003_2nd(rng)   ' High Space
                Case 4: If Cells(R, 13) = "" Then Call BReg004_1st(rng) Else: Call BReg004_2nd(rng)   ' Homepage Bank
                Case 5: If Cells(R, 13) = "" Then Call BReg005_1st(rng) Else: Call BReg005_2nd(rng)   ' Homeko
                Case 6: If Cells(R, 13) = "" Then Call BReg006_1st(rng) Else: Call BReg006_2nd(rng)   ' Mobagene
                Case 7: If Cells(R, 13) = "" Then Call BReg007_1st(rng) Else: Call BReg007_2nd(rng)   ' Bushido
                Case 8: If Cells(R, 13) = "" Then Call BReg008_1st(rng) Else: Call BReg008_2nd(rng)   ' MP7
                Case 9: If Cells(R, 13) = "" Then Call BReg009_1st(rng) Else: Call BReg009_2nd(rng)   ' No1 HP
                Case 10: If Cells(R, 13) = "" Then Call BReg010_1st(rng) Else: Call BReg010_2nd(rng)  ' 00HP
                Case 11: Call BReg011(rng)                                                            ' Chip
                Case 12: If Cells(R, 13) = "" Then Call BReg012_1st(rng) Else Call BReg012_2nd(rng)   ' Forest
                Case 13: Call BReg013(rng)                                                            ' Freepe
                Case 14: If Cells(R, 13) = "" Then Call BReg014_1st(rng) Else Call BReg014_2nd(rng)   ' iTool
                Case 15: If Cells(R, 13) = "" Then Call BReg015_1st(rng) Else Call BReg015_2nd(rng)   ' Mobalove
                Case 16: If Cells(R, 13) = "" Then Call BReg016_1st(rng) Else Call BReg016_2nd(rng)   ' Father of Homeo
                Case 17: Call BReg017(rng)                                                            ' Peps
                Case 18: Call BReg018(rng)                                                            ' Pandam
                Case 21: If Cells(R, 13) = "" Then Call BReg021_1st(rng) Else Call BReg021_2nd(rng)   ' Homeo
                Case 22: If Cells(R, 13) = "" Then Call BReg022_1st(rng) Else Call BReg022_2nd(rng)   ' Horent
                Case 27: Call BReg027(rng)                                                            ' Keitai Kakiko
                Case 28: Call BReg028(rng)                                                            ' At Pocket
                Case 33: If Cells(R, 13) = "" Then Call BReg033_1st(rng) Else: Call BReg033_2nd(rng)  ' 123 HP (Horent)
                Case 34: If Cells(R, 13) = "" Then Call BReg034_1st(rng) Else: Call BReg034_2nd(rng)  ' Magick Space (Horent)
                Case 36: If Cells(R, 13) = "" Then Call BReg036_1st(rng) Else: Call BReg036_2nd(rng)  ' Page 0 (Horent)
                Case 37: If Cells(R, 13) = "" Then Call BReg037_1st(rng) Else: Call BReg037_2nd(rng)  ' SEO Kenja (Horent)
                Case 44: If Cells(R, 13) = "" Then Call BReg044_1st(rng) Else: Call BReg044_2nd(rng)  '
                Case 47: Call BReg047(rng)                                                            '
                Case 48: If Cells(R, 13) = "" Then Call BReg048_1st(rng) Else: Call BReg048_2nd(rng)  '
                Case 49: If Cells(R, 13) = "" Then Call BReg049_1st(rng) Else: Call BReg049_2nd(rng)  '
                Case 50: If Cells(R, 13) = "" Then Call BReg050_1st(rng) Else: Call BReg050_2nd(rng)  '
                Case 51: If Cells(R, 13) = "" Then Call BReg051_1st(rng) Else: Call BReg051_2nd(rng)  '
                Case 52: If Cells(R, 13) = "" Then Call BReg052_1st(rng) Else: Call BReg052_2nd(rng)  '
                Case 53: If Cells(R, 13) = "" Then Call BReg053_1st(rng) Else: Call BReg053_2nd(rng)  '
                Case 54: If Cells(R, 13) = "" Then Call BReg054_1st(rng) Else: Call BReg054_2nd(rng)  '
                Case 55: If Cells(R, 13) = "" Then Call BReg055_1st(rng) Else: Call BReg055_2nd(rng)  '
                Case 56: If Cells(R, 13) = "" Then Call BReg056_1st(rng) Else: Call BReg056_2nd(rng)  '
                Case 57: If Cells(R, 13) = "" Then Call BReg057_1st(rng) Else: Call BReg057_2nd(rng)  '
                Case 58: If Cells(R, 13) = "" Then Call BReg058_1st(rng) Else: Call BReg058_2nd(rng)  '
                Case 59: If Cells(R, 13) = "" Then Call BReg059_1st(rng) Else: Call BReg059_2nd(rng)  '
                Case 60: If Cells(R, 13) = "" Then Call BReg060_1st(rng) Else: Call BReg060_2nd(rng)  '
                Case 61: If Cells(R, 13) = "" Then Call BReg061_1st(rng) Else: Call BReg061_2nd(rng)  '
                Case 62: If Cells(R, 13) = "" Then Call BReg062_1st(rng) Else: Call BReg062_2nd(rng)  '
                Case 63: If Cells(R, 13) = "" Then Call BReg063_1st(rng) Else: Call BReg063_2nd(rng)  '
                Case 64: If Cells(R, 13) = "" Then Call BReg064_1st(rng) Else: Call BReg064_2nd(rng)  '
                Case 65: If Cells(R, 13) = "" Then Call BReg065_1st(rng) Else: Call BReg065_2nd(rng)  '
                Case 66: If Cells(R, 13) = "" Then Call BReg066_1st(rng) Else: Call BReg066_2nd(rng)  '
                Case 67: If Cells(R, 13) = "" Then Call BReg067_1st(rng) Else: Call BReg067_2nd(rng)  '
                Case 68: If Cells(R, 13) = "" Then Call BReg068_1st(rng) Else: Call BReg068_2nd(rng)  '
                Case 69: If Cells(R, 13) = "" Then Call BReg069_1st(rng) Else: Call BReg069_2nd(rng)  '
                Case 70: If Cells(R, 13) = "" Then Call BReg070_1st(rng) Else: Call BReg070_2nd(rng)  '
            End Select
        ElseIf Cells(R, 28) <> "" And Cells(R, 34) <> "" Then
            Select Case Cells(R, 10)   ' media-reg-id
                Case 1: Call BReg001_3rd(rng)
                Case 2: Call BReg002_3rd(rng)
                Case 3: Call BReg003_3rd(rng)
                Case 4: Call BReg004_3rd(rng)
                Case 5: Call BReg005_3rd(rng)
                Case 6: Call BReg006_3rd(rng)
                Case 7: Call BReg007_3rd(rng)
                Case 8: Call BReg008_3rd(rng)
                Case 9: Call BReg009_3rd(rng)
                Case 10: Call BReg010_3rd(rng) ' 00HP
                Case 11: Call BReg011_3rd(rng) ' Chip
                Case 12: Call BReg012_3rd(rng) ' Forest
                Case 13: Call BReg013_3rd(rng) ' Freepe
                Case 14: Call BReg014_3rd(rng) ' iTool
                Case 15: Call BReg015_3rd(rng) ' Mobalove
                Case 16: Call BReg016_3rd(rng) ' Father of Homeo
                Case 17: Call BReg017_3rd(rng) ' Peps
                Case 18: Call BReg018_3rd(rng) ' Pandam
                Case 21: Call BReg021_3rd(rng) ' Homeo
                Case 22: Call BReg022_3rd(rng) ' Horent
                Case 27: Call BReg027_3rd(rng) ' Keitai Kakiko
                Case 28: Call BReg028_3rd(rng) ' At Pocket
                Case 33: Call BReg033_3rd(rng)
                Case 34: Call BReg034_3rd(rng)
                Case 36: Call BReg036_3rd(rng)
                Case 37: Call BReg037_3rd(rng)
                Case 44: Call BReg044_3rd(rng)
                Case 47: Call BReg047_3rd(rng)
                Case 48: Call BReg048_3rd(rng)
                Case 49: Call BReg049_3rd(rng)
                Case 50: Call BReg050_3rd(rng)
                Case 51: Call BReg051_3rd(rng)
                Case 52: Call BReg052_3rd(rng)
                Case 53: Call BReg053_3rd(rng)
                Case 54: Call BReg054_3rd(rng)
                Case 55: Call BReg055_3rd(rng)
                Case 56: Call BReg056_3rd(rng)
                Case 57: Call BReg057_3rd(rng)
                Case 58: Call BReg058_3rd(rng)
                Case 59: Call BReg059_3rd(rng)
                Case 60: Call BReg060_3rd(rng)
                Case 61: Call BReg061_3rd(rng)
                Case 62: Call BReg062_3rd(rng)
                Case 63: Call BReg063_3rd(rng)
                Case 64: Call BReg064_3rd(rng)
                Case 65: Call BReg065_3rd(rng)
                Case 66: Call BReg066_3rd(rng)
                Case 67: Call BReg067_3rd(rng)
                Case 68: Call BReg068_3rd(rng)
                Case 69: Call BReg069_3rd(rng)
                Case 70: Call BReg070_3rd(rng)
            End Select
        End If
    Next
End Sub

Sub BReg001_1st(rng)    ' 60Sec
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 7), "NM", "mail", 1)   ' email
        Call IEInput(ie, Cells(R, 32), "NM", "password", 1) ' media-login-pw
        Call IEDOM(ie, Array("ID", "main_con_txt2", "AT", "input", "alt", "ñ≥óøìoò^")).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg001_2nd(rng)    ' 60Sec
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg001_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg002_1st(rng)    ' Cyber Mode
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12), True)   ' media-reg-url
        Call IEInput(ie, Cells(R, 7), "NM", "mail", 0)   ' email
        Call IEInput(ie, IECaptcha, "NM", "keystring", 0)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg002_2nd(rng)    ' Cyber Mode
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEInput(ie, Cells(R, 31), "NM", "id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "pass", 0) ' media-login-pw
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Call IEClick(ie, "TG", "a", 0)
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Call IEInput(ie, Cells(R, 31), "NM", "id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "pass", 0) ' media-login-pw
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        'Call IEClick(ie, "TG", "a", 3)
        'Call IEWait(ie)
        Call IEClick(ie, "AT", "a", , "accesskey", "1")
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 31), "NM", "cid", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 17), "NM", "ctitle", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Call IEClick(ie, "TG", "a", 1)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("TG", "a", 0)).href  ' media-url
        Cells(R, 4) = Date ' date-new
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg002_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg003_1st(rng)    ' High Space
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 7), "NM", "mail", 1)   ' email
        Call IEInput(ie, Cells(R, 32), "NM", "password", 1) ' media-login-pw
        Call IEDOM(ie, Array("TG", "form", 1, "AT", "input", "type", "submit")).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg003_2nd(rng)    ' High Space
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "sub_title", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg003_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg004_1st(rng)    ' Homepage Bank
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 7), "NM", "mail", 1)   ' email
        Call IEInput(ie, Cells(R, 32), "NM", "password", 1) ' media-login-pw
        Call IEDOM(ie, Array("TG", "form", 1, "AT", "input", "class", "signup_button")).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg004_2nd(rng)    ' Homepage Bank
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call iedoc(ie, Array("ID", "left_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call iedoc(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg004_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg005_1st(rng)    ' Homeko
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 19), "NM", "name", 0) ' media-reg-name
        Call IEInput(ie, "t", "NM", "sei_flg", 0) ' media-reg-sex: [f] female [t] male
        Call IEInput(ie, Cells(R, 24), "NM", "todoufuken", 0)   ' media-reg-area
        Call IEInput(ie, Cells(R, 20), "NM", "birth_year", 0)   ' media-reg-birth-m
        Call IEInput(ie, Cells(R, 21), "NM", "birth_month", 0)   ' media-reg-birth-m
        Call IEInput(ie, Cells(R, 22), "NM", "birth_day", 0)   ' media-reg-birth-d
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)   ' email
        Call IEInput(ie, Cells(R, 7), "NM", "email2", 0)  ' email
        Call IEInput(ie, Cells(R, 26), "NM", "subdomain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "directory", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "password", 0) ' media-login-pw
        Call IEDOM(ie, Array("ID", "f_reg", "AT", "input", "class", "send_btn")).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg005_2nd(rng)    ' Homeko
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEInput(ie, Cells(R, 31), "NM", "usr_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "usr_pass1", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 32), "NM", "usr_pass2", 0) ' media-login-pw
        Call IEInput(ie, IECaptcha, "NM", "captcha_pass", 0)
        Call IEClick(ie, "AT", "input", , "value", "åàíËÅI")
        Call IEWait(ie)
        Call IEClick(ie, "NM", "a", 1)
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 17), "NM", "hp_title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 19), "NM", "h_name", 0)  ' media-reg-name
        Call IEInput(ie, Cells(R, 14), "NM", "bcate", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 24), "NM", "area", 0)   ' media-reg-area
        Call IEInput(ie, Cells(R, 20) & Cells(R, 21) & Cells(R, 22), "NM", "birth", 0) ' media-reg-birth-y, m, d
        Call IEInput(ie, Cells(R, 18), "NM", "naiyou", 0)  ' media-reg-desc
        Call IEClick(ie, "AT", "input", , "value", "åàíË")
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 15), "NM", "type", 0) ' media-reg-category-2
        Call IEClick(ie, "AT", "input", , "value", "åàíË")
        Call IEWait(ie)
        Call IEClick(ie, "NM", "a", 0)
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 30) = IEDOM(ie, Array("ID", "header_menu", "TG", "ul", 0, "TG", "li", 6, "TG", "a", 0)).href  ' media-url
        Cells(R, 4) = Date ' date-new
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg005_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg006_1st(rng)    ' Mobagene
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEClick(ie, "NM", "sex", 0) ' media-reg-sex: [1] female [0] male
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)   ' email
        Call IEInput(ie, Cells(R, 32), "NM", "pass", 0) ' media-login-pw
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg006_2nd(rng)    ' Mobagene
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEInput(ie, Cells(R, 32), "NM", "pass", 0) ' media-login-pw
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 31), "NM", "id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "categoid", 0) ' media-reg-category-1
        Call IEClick(ie, "NM", "read", 1)
        Call IEInput(ie, Cells(R, 19), "NM", "name", 0)  ' media-reg-name
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)   ' email
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Call IEGotoURL(ie, "http://mg1.jp/e/")
        Call IEInput(ie, Cells(R, 31), "NM", "ID", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "PW", 0) ' media-login-pw
        Call IEClick(ie, "ID", "ok")
        Call IEWait(ie)
    Set ie = Nothing
End Sub

Sub BReg006_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg007_1st(rng)    ' Bushido
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 7), "NM", "mail")    ' email
        Call IEInput(ie, Cells(R, 32), "NM", "password", 0) ' media-login-pw
        Call IEClick(ie, "AT", "input", , "class", "img1")
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg007_2nd(rng)    ' Bushido
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "left", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0)  ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0)  ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).href ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).Click
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg007_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg008_1st(rng)    ' MP7
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 7), "NM", "from_mail", 0)   ' email
        Call IEClick(ie, "NM", "sex", 0) ' media-reg-sex: [1] female [0] male
        Call IEClick(ie, "NM", "Submit", 0)
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg008_2nd(rng)    ' MP7
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13), True)   ' media-reg-url-2
        Call IEInput(ie, Cells(R, 31), "NM", "usr_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "usr_pass1", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 32), "NM", "usr_pass2", 0) ' media-login-pw
        Call IEInput(ie, 1, "NM", "sex", 0) ' media-reg-sex: [2] female [1] male
        Call IEInput(ie, IECaptcha, "NM", "img_str", 0)
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Call IEClick(ie, "TG", "a", 0)
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 17), "NM", "hp_title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 19), "NM", "h_name", 0)  ' media-reg-name
        Call IEInput(ie, Cells(R, 14), "NM", "bcate", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 24), "NM", "area", 0)   ' media-reg-area
        Call IEInput(ie, "4", "NM", "age", 0) ' media-reg-birth: [3] 16Å`20 [4] 21Å`25 [5] 26Å`30
        Call IEInput(ie, Cells(R, 18), "NM", "naiyou", 0)  ' media-reg-desc
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 15), "NM", "type", 0) ' media-reg-category-2
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Call IEClick(ie, "NM", "melmag_ad_flg", 0)
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Call IEClick(ie, "TG", "a", 1)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("TG", "center", 0, "TG", "blink", 0, "TG", "a", 0)).href ' media-url
        Cells(R, 29) = "http://mh.mp7.jp/?f=" ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg008_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg009_1st(rng)    ' No1 HP
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 7), "NM", "mail", 0)   ' email
        Call IEInput(ie, Cells(R, 32), "NM", "password", 0) ' media-login-pw
        Call IEClick(ie, "AT", "input", , "type", "image")
        Call IEWait(ie)
        Call ie.Quit
        'Call IEInput(ie, Cells(R, 7), "NM", "mail", 0)   ' email
        'Call IEInput(ie, Cells(R, 32), "NM", "password", 0) ' media-login-pw
        'Call IEClick(ie, "AT", "input", , "type", "image")
    Set ie = Nothing
End Sub

Sub BReg009_2nd(rng)    ' No1 HP
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEInput(ie, Cells(R, 31), "NM", "usr_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "usr_pass1", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 32), "NM", "usr_pass2", 0) ' media-login-pw
        Call IEInput(ie, 1, "NM", "sex", 0) ' media-reg-sex: [2] female [1] male
        Call IEInput(ie, IECaptcha, "NM", "captcha_pass", 0)
        Call IEClick(ie, "AT", "input", , "value", "åàíËÅI")
        Call IEWait(ie)
        Call IEClick(ie, "NM", "a", 1)
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 17), "NM", "hp_title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 19), "NM", "h_name", 0)  ' media-reg-name
        Call IEInput(ie, Cells(R, 14), "NM", "bcate", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 24), "NM", "area", 0)   ' media-reg-area
        Call IEInput(ie, Cells(R, 20) & Cells(R, 21) & Cells(R, 22), "NM", "birth", 0) ' media-reg-birth-y, m, d
        Call IEInput(ie, Cells(R, 18), "NM", "naiyou", 0)  ' media-reg-desc
        Call IEClick(ie, "AT", "input", , "value", "åàíË")
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 15), "NM", "type", 0) ' media-reg-category-2
        Call IEClick(ie, "AT", "input", , "value", "åàíË")
        Call IEWait(ie)
        Call IEClick(ie, "NM", "a", 0)
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 30) = IEDOM(ie, Array("ID", "header_menu", "TG", "ul", 0, "TG", "li", 6, "TG", "a", 0)).href  ' media-url
        Cells(R, 4) = Date ' date-new
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg009_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg010_1st(rng)    ' 00HP
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 7), "ID", "mailsetinp")    ' email
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg010_2nd(rng)    ' 00HP
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13), True)   ' media-reg-url-2
        Call IEInput(ie, Cells(R, 31), "NM", "usr_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "usr_pass1", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 32), "NM", "usr_pass2", 0) ' media-login-pw
        Call IEInput(ie, 1, "NM", "sex", 0) ' media-reg-sex: [2] female [1] male
        Call IEInput(ie, IECaptcha, "NM", "captcha_pass", 0)
        Call IEClick(ie, "AT", "input", , "value", "åàíËÅI")
        Call IEWait(ie)
        Call IEClick(ie, "TG", "a", 1)
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 17), "NM", "hp_title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 19), "NM", "h_name", 0)  ' media-reg-name
        Call IEInput(ie, Cells(R, 14), "NM", "bcate", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 24), "NM", "area", 0)   ' media-reg-area
        Call IEInput(ie, Cells(R, 20) & Cells(R, 21) & Cells(R, 22), "NM", "birth", 0) ' media-reg-birth-y, m, d
        Call IEInput(ie, Cells(R, 18), "NM", "naiyou", 0)  ' media-reg-desc
        Call IEClick(ie, "AT", "input", , "value", "åàíË")
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 15), "NM", "type", 0) ' media-reg-category-2
        Call IEClick(ie, "AT", "input", , "value", "åàíË")
        Call IEWait(ie)
        Call IEClick(ie, "TG", "a", 0)
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 30) = IEDOM(ie, Array("ID", "header_menu", "TG", "ul", 0, "TG", "li", 6, "TG", "a", 0)).href  ' media-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("AT", "div", "class", "login_menu", "TG", "a", 1)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg010_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg011(rng)    ' Chip
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12), True)   ' media-reg-url
        Call IEInput(ie, Cells(R, 31), "NM", "id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "ps", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 7), "NM", "mail", 0)    ' email
        Call IEInput(ie, Cells(R, 19), "NM", "title", 0)  ' media-reg-name
        Call IEClick(ie, "NM", "sex", 1)   ' media-reg-sex: [0] female [1] male
        Call IEClick(ie, "NM", "act", 0)
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 14), "NM", "categoly_id", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 25), "NM", "work", 0) ' media-reg-work
        Call IEInput(ie, Cells(R, 20), "NM", "birth[y]", 0)   ' media-reg-birth-m
        Call IEInput(ie, Cells(R, 21), "NM", "birth[m]", 0)   ' media-reg-birth-m
        Call IEInput(ie, Cells(R, 22), "NM", "birth[d]", 0)   ' media-reg-birth-d
        Call IEInput(ie, Cells(R, 24), "NM", "area", 0)   ' media-reg-area
        Call IEClick(ie, "NM", "kiyaku_check", 0)  ' kiyaku_check
        Call IEInput(ie, IECaptcha, "NM", "captcha_text_a", 0)   ' captcha_text_a
        Call IEClick(ie, "AT", "input", , "value", "ìoò^Ç∑ÇÈ")
        Call IEWait(ie)
        Call IEClick(ie, "TG", "a", 1)
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 19), "NM", "v[0]", 0)  ' media-reg-name
        Call IEInput(ie, Cells(R, 20) & "/" & Cells(R, 21) & "/" & Cells(R, 22), "NM", "v[2]", 0) ' media-reg-birth
        Call IEInput(ie, Cells(R, 25), "NM", "v[8]", 0) ' media-reg-work
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Call IEClick(ie, "TG", "a", 0)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 30) = IEDOM(ie, Array("TG", "a", 0)).href  ' media-url
        Cells(R, 4) = Date ' date-new
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg011_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg012_1st(rng)    ' Forest
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 7), "NM", "mail", 0)  ' email
        Call IEInput(ie, Cells(R, 7), "NM", "mail_conf", 0)  ' email
        Call IEClick(ie, "AT", "input", , "alt", "ëóêM")
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg012_2nd(rng)    ' Forest
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13), True)   ' media-reg-url-2
        Call IEInput(ie, Cells(R, 31), "NM", "fid", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "pw", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 19), "NM", "nm", 0)  ' media-reg-name
        Call IEInput(ie, Cells(R, 25), "NM", "job", 0) ' media-reg-work
        Call IEInput(ie, Cells(R, 24), "NM", "area", 0)   ' media-reg-area
        Call IEInput(ie, Cells(R, 20), "NM", "bd[Year]", 0)   ' media-reg-birth-m
        Call IEInput(ie, Cells(R, 21), "NM", "bd[Month]", 0)   ' media-reg-birth-m
        Call IEInput(ie, Cells(R, 22), "NM", "bd[Day]", 0)   ' media-reg-birth-d
        Call IEInput(ie, 1, "NM", "gender", 0)   ' media-reg-sex: [2] female [1] male
        Call IEInput(ie, 2, "NM", "ques", 0)   ' media-reg-question: [1] êeÇÃãåê© [2] ÉyÉbÉgÇÃñºëO [3] åôÇ¢Ç»êHÇ◊ï® [4] èâóˆÇÃêlÇÃñºëO
        Call IEInput(ie, "kuma", "NM", "answ", 0)   ' media-reg-answer
        Call IEClick(ie, "NM", "agreement", 0)
        Call IEInput(ie, IECaptcha, "NM", "imgkey_r", 0)    ' captcha
        Call IEClick(ie, "NM", "action_cadd", 0)
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 31), "NM", "hpid", 0) ' media-login-id
        Call IEClick(ie, "NM", "action_cadd2", 0)
        Call IEWait(ie)
        Call IEDOM(ie, Array("AT", "div", "class", "halfbox", "TG", "a", 0)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 30) = IEDOM(ie, Array("AT", "div", "class", "kaisetsu", "TG", "table", 1, "TG", "b", 0)).innerText ' media-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "ichiran", "TG", "a", 8)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg012_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg013(rng)    ' Freepe
    Dim ie As Object
    Dim R As Integer
    Dim x As String
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12), True)   ' media-reg-url
        Call IEInput(ie, Cells(R, 31), "NM", "id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 7), "NM", "account", 0)    ' email
        Call IEInput(ie, IECaptcha, "NM", "cnf_key", 0)  ' captcha
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg013_3nd(rng)    ' Freepe
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEInput(ie, Cells(R, 31), "NM", "id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "ps", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 19), "NM", "title", 0)  ' media-reg-name
        Call IEClick(ie, "NM", "sex", 1)   ' media-reg-sex: [0] female [1] male
        Call IEClick(ie, "NM", "act", 0)
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 14), "NM", "categoly_id", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 24), "NM", "work", 0) ' media-reg-work
        Call IEInput(ie, Cells(R, 20), "NM", "birth[y]", 0)   ' media-reg-birth-m
        Call IEInput(ie, Cells(R, 21), "NM", "birth[m]", 0)   ' media-reg-birth-m
        Call IEInput(ie, Cells(R, 22), "NM", "birth[d]", 0)   ' media-reg-birth-d
        Call IEInput(ie, Cells(R, 24), "NM", "area", 0)   ' media-reg-area
        Call IEClick(ie, "NM", "kiyaku_check", 0)  ' kiyaku_check
        Call IEFocus(ie, "NM", "captcha_text_a", 0)    ' captcha_text_a
    Set ie = Nothing
End Sub

Sub BReg014_1st(rng)    ' iTool
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 7), "NM", "data[Mail][mail]", 0)    ' email
        Call IEDOM(ie, Array("ID", "MailAddForm", "AT", "input", "type", "submit")).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg014_2nd(rng)    ' iTool
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEInput(ie, Cells(R, 31), "NM", "data[User][username]", 1) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "data[User][password]", 1) ' media-login-pw
        Call IEClick(ie, "AT", "input", , "value", "åàíË")
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 19), "NM", "data[User][nickname]", 0)  ' media-reg-name
        Call IEInput(ie, Cells(R, 14), "NM", "data[User][category_id]", 0) ' media-reg-category-1
        Call IEInput(ie, 2, "NM", "data[User][sex_id]", 0)   ' media-reg-sex: [3] female [2] male
        Call IEInput(ie, 5, "NM", "data[User][age_id]", 0)   ' media-reg-birth: [4] 15-19 [5] 20-24 [6] 25-29
        Call IEInput(ie, Cells(R, 24), "NM", "data[User][area_id]", 0)   ' media-reg-area
        Call IEClick(ie, "AT", "input", , "value", "åàíË")
        Call IEWait(ie, 200)
        Call IEInput(ie, Cells(R, 16), "NM", "data[Page][title]", 0)  ' media-reg-keyword
        Call IEInput(ie, 7, "NM", "data[Page][imagenavi_id]", 0)   ' [7] Hiyoko
        Call IEInput(ie, Cells(R, 17), "NM", "data[Page][comment]", 0)  ' media-reg-title
        Call IEInput(ie, 1, "NM", "data[Page][coloritem_id]", 0)   ' [1] Skyblue
        Call IEInput(ie, Cells(R, 18), "NM", "data[Page][introduce]", 0)  ' media-reg-desc
        Call IEClick(ie, "AT", "input", , "value", "åàíË")
        Call IEWait(ie, 200)
        'Call IEClick(ie, "AT", "input", , "value", "åàíË")
        Call IEClick(ie, "ID", "n3")
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 30) = IEDOM(ie, Array("TG", "form", 0, "TG", "input", 0)).Value  ' media-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "login1", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg014_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg015_1st(rng)    ' Mobalove
    Dim ie As Object
    Dim R As Integer
    Dim x As String
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 31), "NM", "id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "pw", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEClick(ie, "NM", "BT_EDIT_OK", 0)
        Call IEWait(ie)
        Call IEClick(ie, "NM", "BT_CONFIRM_OK", 0)
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg015_2nd(rng)    ' Mobalove
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEClick(ie, "NM", "BT_OK", 0)
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 15), "NM", "category", 0) ' media-reg-category-2
        Call IEClick(ie, "NM", "BT_OK", 0)
        Call IEWait(ie)
        Call IEClick(ie, "TG", "a", 0)
        Call IEWait(ie)
        'Cells(R, 29) = IEDOM(ie, Array("TG", "a", 0)).href & "&pw=" & Cells(R, 7) ' media-login-url
        'Cells(R, 30) = ie.document.url  ' media-url
        'Cells(R, 4) = Date ' date-new
        'Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg015_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg016_1st(rng)    ' Father of Homeo
    Dim ie As Object
    Dim R As Integer
    Dim x As String
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 19), "NM", "name", 0)  ' media-reg-name
        Call IEInput(ie, "t", "NM", "sei_flg", 0)   ' media-reg-sex: [f] female [t] male
        Call IEInput(ie, Cells(R, 24), "NM", "todoufuken", 0)   ' media-reg-area
        Call IEInput(ie, Cells(R, 20), "NM", "birth_year", 0)   ' media-reg-birth-m
        Call IEInput(ie, Cells(R, 21), "NM", "birth_month", 0)   ' media-reg-birth-m
        Call IEInput(ie, Cells(R, 22), "NM", "birth_day", 0)   ' media-reg-birth-d
        Call IEInput(ie, Cells(R, 7), "NM", "email", 1)    ' email
        Call IEInput(ie, Cells(R, 7), "NM", "email2", 0)    ' email
        Call IEInput(ie, Cells(R, 32), "NM", "password", 1) ' media-login-pw
        Call IEInput(ie, Cells(R, 31), "NM", "login_id", 0) ' media-login-id
        Call IEClick(ie, "AT", "input", , "value", "Å@ëóÅ@êMÅ@")
        Call IEWait(ie)
        Call IEClick(ie, "AT", "input", , "value", "ìoò^")
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg016_2nd(rng)    ' Father of Homeo
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Cells(R, 29) = "http://pasiphae.biz/" ' media-login-url
        Cells(R, 30) = "http://" & Cells(R, 31) & ".pasiphae.biz/"  ' media-url
        Cells(R, 4) = Date ' date-new
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg016_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg017(rng)    ' Peps
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12), True)   ' media-reg-url
        Call IEInput(ie, Cells(R, 31), "NM", "id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "ps", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 7), "NM", "mail", 0)    ' email
        Call IEInput(ie, Cells(R, 17), "NM", "title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 14), "NM", "cate_l", 0)  ' media-reg-category-1
        Call IEClick(ie, "NM", "act", 0)
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 15), "NM", "categoly_id", 0) ' media-reg-category-2
        Call IEInput(ie, Cells(R, 20), "NM", "birth[y]", 0)   ' media-reg-birth-y
        Call IEInput(ie, Cells(R, 21), "NM", "birth[m]", 0)   ' media-reg-birth-m
        Call IEInput(ie, Cells(R, 22), "NM", "birth[d]", 0)   ' media-reg-birth-d
        Call IEClick(ie, "NM", "sex", 1)   ' media-reg-sex: [0] female [1] male
        Call IEInput(ie, Cells(R, 24), "NM", "area", 0)   ' media-reg-area
        Call IEClick(ie, "NM", "kiyaku_check", 0)  ' kiyaku_check
        Call IEInput(ie, IECaptcha, "NM", "captcha_text_a", 0)   ' captcha_text_a
        Call IEClick(ie, "AT", "input", , "value", "ìoò^Ç∑ÇÈ")
        Call IEWait(ie)
        Call IEClick(ie, "NM", "select_tmpl", 0)
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url  ' media-login-url
        Cells(R, 30) = IEDOM(ie, Array("TG", "a", 1)).href ' media-url
        Cells(R, 4) = Date ' date-new
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg017_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg018(rng)    ' Pandam
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12), True)   ' media-reg-url
        Call IEInput(ie, Cells(R, 19), "NM", "name", 0)  ' media-reg-name
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEClick(ie, "AT", "input", , "value", "ëóêM")
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0)  ' media-reg-domain
        Call IEInput(ie, Cells(R, 14), "NM", "sub_domain", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 31), "NM", "account", 0) ' media-login-id
        Call IEClick(ie, "AT", "input", , "value", "ëóêM")
        Call IEWait(ie)
        Call IEInput(ie, IECaptcha, "NM", "pass1", 0) ' captcha
        Call IEClick(ie, "AT", "input", , "value", "ÇÕÇ¢")
        Call IEWait(ie)
        Cells(R, 29) = "http://" & Cells(R, 26) & "/" & Cells(R, 14) & "/pandam.php" ' media-login-url
        Cells(R, 30) = "http://" & Cells(R, 26) & "/" & Cells(R, 14) & "/" ' media-url
        Cells(R, 4) = Date ' date-new
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg018_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg021_1st(rng)    ' Homeo
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 19), "NM", "name", 0)  ' media-reg-name
        Call IEInput(ie, "t", "NM", "sei_flg", 0)   ' media-reg-sex: [f] female [t] male
        Call IEInput(ie, Cells(R, 24), "NM", "todoufuken", 0)   ' media-reg-area
        Call IEInput(ie, Cells(R, 20), "NM", "birth_year", 0)   ' media-reg-birth-m
        Call IEInput(ie, Cells(R, 21), "NM", "birth_month", 0)   ' media-reg-birth-m
        Call IEInput(ie, Cells(R, 22), "NM", "birth_day", 0)   ' media-reg-birth-d
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEInput(ie, Cells(R, 7), "NM", "email2", 0)    ' email
        Call IEInput(ie, Cells(R, 32), "NM", "password", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 31), "NM", "login_id", 0) ' media-login-id
        Call IEClick(ie, "AT", "input", , "class", "send_btn")
        Call IEWait(ie)
        Call IEClick(ie, "AT", "input", , "value", "Å@ëóÅ@êMÅ@")
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg021_2nd(rng)    ' Homeo
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Cells(R, 29) = "http://enceladus.biz/" ' media-login-url
        Cells(R, 30) = "http://" & Cells(R, 31) & ".enceladus.biz/"  ' media-url
        Cells(R, 4) = Date ' date-new
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg021_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg022_1st(rng)    ' Horent
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 7), "NM", "mail", 1)    ' email
        Call IEInput(ie, Cells(R, 32), "NM", "password", 1) ' media-login-pw
        Call IEClick(ie, "AT", "input", , "value", "ìoò^Ç∑ÇÈ")
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg022_2nd(rng)    ' Horent
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "manage", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0)  ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0)  ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).href ' media-url
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg022_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg027(rng)    ' Keitai Kakiko
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEClick(ie, "NM", "agree", 0)
        Call IEInput(ie, Cells(R, 31), "NM", "account", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEInput(ie, Cells(R, 19), "NM", "name", 0)  ' media-reg-name
        Call IEClick(ie, "NM", "join", 0)
        Call IEWait(ie)
        Cells(R, 29) = IEDOM(ie, Array("TG", "center", 1, "TG", "a", 1)).href ' media-login-url
        Cells(R, 30) = IEDOM(ie, Array("TG", "center", 1, "TG", "a", 0)).href ' media-url
        Cells(R, 4) = Date ' date-new
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg027_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg028(rng)    ' At Pocket
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 31), "NM", "id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "pass", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 17), "NM", "title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEInput(ie, 2, "NM", "age", 0)   ' media-reg-birth: [1] 10 [2] 20 [3] 30
        Call IEInput(ie, 1, "NM", "sex", 0)   ' media-reg-sex: [2] female [1] male [3] other
        Call IEInput(ie, Cells(R, 24), "NM", "area", 0)   ' media-reg-area
        Call IEInput(ie, Cells(R, 18), "NM", "comment", 0)  ' media-reg-desc
        Call IEInput(ie, Cells(R, 14), "NM", "cate", 0) ' media-reg-category-1
        Call IEClick(ie, "AT", "input", , "value", "çÏê¨")
        Call IEWait(ie)
        Call IEClick(ie, "AT", "input", , "value", "äÆóπ")
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("TG", "a", 1)).href ' media-url
        Cells(R, 4) = Date ' date-new
        Call IEWait(ie)
        Call IEClick(ie, "TG", "a", 0)
        Call IEWait(ie)
        Cells(R, 29) = "http://atpk.jp/login.php?id=" & Cells(R, 31) & "&pass=" & Cells(R, 32) & "&mode=login" ' media-login-url
        Call IEWait(ie)
        Call IEClick(ie, "TG", "a", 2)
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 18), "NM", "input_header", 0)  ' media-reg-desc
        Call IEClick(ie, "AT", "input", , "value", "ê›íË")
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg028_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg033_1st(rng)    ' 123 HP (Horent)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 7), "NM", "mail", 1)    ' email
        Call IEInput(ie, Cells(R, 32), "NM", "password", 1) ' media-login-pw
        Call IEDOM(ie, Array("TG", "form", 1, "AT", "input", "type", "image")).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg033_2nd(rng)    ' 123 HP (Horent)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "right_content_middle", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0)  ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0)  ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).href ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 1)).Click
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg033_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg034_1st(rng)    ' Magic Space (Horent)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 7), "NM", "mail", 0)    ' email
        Call IEInput(ie, Cells(R, 32), "NM", "password", 0) ' media-login-pw
        Call IEDOM(ie, Array("TG", "form", 0, "AT", "input", "type", "image")).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg034_2nd(rng)    ' Magic Space (Horent)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "left_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0)  ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0)  ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).href ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 1)).Click
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg034_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg036_1st(rng)    ' Page 0 (Horent)
    Dim ie As Object
    Dim R As Integer
    Dim x As String
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 7), "NM", "mail", 1)    ' email
        Call IEInput(ie, Cells(R, 32), "NM", "password", 1) ' media-login-pw
        Call IEClick(ie, "AT", "input", , "value", "êVãKìoò^")
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg036_2nd(rng)    ' Page 0 (Horent)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "content_box", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0)  ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0)  ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).href ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 1)).Click
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg036_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg037_1st(rng)    ' SEO Kenja (Horent)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 7), "NM", "mail", 1)    ' email
        Call IEInput(ie, Cells(R, 32), "NM", "password", 1) ' media-login-pw
        Call IEDOM(ie, Array("TG", "form", 1, "AT", "input", "type", "image")).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg037_2nd(rng)    ' SEO Kenja (Horent)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0)  ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0)  ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).href ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 1)).Click
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg037_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg042_1st(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 31), "NM", "user_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "user_pass", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEInput(ie, Cells(R, 17), "NM", "title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0)  ' media-reg-category-1
        Call IEInput(ie, IECaptcha, "NM", "phrase", 0)  ' captcha
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg042_2nd(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0)  ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).href ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 1)).Click
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg042_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg043_1st(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 31), "NM", "user_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "user_pass", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEInput(ie, Cells(R, 17), "NM", "title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0)  ' media-reg-category-1
        Call IEInput(ie, IECaptcha, "NM", "phrase", 0)  ' captcha
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg043_2nd(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0)  ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).href ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 1)).Click
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg043_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg044_1st(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12), True)   ' media-reg-url
        Call IEInput(ie, Cells(R, 31), "NM", "user_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "user_pass", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEInput(ie, Cells(R, 17), "NM", "title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0)  ' media-reg-category-1
        Call IEInput(ie, IECaptcha, "NM", "phrase", 0)  ' captcha
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg044_2nd(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEGotoURL(ie, "http://www.bjwbggw.com/login.php")
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 31), "NM", "id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "pw", 0) ' media-login-pw
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("TG", "a", 0)).href ' media-url
        Call IEClick(ie, "TG", "a", 3)
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 18), "NM", "up", 0)  ' media-reg-desc
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Cells(R, 29) = "http://www.bjwbggw.com/login.php?id=" & Cells(R, 31) & "&pw=" & Cells(R, 32) ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg044_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg045_1st(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 31), "NM", "user_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "user_pass", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEInput(ie, Cells(R, 17), "NM", "title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0)  ' media-reg-category-1
        Call IEInput(ie, IECaptcha, "NM", "phrase", 0)  ' captcha
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg045_2nd(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0)  ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).href ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 1)).Click
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg045_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg046_1st(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 31), "NM", "user_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "user_pass", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEInput(ie, Cells(R, 17), "NM", "title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0)  ' media-reg-category-1
        Call IEInput(ie, IECaptcha, "NM", "phrase", 0)  ' captcha
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg046_2nd(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0)  ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).href ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 1)).Click
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg046_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg047(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12), True)   ' media-reg-url
        Call IEClick(ie, "NM", "kiyaku", 0)
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 31), "NM", "id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "ps", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 7), "NM", "mail", 0)    ' email
        Call IEInput(ie, Cells(R, 14), "NM", "purpose", 0)  ' media-reg-category-1
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Cells(R, 29) = "http://o.z-z.jp/menu1.cgi?ps=" & Cells(R, 32) & "&id=" & Cells(R, 31) & "&master=1" ' media-login-url
        Cells(R, 30) = IEDOM(ie, Array("TG", "a", 0)).href ' media-url
        Cells(R, 4) = Date ' date-new
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg047_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg048_1st(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEInput(ie, Cells(R, 32), "NM", "password", 0) ' media-login-pw
        Call IEClick(ie, "AT", "input", , "value", "ÉçÉOÉCÉì")
        Call IEWait(ie)
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg048_2nd(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 31), "NM", "subdomain", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category_id", 0)  ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "title", 0)  ' media-reg-title
        Call IEClick(ie, "AT", "input", , "type", "image")
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "main", "TG", "table", 0, "TG", "a", 0)).href ' media-url
        Cells(R, 29) = "http://mobile-sp.com/" ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "login", "TG", "a", 1)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg048_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg049_1st(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 7), "NM", "email", 1)    ' email
        Call IEInput(ie, Cells(R, 32), "NM", "password", 1) ' media-login-pw
        Call IEClick(ie, "AT", "input", , "value", "Ç±ÇÃì‡óeÇ≈ñ≥óøìoò^")
        Call IEWait(ie)
        Call IEClick(ie, "AT", "input", , "value", "Ç±ÇÃì‡óeÇ≈ñ≥óøìoò^")
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg049_2nd(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 31), "NM", "subdomain", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category_id", 0)  ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "title", 0)  ' media-reg-title
        Call IEClick(ie, "AT", "input", , "type", "image")
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "main", "TG", "table", 0, "TG", "a", 0)).href ' media-url
        Cells(R, 29) = "http://thebe.biz/" ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "login", "TG", "a", 1)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg049_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg050_1st(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 19), "NM", "name", 0) ' media-reg-name
        Call IEInput(ie, "t", "NM", "sei_flg", 0) ' media-reg-sex: [f] female [t] male
        Call IEInput(ie, Cells(R, 24), "NM", "todoufuken", 0)   ' media-reg-area
        Call IEInput(ie, Cells(R, 20), "NM", "birth_year", 0)   ' media-reg-birth-m
        Call IEInput(ie, Cells(R, 21), "NM", "birth_month", 0)   ' media-reg-birth-m
        Call IEInput(ie, Cells(R, 22), "NM", "birth_day", 0)   ' media-reg-birth-d
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)   ' email
        Call IEInput(ie, Cells(R, 7), "NM", "email2", 0)  ' email
        Call IEInput(ie, Cells(R, 26), "NM", "subdomain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "directory", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "password", 0) ' media-login-pw
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Call IEClick(ie, "AT", "input", , "value", "Å@ìoÅ@ò^Å@")
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg050_2nd(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)   ' email
        Call IEInput(ie, Cells(R, 32), "NM", "password", 0) ' media-login-pw
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26) & "@" & Cells(R, 31), "NM", "site", 0)
        Call IEClick(ie, "AT", "input", , "value", "ï“èW")
        Call IEWait(ie)
        Call IEClick(ie, "TG", "a", 1)
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 14), "NM", "category_id", 0)  ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "page_title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 16), "NM", "keywords", 0)  ' media-reg-keyword
        Call IEInput(ie, Cells(R, 17), "NM", "description", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 32), "NM", "password", 0) ' media-login-pw
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Call IEClick(ie, "TG", "a", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("TG", "a", 0)).href   ' media-url
        Cells(R, 29) = "http://frispe-hp.com/login/" ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg050_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg051_1st(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 19), "NM", "name", 0) ' media-reg-name
        Call IEInput(ie, "t", "NM", "sei_flg", 0) ' media-reg-sex: [f] female [t] male
        Call IEInput(ie, Cells(R, 24), "NM", "todoufuken", 0)   ' media-reg-area
        Call IEInput(ie, Cells(R, 20), "NM", "birth_year", 0)   ' media-reg-birth-m
        Call IEInput(ie, Cells(R, 21), "NM", "birth_month", 0)   ' media-reg-birth-m
        Call IEInput(ie, Cells(R, 22), "NM", "birth_day", 0)   ' media-reg-birth-d
        Call IEInput(ie, Cells(R, 7), "NM", "email", 1)   ' email
        Call IEInput(ie, Cells(R, 7), "NM", "email2", 0)  ' email
        Call IEInput(ie, Cells(R, 32), "NM", "password", 1) ' media-login-pw
        Call IEInput(ie, Cells(R, 31), "NM", "login_id", 0) ' media-login-id
        Call IEClick(ie, "AT", "input", , "alt", "ìoò^")
        Call IEWait(ie)
        Call IEClick(ie, "AT", "input", , "value", "ìoò^")
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg051_2nd(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEInput(ie, Cells(R, 32), "NM", "password", 0) ' media-login-pw
        Call IEClick(ie, "AT", "input", , "type", "image")
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 31), "NM", "login_id", 0)
        Call IEClick(ie, "AT", "input", , "value", "ï“èW")
        Call IEWait(ie)
        Call IEClick(ie, "TG", "a", 1)
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 14), "NM", "category_id", 0)  ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "page_title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 16), "NM", "keywords", 0)  ' media-reg-keyword
        Call IEInput(ie, Cells(R, 17), "NM", "description", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 32), "NM", "password", 0) ' media-login-pw
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Call IEClick(ie, "TG", "a", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("TG", "a", 0)).href   ' media-url
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg051_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg052_1st(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 31), "NM", "user_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "user_pass", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEInput(ie, Cells(R, 17), "NM", "title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0)  ' media-reg-category-1
        Call IEInput(ie, IECaptcha, "NM", "phrase", 0)  ' captcha
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg052_2nd(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEInput(ie, Cells(R, 32), "NM", "password", 0) ' media-login-pw
        Call IEClick(ie, "AT", "input", , "type", "image")
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 31), "NM", "site", 0)
        Call IEClick(ie, "AT", "input", , "value", "ï“èW")
        Call IEWait(ie)
        Call IEClick(ie, "TG", "a", 1)
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 14), "NM", "category_id", 0)  ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "page_title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 16), "NM", "keywords", 0)  ' media-reg-keyword
        Call IEInput(ie, Cells(R, 17), "NM", "description", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 32), "NM", "password", 0) ' media-login-pw
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Call IEClick(ie, "TG", "a", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("TG", "a", 0)).href   ' media-url
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        'Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg052_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg053_1st(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEClick(ie, "NM", "genre", 0)
        Call IEClick(ie, "NM", "type", 0)
        Call IEInput(ie, Cells(R, 31), "NM", "login_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "password", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Call IEClick(ie, "AT", "input", , "value", "ëóêM")
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg053_2nd(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEGotoURL(ie, "http://unirank.biz/login.html")
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 31), "NM", "login_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "password", 0) ' media-login-pw
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("AT", "input", "type", "text")).Value ' media-url
        Cells(R, 29) = "http://unirank.biz/login.html"  ' media-login-url
        Call IEClick(ie, "TG", "a", 2)
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 18), "NM", "toppage_body2", 0)  ' media-reg-desc
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Cells(R, 4) = Date ' date-new
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg053_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg054_1st(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 7), "NM", "email", 1)   ' email
        Call IEInput(ie, Cells(R, 19), "NM", "name", 0) ' media-reg-name
        Call IEInput(ie, "t", "NM", "sei_flg", 0) ' media-reg-sex: [f] female [t] male
        Call IEInput(ie, Cells(R, 24), "NM", "todoufuken", 0)   ' media-reg-area
        Call IEInput(ie, Cells(R, 20), "NM", "birth_year", 0)   ' media-reg-birth-m
        Call IEInput(ie, Cells(R, 21), "NM", "birth_month", 0)   ' media-reg-birth-m
        Call IEInput(ie, Cells(R, 22), "NM", "birth_day", 0)   ' media-reg-birth-d
        Call IEInput(ie, Cells(R, 31), "NM", "login_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "password", 0) ' media-login-pw
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Call IEClick(ie, "AT", "input", , "value", "ëóêM")
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg054_2nd(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEGotoURL(ie, "http://zakkuzaku.net/login.html")
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 31), "NM", "login_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "password", 0) ' media-login-pw
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("AT", "input", "type", "text")).Value ' media-url
        Cells(R, 29) = "http://zakkuzaku.net/login.html"  ' media-login-url
        Call IEClick(ie, "TG", "a", 2)
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 18), "NM", "body1", 0)  ' media-reg-desc
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Cells(R, 4) = Date ' date-new
        Call ie.Quit
    Set ie = Nothing
End Sub


Sub BReg054_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg055_1st(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 7), "NM", "email", 1)   ' email
        Call IEInput(ie, Cells(R, 19), "NM", "name", 0) ' media-reg-name
        Call IEInput(ie, "t", "NM", "sei_flg", 0) ' media-reg-sex: [f] female [t] male
        Call IEInput(ie, Cells(R, 24), "NM", "todoufuken", 0)   ' media-reg-area
        Call IEInput(ie, Cells(R, 20), "NM", "birth_year", 0)   ' media-reg-birth-m
        Call IEInput(ie, Cells(R, 21), "NM", "birth_month", 0)   ' media-reg-birth-m
        Call IEInput(ie, Cells(R, 22), "NM", "birth_day", 0)   ' media-reg-birth-d
        Call IEInput(ie, Cells(R, 31), "NM", "login_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "password", 0) ' media-login-pw
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Call IEClick(ie, "AT", "input", , "value", "ëóêM")
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg055_2nd(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEGotoURL(ie, "http://unimo.biz/login.html")
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 31), "NM", "login_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "password", 0) ' media-login-pw
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("AT", "input", "type", "text")).Value ' media-url
        Cells(R, 29) = "http://unimo.biz/login.html"  ' media-login-url
        Call IEClick(ie, "TG", "a", 2)
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 18), "NM", "body1", 0)  ' media-reg-desc
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        Cells(R, 4) = Date ' date-new
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg055_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg056_1st(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 31), "NM", "user_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "user_pass", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEInput(ie, Cells(R, 17), "NM", "title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0)  ' media-reg-category-1
        Call IEInput(ie, IECaptcha, "NM", "phrase", 0)  ' captcha
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg056_2nd(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0)  ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).href ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 1)).Click
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub


Sub BReg056_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg057_1st(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 31), "NM", "user_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "user_pass", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEInput(ie, Cells(R, 17), "NM", "title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0)  ' media-reg-category-1
        Call IEInput(ie, IECaptcha, "NM", "phrase", 0)  ' captcha
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg057_2nd(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0)  ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).href ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 1)).Click
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub


Sub BReg057_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg058_1st(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 31), "NM", "user_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "user_pass", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEInput(ie, Cells(R, 17), "NM", "title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0)  ' media-reg-category-1
        Call IEInput(ie, IECaptcha, "NM", "phrase", 0)  ' captcha
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg058_2nd(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0)  ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).href ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 1)).Click
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub


Sub BReg058_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg059_1st(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 31), "NM", "user_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "user_pass", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEInput(ie, Cells(R, 17), "NM", "title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0)  ' media-reg-category-1
        Call IEInput(ie, IECaptcha, "NM", "phrase", 0)  ' captcha
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg059_2nd(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0)  ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).href ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 1)).Click
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub


Sub BReg059_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg060_1st(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 31), "NM", "user_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "user_pass", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEInput(ie, Cells(R, 17), "NM", "title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0)  ' media-reg-category-1
        Call IEInput(ie, IECaptcha, "NM", "phrase", 0)  ' captcha
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg060_2nd(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0)  ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).href ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 1)).Click
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub


Sub BReg060_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg061_1st(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 31), "NM", "user_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "user_pass", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEInput(ie, Cells(R, 17), "NM", "title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0)  ' media-reg-category-1
        Call IEInput(ie, IECaptcha, "NM", "phrase", 0)  ' captcha
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg061_2nd(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0)  ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).href ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 1)).Click
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub


Sub BReg061_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg062_1st(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 31), "NM", "user_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "user_pass", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEInput(ie, Cells(R, 17), "NM", "title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0)  ' media-reg-category-1
        Call IEInput(ie, IECaptcha, "NM", "phrase", 0)  ' captcha
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg062_2nd(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0)  ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).href ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 1)).Click
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub


Sub BReg062_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg063_1st(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 31), "NM", "user_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "user_pass", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEInput(ie, Cells(R, 17), "NM", "title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0)  ' media-reg-category-1
        Call IEInput(ie, IECaptcha, "NM", "phrase", 0)  ' captcha
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg063_2nd(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0)  ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).href ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 1)).Click
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub


Sub BReg063_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg064_1st(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 31), "NM", "user_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "user_pass", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEInput(ie, Cells(R, 17), "NM", "title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0)  ' media-reg-category-1
        Call IEInput(ie, IECaptcha, "NM", "phrase", 0)  ' captcha
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg064_2nd(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0)  ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).href ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 1)).Click
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub


Sub BReg064_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg065_1st(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 31), "NM", "user_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "user_pass", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEInput(ie, Cells(R, 17), "NM", "title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0)  ' media-reg-category-1
        Call IEInput(ie, IECaptcha, "NM", "phrase", 0)  ' captcha
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg065_2nd(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0)  ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).href ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 1)).Click
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub


Sub BReg065_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg066_1st(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 31), "NM", "user_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "user_pass", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEInput(ie, Cells(R, 17), "NM", "title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0)  ' media-reg-category-1
        Call IEInput(ie, IECaptcha, "NM", "phrase", 0)  ' captcha
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg066_2nd(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0)  ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).href ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 1)).Click
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub


Sub BReg066_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg067_1st(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 31), "NM", "user_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "user_pass", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEInput(ie, Cells(R, 17), "NM", "title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0)  ' media-reg-category-1
        Call IEInput(ie, IECaptcha, "NM", "phrase", 0)  ' captcha
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg067_2nd(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0)  ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).href ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 1)).Click
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub


Sub BReg067_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg068_1st(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 31), "NM", "user_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "user_pass", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEInput(ie, Cells(R, 17), "NM", "title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0)  ' media-reg-category-1
        Call IEInput(ie, IECaptcha, "NM", "phrase", 0)  ' captcha
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg068_2nd(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0)  ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).href ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 1)).Click
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub


Sub BReg068_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg069_1st(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 31), "NM", "user_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "user_pass", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEInput(ie, Cells(R, 17), "NM", "title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0)  ' media-reg-category-1
        Call IEInput(ie, IECaptcha, "NM", "phrase", 0)  ' captcha
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg069_2nd(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0)  ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).href ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 1)).Click
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub


Sub BReg069_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg070_1st(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 12))    ' media-reg-url
        Call IEInput(ie, Cells(R, 31), "NM", "user_id", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 32), "NM", "user_pass", 0) ' media-login-pw
        Call IEInput(ie, Cells(R, 7), "NM", "email", 0)    ' email
        Call IEInput(ie, Cells(R, 17), "NM", "title", 0)  ' media-reg-title
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0)  ' media-reg-category-1
        Call IEInput(ie, IECaptcha, "NM", "phrase", 0)  ' captcha
        Call IEClick(ie, "AT", "input", , "type", "submit")
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub

Sub BReg070_2nd(rng)    '
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0)  ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "a", 0)).href ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "a", 1)).Click
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        'Call ie.Quit
    Set ie = Nothing
End Sub


Sub BReg070_3rd(rng)
    Dim ie As Object
    Dim R As Integer
    
    R = rng.Row
    Set ie = IENew(Cells(R, 13))    ' media-reg-url-2
        Call IEDOM(ie, Array("ID", "main_con_txt2", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEDOM(ie, Array("ID", "left", "AT", "div", "class", "menu_contents", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call IEInput(ie, Cells(R, 26), "NM", "domain", 0) ' media-reg-domain
        Call IEInput(ie, Cells(R, 31), "NM", "dir_name", 0) ' media-login-id
        Call IEInput(ie, Cells(R, 14), "NM", "category", 0) ' media-reg-category-1
        Call IEInput(ie, Cells(R, 17), "NM", "name", 0)  ' media-reg-title
        Call IEFocus(ie, "NM", "domain", 0)
        Call Application.SendKeys("{Tab}")
        Call IEWait(ie, 1500)
        Call IEClick(ie, "NM", "submit", 0)
        Call IEWait(ie)
        Cells(R, 30) = IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 0)).href   ' media-url
        Call IEDOM(ie, Array("ID", "right_content", "TG", "table", 0, "TG", "a", 1)).Click
        Call IEWait(ie)
        Cells(R, 29) = ie.document.url ' media-login-url
        Cells(R, 4) = Date ' date-new
        Call IEDOM(ie, Array("ID", "header_menu_content", "TG", "a", 0)).Click
        Call IEWait(ie)
        Call ie.Quit
    Set ie = Nothing
End Sub

Sub zzz()  '
    Dim rng As Range
    Dim ie As Object
    Dim R As Integer
    
    On Error Resume Next
        For Each rng In Selection
            R = rng.Row
        
            Set ie = IENew(Cells(R, 1))
                Cells(R, 2) = IEDOM(ie, Array("ID", "article", "TG", "a", 3)).href
                Cells(R, 3) = IEDOM(ie, Array("ID", "article", "TG", "a", 4)).href
                Cells(R, 4) = IEDOM(ie, Array("ID", "article", "TG", "a", 5)).href
                Cells(R, 5) = IEDOM(ie, Array("ID", "article", "TG", "a", 6)).href
                Cells(R, 6) = IEDOM(ie, Array("ID", "article", "TG", "a", 7)).href
                Cells(R, 7) = IEDOM(ie, Array("ID", "article", "TG", "a", 8)).href
                Cells(R, 8) = IEDOM(ie, Array("ID", "article", "TG", "a", 9)).href
                Cells(R, 9) = IEDOM(ie, Array("ID", "article", "TG", "a", 10)).href
                Cells(R, 10) = IEDOM(ie, Array("ID", "article", "TG", "a", 11)).href
                Cells(R, 11) = IEDOM(ie, Array("ID", "article", "TG", "a", 12)).href
                Cells(R, 12) = IEDOM(ie, Array("ID", "article", "TG", "a", 13)).href
                Call ie.Quit
            Set ie = Nothing
        Next
    On Error GoTo 0
End Sub

