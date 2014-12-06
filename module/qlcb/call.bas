Attribute VB_Name = "qlcb_call"
'<< __CommandQLCB__
'Case "cd": Call CellDeleteBlankRows
'Case "cv": Call CellFromAlignToValign
'Case "ci": Call CellInsert
'Case "cs": Call CellSort
'Case "ch": Call CellRowHeight
'Case "cw": Call CellClumnWidth
'Case "disp0": Call DisplayNone
'Case "disp9": Call DisplayNone2
'Case "disp1": Call DisplayAll
'Case "f": Call FormatStringSelection
'Case "img": Call ImageFromTextSelection
'Case "is": Call InitializeSheet
'Case "if": Call InitializeFont
'Case "lo": Call LinkOpenSelection
'Case "ld": Call LinkDelete
'Case "mail": Call MailBASPSelection
'Case "menu0": Call MenuHide
'Case "menu9": Call MenuHide2
'Case "menu1": Call MenuUnhide
'Case "m1": Call MoveBlock
'Case "m4": Call MoveBlock4
'Case "m5": Call MoveBlock5
'Case "m6": Call MoveBlock6
'Case "m2": Call MoveBlock2
'Case "sd": Call SheetDeleteEmpty
'Case "sr": Call SheetRename
'Case "ss", "split": Call StrSplit2
'Case "sls": Call SheetListName
'Case "sys": Call SystemListComponentAndProcedure
'Case "a0": Call QLCBHereDoc
'Case "a2": Call QLCBMain(pstrCom, hdcQLCBMTOS)
'Case "y", "yank": Call StrYank
'Case "yy" "yank2": Call StrYank2
'__CommandQLCB__
'<< __CommandQLCBMTOS__
'Commands:
'---------
'Case "q6": Call QLCBMTOSFillEntryID
'Case "q7": Call QLCBMTOSFillEntryBasename
'Case "q8": Call QLCBMTOSFillEntryDate
'Case "q9": Call QLCBMTOSFillEntryDateNow
'Case "q2": Call QLCBMain(pstrCom, hdcQLCB)
'__CommandQLCBMTOS__

Sub QLCBAlias( _
    Optional pstrCom _
    )
    On Error Resume Next
        Select Case pstrCom
            Case "a0": Call QLCBHereDoc
            Case "a1": Call QLCBMain(pstrCom, hdcQLCB)
            Case "a2": Call QLCBMain(pstrCom, hdcQLCBMTOS)
            Case "alias": Call SystemWhich("QLCBAlias")
            Case "af": Call Selection.AutoFilter
            Case "bs": Call BookShare
            Case "ca": Call CellShowAll
            Case "cc": Call CellCount
            Case "cd": Call CellDeleteBlankRows
            Case "ci": Call CellInsert
            Case "ch": Call CellRowHeight
            Case "cs": Call CellSort
            Case "cv": Call CellFromAlignToValign
            Case "cw": Call CellColumnWidth
            Case "count": Call CellCount
            Case "disp0": Call DisplayNone
            Case "disp9": Call DisplayNone2
            Case "disp1": Call DisplayAll
            Case "date": Call MsgBox(Now)
            Case "f": Call FormatStringSelection
            Case "fsplit": Call FormatStringSelection(, True)
            Case "img": Call ImageFromTextSelection
            Case "is": Call InitializeSheet
            Case "if": Call InitializeFont
            Case "lo": Call LinkOpenSelection
            Case "ld": Call LinkDelete
            Case "mail": Call MailBASPSelection
            Case "menu0": Call MenuHide
            Case "menu9": Call MenuHide2
            Case "menu1": Call MenuUnhide
            Case "m1": Call MoveBlock
            Case "m4": Call MoveBlock4
            Case "m5": Call MoveBlock5
            Case "m6": Call MoveBlock6
            Case "m2": Call MoveBlock2
            Case "ping": Call ShellCmd3("ping " & ActiveCell.Value)
            Case "pivot": Call AnalPivot
            Case "r": Call DisplayRefresh
            Case "rg", "reg": Call BReg
            Case "sd": Call SheetDeleteEmpty
            Case "sls": Call SheetListName
            Case "share": Call BookShare
            Case "shrink": Call CellShrink
            Case "sort": Call CellSort
            Case "sr": Call SheetRename
            Case "ss": Call StrSplit2
            Case "sc": Call SheetCopyToNewBook
            Case "split": Call StrSplit2
            Case "sys": Call SystemListMacro
            Case "time": Call MsgBox(Now)
            Case "which": Call SystemWhich(ActiveCell.Value)
            Case "whois": Call ShellZsh3("whois " & ActiveCell.Value)
            Case "y", "yank": Call StrYank
            Case "yy", "yank2": Call StrYank2
            Case "0111", "0100", "0010", "0110", "0101", "0011": Call AliasQLCBMTOSProcess(pstrCom)
            Case 1000 To 1999: Call AliasQLCBMTOSInit(pstrCom)
            Case "mt1": Call QLCBMTOSFillEntryID
            Case "mt2": Call QLCBMTOSFillEntryBasename
            Case "mt3": Call QLCBMTOSFillEntryDate
            Case "mt4": Call QLCBMTOSFillEntryDateNow
            Case Else:
                If RegexBASPMatch("/^w /", pstrCom) = 1 Then Call ShellCmd3(RegexBASPReplace("s/^w (.+?)$/$1/g", pstrCom))
                If RegexBASPMatch("/^x /", pstrCom) = 1 Then Call Application.Run(RegexBASPReplace("s/^x (.+?)$/$1/g", pstrCom))
                If RegexBASPMatch("/^z /", pstrCom) = 1 Then Call ShellZsh3(RegexBASPReplace("s/^z (.+?)$/$1/g", pstrCom))
                If RegexBASPMatch("/^sp /", pstrCom) = 1 Then Call _
                    AnalSumproduct( _
                        RegexBASPReplace("s/^sp (.+?),(.*?)$/$1/g", pstrCom), _
                        RegexBASPReplace("s/^sp (.+?),(.*?)$/$2/g", pstrCom))
                If RegexBASPMatch("/^sr /", pstrCom) = 1 Then Call SheetRename(RegexBASPReplace("s/^sr (.+?)$/$1/g", pstrCom))
                If RegexBASPMatch("/^ss /", pstrCom) = 1 Then Call StrSplit(RegexBASPReplace("s/^ss (.+?)$/$1/g", pstrCom))
                If RegexBASPMatch("/^split /", pstrCom) = 1 Then Call StrSplit(RegexBASPReplace("s/^split (.+?)$/$1/g", pstrCom))
                If RegexBASPMatch("/^ch /", pstrCom) = 1 Then Call CellRowHeight(RegexBASPReplace("s/^ch (.+?)$/$1/g", pstrCom))
                If RegexBASPMatch("/^cw /", pstrCom) = 1 Then Call CellColumnWidth(RegexBASPReplace("s/^cw (.+?)$/$1/g", pstrCom))
                If RegexBASPMatch("/^cs /", pstrCom) = 1 Then Call CellSelect(RegexBASPReplace("s/^cs (.+?)$/$1/g", pstrCom))
                If RegexBASPMatch("/^cu /", pstrCom) = 1 Then Call CellUnique(RegexBASPReplace("s/^cu (.+?)$/$1/g", pstrCom))
                If RegexBASPMatch("/^uniq /", pstrCom) = 1 Then Call CellUnique(RegexBASPReplace("s/^uniq (.+?)$/$1/g", pstrCom))
                If RegexBASPMatch("/^ca /", pstrCom) = 1 Then Call Range(RegexBASPReplace("s/^ca (.+?)$/$1/g", pstrCom)).Activate
                If RegexBASPMatch("/^goto /", pstrCom) = 1 Then Call Range(RegexBASPReplace("s/^goto (.+?)$/$1/g", pstrCom)).Activate
                If RegexBASPMatch("/^ping /", pstrCom) = 1 Then Call ShellCmd3(RegexBASPReplace("s/^ping (.+?)$/ping $1/g", pstrCom))
                If RegexBASPMatch("/^which /", pstrCom) = 1 Then Call SystemWhich(RegexBASPReplace("s/^which (.+?)$/$1/g", pstrCom))
                If RegexBASPMatch("/^whois /", pstrCom) = 1 Then Call ShellZsh3(RegexBASPReplace("s/^whois (.+?)$/whois $1/g", pstrCom))
                If RegexBASPMatch("/^zm/", pstrCom) = 1 Then Call DisplayZoom(RegexBASPReplace("s/^zm (.+?)$/$1/g", pstrCom))
                If RegexBASPMatch("/^zoom/", pstrCom) = 1 Then Call DisplayZoom(RegexBASPReplace("s/^zoom (.+?)$/$1/g", pstrCom))
                If RegexBASPMatch("/^u /", pstrCom) = 1 Then Call QLCBRemoteUpload(pstrCom)
                If RegexBASPMatch("/^f /", pstrCom) = 1 Then Call FormatStringSelection(RegexBASPReplace("s/^f (.+?)$/$1/g", pstrCom))
                If RegexBASPMatch("/^fsplit /", pstrCom) = 1 Then Call FormatStringSelection(RegexBASPReplace("s/^fsplit (.+?)$/$1/g", pstrCom), True)
                If RegexBASPMatch("/^m /", pstrCom) = 1 Then Call _
                    MatchVlookup( _
                        RegexBASPReplace("s/^m (.+?),(.+?),(.+?)$/$1/g", pstrCom), _
                        RegexBASPReplace("s/^m (.+?),(.+?),(.+?)$/$2/g", pstrCom), _
                        RegexBASPReplace("s/^m (.+?),(.+?),(.+?)$/$3/g", pstrCom))
                If RegexBASPMatch("/^m1 /", pstrCom) = 1 Then Call _
                    MatchVlookup1_1( _
                        RegexBASPReplace("s/^m1 (.+?),(.+?)$/$1/g", pstrCom), _
                        RegexBASPReplace("s/^m1 (.+?),(.+?)$/$2/g", pstrCom))
                If RegexBASPMatch("/^mm /", pstrCom) = 1 Then Call _
                    MatchVlookup2( _
                        RegexBASPReplace("s/^mm (.+?),(.+?),(.*?)$/$1/g", pstrCom), _
                        RegexBASPReplace("s/^mm (.+?),(.+?),(.*?)$/$2/g", pstrCom), _
                        RegexBASPReplace("s/^mm (.+?),(.+?),(.*?)$/$3/g", pstrCom))
                If RegexBASPMatch("/^mmm /", pstrCom) = 1 Then Call _
                    MatchVlookup3( _
                        RegexBASPReplace("s/^mmm (.+?),(.+?)$/$1/g", pstrCom), _
                        RegexBASPReplace("s/^mmm (.+?),(.+?)$/$2/g", pstrCom))
                If RegexBASPMatch("/^mhwc /", pstrCom) = 1 Then Call _
                    MatchHWC( _
                        RegexBASPReplace("s/^mhwc (.+?),(.+?)$/$1/g", pstrCom), _
                        RegexBASPReplace("s/^mhwc (.+?),(.+?)$/$2/g", pstrCom))
                If RegexBASPMatch("/^merge /", pstrCom) = 1 Then Call _
                    CellMerge( _
                        RegexBASPReplace("s/^merge (.+?)$/$1/g", pstrCom))
                If RegexBASPMatch("/^cm /", pstrCom) = 1 Then Call _
                    CellMerge( _
                        RegexBASPReplace("s/^cm (.+?)$/$1/g", pstrCom))
                If RegexBASPMatch("/^dup /", pstrCom) = 1 Then Call CellIsDuplicated(RegexBASPReplace("s/^dup (.+?)$/$1/g", pstrCom))
                If RegexBASPMatch("/^md /", pstrCom) = 1 Then Call _
                    MatchDuplicated( _
                        RegexBASPReplace("s/^md (.+?),(.+?),(.+?)$/$1/g", pstrCom), _
                        RegexBASPReplace("s/^md (.+?),(.+?),(.+?)$/$2/g", pstrCom), _
                        RegexBASPReplace("s/^md (.+?),(.+?),(.+?)$/$3/g", pstrCom))
                If RegexBASPMatch("/^md2 /", pstrCom) = 1 Then Call _
                    MatchDuplicated2( _
                        RegexBASPReplace("s/^md2 (.+?),(.+?),(.+?)$/$1/g", pstrCom), _
                        RegexBASPReplace("s/^md2 (.+?),(.+?),(.+?)$/$2/g", pstrCom), _
                        RegexBASPReplace("s/^md2 (.+?),(.+?),(.+?)$/$3/g", pstrCom))
                If RegexBASPMatch("/^cp /", pstrCom) = 1 Then Call _
                    CellPrimaryKey( _
                        RegexBASPReplace("s/^cp (.+?),(.+?),(.*?)$/$1/g", pstrCom), _
                        RegexBASPReplace("s/^cp (.+?),(.+?),(.*?)$/$2/g", pstrCom), _
                        RegexBASPReplace("s/^cp (.+?),(.+?),(.*?)$/$3/g", pstrCom))
        End Select
    On Error GoTo 0
End Sub

Sub AliasQLCBMTOSProcess( _
    Optional pstrCom _
    )
    If SheetCheck(cstrWSName1) = False Then Call QLCBErr001
    Dim x As String
    Dim y As String
    Dim z As String
    
    x = RegexBASPReplace("s/.(.)../$1/g", pstrCom)
    z = RegexBASPReplace("s/...(.)/$1/g", pstrCom)
    y = RegexBASPReplace("s/..(.)./$1/g", pstrCom)
    If x = 1 Then
        Call QLCBMTOSImage
        If z = 1 Then
            Call QLCBMTOSImageUpload
        End If
        Call QLCBEnd(pstrCom)
    ElseIf y = 1 Then
        Call QLCBMTOSText(pstrCom)
        If z = 1 Then
            Call QLCBMTOSTextUpload
        End If
        Call QLCBEnd(pstrCom)
    End If
End Sub

Sub AliasQLCBMTOSInit( _
    Optional pstrCom _
    )
    Dim x As String
    
    x = ThisWorkbook.Path & "\tmpl\mtos\" & RegexBASPReplace("s/^.(.{3})/$1/g", pstrCom)
    If FileCheck2(x, 1) = False Then Exit Sub
    If FileAddData(x, cstrMediaNameCell, cstrWSName1, pstrCom) = 3 Then Exit Sub
    Call QLCBMTOSFillEntryBasename
    Call QLCBMTOSFillEntryDateNow
    Call QLCBMTOSFillEntryID
End Sub

