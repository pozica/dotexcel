Attribute VB_Name = "util_excel_keybind"
Public cblnEmacsModeViewMode As Boolean

Sub Keybind()
    With Application
        .OnKey "^{u}", "SheetMovePrevious"
        .OnKey "^{i}", "SheetMoveNext"
        .OnKey "^{q}", "SheetDeleteExclusion"
        .OnKey "^{t}", "SheetAdd"
        .OnKey "%{t}", "SheetCopy"
        .OnKey "^%{t}", "SheetCopyToOtherBook"
        .OnKey "^{=}", "CellSelectFunction"
        .OnKey "%{q}", "Initialize"
        .OnKey "+%{q}", "InitializeDisplay"
        .OnKey "+^%{q}", "InitializeDisplay2"
        .OnKey "%{x}", "QLCBMain"
        .OnKey "^{c}", "KeybindCcMode"
        .OnKey "{i}"
        .OnKey "{o}"
        .OnKey "{n}"
        .OnKey "{t}"
        .OnKey "{l}"
        .OnKey "{r}"
        .OnKey "{b}"
        .OnKey "{u}"
        .OnKey "{d}"
    End With
    Call EmacsMode
End Sub

Sub KeybindCcMode()
    With Application
        .OnKey "{i}", "KeybindCcModeBorderCross"
        .OnKey "{o}", "KeybindCcModeBorderSquare"
        .OnKey "{n}", "KeybindCcModeBorderSquareCross"
        .OnKey "{t}", "KeybindCcModeBorderTop"
        .OnKey "{l}", "KeybindCcModeBorderLeft"
        .OnKey "{r}", "KeybindCcModeBorderRight"
        .OnKey "{b}", "KeybindCcModeBorderBottom"
        .OnKey "{u}", "KeybindCcModeBorderDiagonalUp"
        .OnKey "{d}", "KeybindCcModeBorderDiagonalDown"
    End With
End Sub

' C-c i
Sub KeybindCcModeBorderCross()
    Call BorderCross
    Call Keybind
End Sub

' C-c o
Sub KeybindCcModeBorderSquare()
    Call BorderSquare
    Call Keybind
End Sub

' C-c n
Sub KeybindCcModeBorderSquareCross()
    Call BorderSquareCross
    Call Keybind
End Sub

' C-c t
Sub KeybindCcModeBorderTop()
    Call BorderTop
    Call Keybind
End Sub

' C-c l
Sub KeybindCcModeBorderLeft()
    Call BorderLeft
    Call Keybind
End Sub

' C-c r
Sub KeybindCcModeBorderRight()
    Call BorderRight
    Call Keybind
End Sub

' C-c b
Sub KeybindCcModeBorderBottom()
    Call BorderBottom
    Call Keybind
End Sub

' C-c u
Sub KeybindCcModeBorderDiagonalUp()
    Call BorderDiagonalUp
    Call Keybind
End Sub

' C-c b
Sub KeybindCcModeBorderDiagonalDown()
    Call BorderDiagonalDown
    Call Keybind
End Sub

Sub EmacsMode()
    With Application
        .OnKey "^{o}", "EmacsModeF2"
        .OnKey "^{m}", "EmacsModeReturn"
        .OnKey "^{f}", "EmacsModeForwardCell"
        .OnKey "^{b}", "EmacsModeBackwardCell"
        .OnKey "^{p}", "EmacsModePreviousLine"
        .OnKey "^{n}", "EmacsModeNextLine"
        .OnKey "%{p}", "EmacsModeXlUp"
        .OnKey "%{n}", "EmacsModeXlDown"
        .OnKey "%{f}", "EmacsModeXlToRight"
        .OnKey "%{b}", "EmacsModeXlToLeft"
        .OnKey "%{<}", "EmacsModeBeginningOfUsedRangeRow"
        .OnKey "%{>}", "EmacsModeEndOfUsedRangeRow"
        .OnKey "+^{f}", "EmacsModeForwardCellSelect"
        .OnKey "+^{b}", "EmacsModeBackwardCellSelect"
        .OnKey "+^{p}", "EmacsModePreviousLineSelect"
        .OnKey "+^{n}", "EmacsModeNextLineSelect"
        .OnKey "+%{p}", "EmacsModeXlUpSelect"
        .OnKey "+%{n}", "EmacsModeXlDownSelect"
        .OnKey "+%{f}", "EmacsModeXlToRightSelect"
        .OnKey "+%{b}", "EmacsModeXlToLeftSelect"
        .OnKey "^%{<}", "EmacsModeBeginningOfUsedRangeRowSelect"
        .OnKey "^%{>}", "EmacsModeEndOfUsedRangeRowSelect"
        .OnKey "^%{p}", "EmacsModeSheetPrevious"
        .OnKey "^%{n}", "EmacsModeSheetNext"
        .OnKey "+^%{p}", "EmacsModeSheetPrevious2"
        .OnKey "+^%{n}", "EmacsModeSheetNext2"
        .OnKey "^{/}", "EmacsModeUndo"
        .OnKey "^{_}", "EmacsModeUndo"
        .OnKey "^{g}", "EmacsModeKeyboardQuit"
        .OnKey "^{y}", "EmacsModeYank"
        .OnKey "^{w}", "EmacsModeKillRegion"
        .OnKey "%{w}", "EmacsModeKillRingSave"
        .OnKey "%{d}", "EmacsModeKillCurrentRegion"
        .OnKey "^{k}", "EmacsModeKillLine"
        .OnKey "^{j}", "EmacsModeKillVerticalLine"
        .OnKey "^{h}", "EmacsModeDeleteBackwardChar"
        .OnKey "^{a}", "EmacsModeBeginningOfUsedRangeLine"
        .OnKey "^{e}", "EmacsModeEndOfUsedRangeLine"
        .OnKey "^{v}", "EmacsModeScrollUp"
        .OnKey "%{v}", "EmacsModeScrollDown"
        .OnKey "^{l}", "EmacsModeRecenter"
        .OnKey "^%{l}", "EmacsModeRebottom"
        .OnKey "^{s}", "EmacsModeSearch"
        .OnKey "%{g}", "EmacsModeGoto"
        .OnKey "^{x}", "EmacsModeCxMode"
        .OnKey "^{z}", "EmacsModeCzMode"
        .OnKey "%{s}", "EmacsModeViewMode"
        .OnKey "{SCROLLLOCK}", "EmacsModeViewMode"
        .OnKey "{2}"
        .OnKey "^{2}"
        .OnKey "{3}"
        .OnKey "+{s}"
        .OnKey "{(}"
        .OnKey "{)}"
        .OnKey "{b}"
        .OnKey "{c}"
        .OnKey "{k}"
        .OnKey "+{q}"
        .OnKey "+{v}"
        .OnKey "+^{u}"
        .OnKey "+^{i}"
        .OnKey "{l}"
        .OnKey "{h}"
        .OnKey "{k}"
        .OnKey "{j}"
        .OnKey "{p}"
        .OnKey "{n}"
        .OnKey "+{l}"
        .OnKey "+{h}"
        .OnKey "+{k}"
        .OnKey "+{j}"
        .OnKey "+{p}"
        .OnKey "+{n}"
        .OnKey "+^{l}"
        .OnKey "+^{h}"
        .OnKey "+^{k}"
        .OnKey "+^{j}"
        .OnKey "{p}"
        .OnKey "{n}"
        .OnKey "+{p}"
        .OnKey "+{n}"
        .OnKey "{ }"
        .OnKey "{u}"
        .OnKey "{y}"
        .OnKey "{b}"
        .OnKey "+{ESC}", "Enable_Keys"
    End With
End Sub

'' define EmacsMode commands
' C-o
Sub EmacsModeF2(): Application.SendKeys "{F2}": End Sub

' C-m
Sub EmacsModeReturn(): Application.SendKeys "{RETURN}": End Sub

' C-f
Sub EmacsModeForwardCell()
    Application.SendKeys "{RIGHT}"
End Sub

' C-b
Sub EmacsModeBackwardCell()
    If ActiveCell.Column <> 1 Then ActiveCell.Offset(0, -1).Select
End Sub

' C-p
Sub EmacsModePreviousLine()
    If ActiveCell.Row <> 1 Then ActiveCell.Offset(-1, 0).Select
End Sub

' C-n
Sub EmacsModeNextLine()
    Application.SendKeys "{DOWN}"
End Sub

' M-p
Sub EmacsModeXlUp(): Selection.End(xlUp).Select: End Sub

' M-n
Sub EmacsModeXlDown(): Selection.End(xlDown).Select: End Sub

' M-f
Sub EmacsModeXlToRight(): Selection.End(xlToRight).Select: End Sub

' M-b
Sub EmacsModeXlToLeft(): Selection.End(xlToLeft).Select: End Sub

' M-<
Sub EmacsModeBeginningOfUsedRangeRow(): Application.SendKeys "^{HOME}": End Sub

' M->
Sub EmacsModeEndOfUsedRangeRow(): ActiveCell.SpecialCells(xlLastCell).Select: End Sub

' S-C-f
Sub EmacsModeForwardCellSelect()
    If ActiveCell.MergeCells Then If ActiveCell.MergeArea.End(xlToRight).Column = Columns.count Then Exit Sub
    If ActiveCell.Column <> Columns.count Then
        Range( _
            Cells(Selection(1).Row, Selection(1).Column), _
            Cells(Selection(Selection.count).Row, Selection(Selection.count).Column).Offset(0, 1) _
        ).Select
    End If
End Sub

' S-C-b
Sub EmacsModeBackwardCellSelect()
    If ActiveCell.Column <> 1 Then
        Range( _
            Cells(Selection(1).Row, Selection(1).Column).Offset(0, -1), _
            Cells(Selection(Selection.count).Row, Selection(Selection.count).Column) _
        ).Select
    End If
End Sub

' S-C-p
Sub EmacsModePreviousLineSelect()
    If ActiveCell.Row <> 1 Then
        Range( _
            Cells(Selection(1).Row, Selection(1).Column).Offset(-1, 0), _
            Cells(Selection(Selection.count).Row, Selection(Selection.count).Column) _
        ).Select
    End If
End Sub

' S-C-n
Sub EmacsModeNextLineSelect()
    If ActiveCell.MergeCells Then If ActiveCell.MergeArea.End(xlDown).Row = Rows.count Then Exit Sub
    If ActiveCell.Row <> Rows.count Then
        Range( _
            Cells(Selection(1).Row, Selection(1).Column), _
            Cells(Selection(Selection.count).Row, Selection(Selection.count).Column).Offset(1, 0) _
        ).Select
    End If
End Sub

' S-M-f
Sub EmacsModeXlToRightSelect()
    Range( _
        Cells(Selection(1).Row, Selection(1).Column), _
        Cells(Selection(Selection.count).Row, Selection(Selection.count).Column).End(xlToRight) _
    ).Select
End Sub

' S-M-b
Sub EmacsModeXlToLeftSelect()
    Range( _
        Cells(Selection(1).Row, Selection(1).Column).End(xlToLeft), _
        Cells(Selection(Selection.count).Row, Selection(Selection.count).Column) _
    ).Select
End Sub

' S-M-p
Sub EmacsModeXlUpSelect()
    Range( _
        Cells(Selection(1).Row, Selection(1).Column).End(xlUp), _
        Cells(Selection(Selection.count).Row, Selection(Selection.count).Column) _
    ).Select
End Sub

' S-M-n
Sub EmacsModeXlDownSelect()
    Range( _
        Cells(Selection(1).Row, Selection(1).Column), _
        Cells(Selection(Selection.count).Row, Selection(Selection.count).Column).End(xlDown) _
    ).Select
End Sub

' C-M-<
Sub EmacsModeBeginningOfUsedRangeRowSelect()
    Range( _
        Cells(1, 1), _
        Cells(Selection(Selection.count).Row, Selection(Selection.count).Column) _
    ).Select
End Sub

' C-M->
Sub EmacsModeEndOfUsedRangeRowSelect()
    Range( _
        Cells(Selection(1).Row, Selection(1).Column), _
        Cells(ActiveCell.SpecialCells(xlLastCell).Row, ActiveCell.SpecialCells(xlLastCell).Column) _
    ).Select
End Sub

' C-u, C-M-p
Sub EmacsModeSheetPrevious()
    SheetPrevious
End Sub

' C-i, C-M-n
Sub EmacsModeSheetNext()
    SheetNext
End Sub

' S-C-u, S-C-M-p
Sub EmacsModeSheetPrevious2()
    SheetPrevious2
End Sub

' S-C-i, S-C-M-n
Sub EmacsModeSheetNext2()
    SheetNext2
End Sub

' C-g
Sub EmacsModeKeyboardQuit()
    Application.CutCopyMode = False
End Sub

' M-g
Sub EmacsModeGoto()
    SystemGoto (ActiveCell.Value)
End Sub

' C-y
Sub EmacsModeYank()
    StrYank
End Sub

' C-k
Sub EmacsModeKill()
    Selection.ClearContents
End Sub

' C-w
Sub EmacsModeKillRegion()
    StrKillRegion
End Sub

' M-w
Sub EmacsModeKillRingSave()
    StrKillRingSave
End Sub

' M-d
Sub EmacsModeKillCurrentRegion()
    StrKillCurrentRegion
End Sub

' C-k
Sub EmacsModeKillLine()
    StrKillLine
End Sub

' C-j
Sub EmacsModeKillVerticalLine()
    StrKillVerticalLine
End Sub

' C-h
Sub EmacsModeDeleteBackwardChar()
    StrDeleteBackwardChar
End Sub

' C-/ OR C-_
Sub EmacsModeUndo()
    On Error Resume Next
        Application.Undo
    On Error GoTo 0
End Sub

' C-a move to left of UsedRange
Sub EmacsModeBeginningOfUsedRangeLine()
    Cells(ActiveCell.Row, ActiveSheet.UsedRange.Column).Activate
End Sub

' C-e move to right of UsedRange
Sub EmacsModeEndOfUsedRangeLine()
    Cells(ActiveCell.Row, _
        ActiveSheet.UsedRange.Columns _
        (ActiveSheet.UsedRange.Columns.count).Column).Activate
End Sub

' C-v move forward with one window
Sub EmacsModeScrollUp()
    Dim RowNum As Long
    Dim ColNum As Long
    On Error Resume Next
        With ActiveWindow
             RowNum = .ActiveCell.Row - .VisibleRange.Row + 1
             ColNum = .ActiveCell.Column
             .LargeScroll Down:=1
             .VisibleRange.Cells(RowNum, ColNum).Activate
        End With
    On Error GoTo 0
End Sub

' C-z move backward with one window
Sub EmacsModeScrollDown()
    Dim RowNum As Long
    Dim ColNum As Long
    On Error Resume Next
        With ActiveWindow
            RowNum = .ActiveCell.Row - .VisibleRange.Row + 1
            ColNum = .ActiveCell.Column
            .LargeScroll up:=1
            .VisibleRange.Cells(RowNum, ColNum).Activate
        End With
    On Error GoTo 0
End Sub

' C-l centerize on display
Sub EmacsModeRecenter()
    Dim x As Long
    With ActiveWindow
        x = Int(ActiveCell.Row - (.VisibleRange.Height / ActiveCell.Height) / 2.5)
        If x > 0 Then
            .ScrollRow = x
        End If
    End With
End Sub
 
' C-M-l
Sub EmacsModeRebottom()
    Dim x As Long
    With ActiveWindow
        x = Int(ActiveCell.Row - (.VisibleRange.Height / ActiveCell.Height) / 1.1)
        If x > 0 Then
            .ScrollRow = x
        End If
    End With
End Sub
 
 ' C-s search dialogue
 Sub EmacsModeSearch()
    Call Application.CommandBars.FindControl(ID:=1849).Execute
 End Sub

'' define C-x mode key bindings
Sub EmacsModeCxMode()
    With Application
        .OnKey "{(}"
        .OnKey "{)}"
        .OnKey "{b}", "EmacsModeCxModeSheetAdd"
        .OnKey "{k}", "EmacsModeCxModeBookClose"
        .OnKey "^{s}", "EmacsModeCxModeSaveFile"
        .OnKey "^{w}", "EmacsModeCxModeWriteFile"
        .OnKey "^{f}", "EmacsModeCxModeFindFile"
        .OnKey "^{p}", "EmacsModeCxModePrintFile"
        .OnKey "^{g}", "EmacsMode"
        .OnKey "^{x}" ' cut
        .OnKey "^{v}" ' paste
        .OnKey "^{z}" ' undo
    End With
End Sub

' C-x b
Sub EmacsModeCxModeSheetAdd()
    Call SheetAdd
    Call EmacsMode
End Sub

' C-x k
Sub EmacsModeCxModeBookClose()
    Call SheetDelete
    Call EmacsMode
End Sub

'' define C-x mode commands
' C-x C-s save
Sub EmacsModeCxModeSaveFile()
    On Error Resume Next
        ActiveWorkbook.Save
        Call EmacsMode
    On Error GoTo 0
End Sub

' C-x C-w save with naming
Sub EmacsModeCxModeWriteFile()
    On Error Resume Next
        Application.Dialogs(xlDialogSaveAs).Show
        Call EmacsMode
    On Error GoTo 0
End Sub

' C-x C-f open file
Sub EmacsModeCxModeFindFile()
    Application.Dialogs(xlDialogOpen).Show
    Call EmacsMode
End Sub

' C-x C-p print dialogue
Sub EmacsModeCxModePrintFile()
    Application.Dialogs(xlDialogPrint).Show
    Call EmacsMode
End Sub

'' define C-z mode key bindings
Sub EmacsModeCzMode()
    With Application
        .OnKey "+{q}", "EmacsModeCzModeBookMaximized"
        .OnKey "+{s}", "EmacsModeCzModeSheetFreezePanes"
        .OnKey "+{v}", "EmacsModeCzModeBookTiled"
        .OnKey "{c}", "EmacsModeCzModeBookAdd"
        .OnKey "{k}", "EmacsModeCzModeBookClose"
    End With
End Sub

' C-z c
Sub EmacsModeCzModeBookAdd()
    Call BookAdd
    Call EmacsMode
End Sub

' C-z Q
Sub EmacsModeCzModeBookMaximized()
    Call BookMaximized
    Call EmacsMode
End Sub

' C-x S
Sub EmacsModeCzModeSheetFreezePanes()
    Call SheetFreezePanes
    Call EmacsMode
End Sub

' C-z V
Sub EmacsModeCzModeBookTiled()
    Call BookTiled
    Call EmacsMode
End Sub

' C-z k
Sub EmacsModeCzModeBookClose()
    Call BookClose
    Call EmacsMode
End Sub

'' define View mode key bindings
' M-s
Sub EmacsModeViewMode()
    Dim x As String
    Dim y As String
    
    If cblnEmacsModeViewMode = False Then
        With Application
            .OnKey "{l}", "EmacsModeViewModeScrollToRight"
            .OnKey "{h}", "EmacsModeViewModeScrollToLeft"
            .OnKey "{k}", "EmacsModeViewModeScrollDown"
            .OnKey "{j}", "EmacsModeViewModeScrollUp"
            .OnKey "^{l}", "EmacsModeForwardCell"
            .OnKey "^{h}", "EmacsModeBackwardCell"
            .OnKey "^{k}", "EmacsModePreviousLine"
            .OnKey "^{j}", "EmacsModeNextLine"
            .OnKey "+^{l}", "EmacsModeForwardCellSelect"
            .OnKey "+^{h}", "EmacsModeBackwardCellSelect"
            .OnKey "+^{k}", "EmacsModePreviousLineSelect"
            .OnKey "+^{j}", "EmacsModeNextLineSelect"
            .OnKey "{p}", "EmacsModeBeginningOfUsedRangeRow"
            .OnKey "{n}", "EmacsModeEndOfUsedRangeRow"
            .OnKey "+{p}", "EmacsModeBeginningOfUsedRangeRowSelect"
            .OnKey "+{n}", "EmacsModeEndOfUsedRangeRowSelect"
            .OnKey "{ }", "EmacsModeScrollUp"
            .OnKey "{u}", "EmacsModeScrollDown"
            .OnKey "{y}", "EmacsModeSheetPrevious"
            .OnKey "{b}", "EmacsModeSheetNext"
            .OnKey "{g}", "EmacsMode"
        End With
        Application.ReferenceStyle = xlR1C1
        cblnEmacsModeViewMode = True
    Else
        Application.ReferenceStyle = xlA1
        cblnEmacsModeViewMode = False
        Keybind
    End If
End Sub

' M-s j
Sub EmacsModeViewModeScrollDown()
    ActiveWindow.SmallScroll Down:=-10
End Sub

' M-s k
Sub EmacsModeViewModeScrollUp()
    ActiveWindow.SmallScroll up:=-10
End Sub

' M-s h
Sub EmacsModeViewModeScrollToLeft()
    ActiveWindow.SmallScroll ToLeft:=-10
End Sub

' M-s l
Sub EmacsModeViewModeScrollToRight()
    ActiveWindow.SmallScroll ToRight:=-10
End Sub

Sub Enable_Keys()
    Dim StartKeyCombination As Variant
    Dim KeysArray As Variant
    Dim Key As Variant
    Dim i As Long

    On Error Resume Next
        For Each StartKeyCombination In Array("+", "^", "%", "+^", "+%", "^%", "+^%")
            KeysArray = Array("{BS}", "{BREAK}", "{CAPSLOCK}", "{CLEAR}", "{DEL}", _
                        "{DOWN}", "{END}", "{ENTER}", "~", "{ESC}", "{HELP}", "{HOME}", _
                        "{INSERT}", "{LEFT}", "{NUMLOCK}", "{PGDN}", "{PGUP}", _
                        "{RETURN}", "{RIGHT}", "{SCROLLLOCK}", "{TAB}", "{UP}")
            For Each Key In KeysArray
                Application.OnKey StartKeyCombination & Key
            Next Key
            For i = 0 To 255
                Application.OnKey StartKeyCombination & Chr$(i)
            Next i
            For i = 1 To 15
                Application.OnKey StartKeyCombination & "{F" & i & "}"
            Next i
        Next StartKeyCombination
        For i = 1 To 15
            Application.OnKey "{F" & i & "}"
        Next i
        Application.OnKey "{PGDN}"
        Application.OnKey "{PGUP}"
End Sub

