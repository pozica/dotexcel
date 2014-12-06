Attribute VB_Name = "util_shell_cmd"
Sub SystemListMacro()
    Dim i
    Dim j
    Dim c
    Dim R
    Dim POL
    
    On Error Resume Next
        R = ActiveCell.Row
        c = ActiveCell.Column
        With Workbooks("init.xla")
            For i = 1 To .VBProject.VBComponents.count
                If .VBProject.VBComponents(i).Type = 1 Then
                    
                    With .VBProject.VBComponents(i).CodeModule
                        POL = ""
                        For j = 1 To .CountOfLines
                            If POL <> .ProcOfLine(j, 0) Then
                                POL = .ProcOfLine(j, 0)
                                Cells(R, c) = POL
                                R = R + 1
                            End If
                        Next
                    End With
                End If
            Next
            Call Cells(R - 1, c).CurrentRegion.Sort( _
                Key1:=Cells(Selection.CurrentRegion.Row, c), _
                Order1:=xlAscending, _
                Header:=xlGuess, _
                OrderCustom:=1, _
                MatchCase:=False, _
                Orientation:=xlTopToBottom, _
                SortMethod:=xlPinYin, _
                DataOption1:=xlSortNormal _
            )
            Cells(R + 2, c).Select
            EmacsModeRebottom
        End With
    On Error GoTo 0
End Sub

Sub SystemWhich( _
    Optional pstrMacro _
    )
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    Dim strData As String
    Dim varData As Variant

    On Error Resume Next
        If IsError(pstrMacro) Then pstrMacro = InputBox("Which Macro ?")
        With Workbooks("init.xla")
            For i = 1 To .VBProject.VBComponents.count
                With .VBProject.VBComponents(i).CodeModule
                    j = .ProcBodyLine(pstrMacro, 0)
                    If j <> 0 Then
                        k = .ProcCountLines(pstrMacro, 0)
                        strData = .Lines(j, k)
                        j = 0
                        k = 0
                    End If
                End With
            Next
            varData = Split(strData, vbCrLf)
            For l = 0 To UBound(varData)
                ActiveCell.Offset(l, 0).Value = varData(l)
            Next
            ActiveCell.Offset(UBound(varData) + 1, 0).Select
            EmacsModeRebottom
        End With
    On Error GoTo 0
End Sub

Sub SystemGoto( _
    Optional pstrRef _
    )
    On Error Resume Next
        If IsError(pstrRef) Or pstrRef = "" Then pstrRef = InputBox("Which reference ?")
        Application.GoTo Reference:=pstrRef
    On Error GoTo 0
End Sub


