Attribute VB_Name = "object_excel_ui_border"
Sub BorderSquare()
    On Error Resume Next
        With Selection
            If _
                .Borders(xlEdgeLeft).LineStyle = xlNone _
                Or .Borders(xlEdgeTop).LineStyle = xlNone _
                Or .Borders(xlEdgeBottom).LineStyle = xlNone _
                Or .Borders(xlEdgeRight).LineStyle = xlNone _
            Then
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeRight).LineStyle = xlContinuous
            Else
                .Borders(xlEdgeLeft).LineStyle = xlNone
                .Borders(xlEdgeTop).LineStyle = xlNone
                .Borders(xlEdgeBottom).LineStyle = xlNone
                .Borders(xlEdgeRight).LineStyle = xlNone
            End If
        End With
    On Error GoTo 0
End Sub
    
Sub BorderSquareCross()
    If Selection.Borders.LineStyle = xlNone Then
        Selection.Borders.LineStyle = xlContinuous
    Else
        Selection.Borders.LineStyle = xlNone
    End If
End Sub
    
Sub BorderCross()
    On Error Resume Next
        With Selection
            If _
                .Borders(xlInsideHorizontal).LineStyle = xlNone _
                Or .Borders(xlInsideVertical).LineStyle = xlNone _
            Then
                .Borders(xlInsideHorizontal).LineStyle = xlContinuous
                .Borders(xlInsideVertical).LineStyle = xlContinuous
            Else
                .Borders(xlInsideHorizontal).LineStyle = xlNone
                .Borders(xlInsideVertical).LineStyle = xlNone
            End If
        End With
    On Error GoTo 0
End Sub
    
Sub BorderLeft()
    If Selection.Borders(xlEdgeLeft).LineStyle = xlNone Then
        Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Else
        Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    End If
End Sub
    
Sub BorderTop()
    If Selection.Borders(xlEdgeTop).LineStyle = xlNone Then
        Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Else
        Selection.Borders(xlEdgeTop).LineStyle = xlNone
    End If
End Sub
    
Sub BorderBottom()
    If Selection.Borders(xlEdgeBottom).LineStyle = xlNone Then
        Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Else
        Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    End If
End Sub
    
Sub BorderRight()
    If Selection.Borders(xlEdgeRight).LineStyle = xlNone Then
        Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Else
        Selection.Borders(xlEdgeRight).LineStyle = xlNone
    End If
End Sub

Sub BorderDiagonalUp()
    If Selection.Borders(xlDiagonalUp).LineStyle = xlNone Then
        Selection.Borders(xlDiagonalUp).LineStyle = xlContinuous
    Else
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    End If
End Sub
    
Sub BorderDiagonalDown()
    If Selection.Borders(xlDiagonalDown).LineStyle = xlNone Then
        Selection.Borders(xlDiagonalDown).LineStyle = xlContinuous
    Else
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    End If
End Sub

