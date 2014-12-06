Attribute VB_Name = "util_excel_analysis"
Function AnalSumproduct( _
    pstrNum As Long, _
    pstrSheet As String _
)
    Dim i As Long
    AnalSumproduct = "=SUMPRODUCT("
    For i = 0 To pstrNum - 2
        AnalSumproduct = AnalSumproduct & "(" & pstrSheet & "!RC[-" & pstrNum - i & "]:R[39998]C[-" & pstrNum - i & "]=RC[-" & pstrNum - i & "])*"
    Next i
    AnalSumproduct = AnalSumproduct & "(" & pstrSheet & "!RC[-" & pstrNum - i & "]:R[39998]C[-" & pstrNum - i & "]=RC[-" & pstrNum - i & "])"
    AnalSumproduct = AnalSumproduct & ")"
    ActiveCell.FormulaR1C1 = AnalSumproduct
End Function

Sub AnalPivot()
    Application.CommandBars.FindControl(ID:=2915).Execute
End Sub

Sub AnalPivotRefresh()
    Application.CommandBars.FindControl(ID:=459).Execute
End Sub
