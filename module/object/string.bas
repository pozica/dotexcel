Attribute VB_Name = "object_string"
Option Explicit

Function SplitPlus( _
    pstrCell, _
    pstrCode, _
    pstrAdd _
    )
    Dim StrSplit
    Dim i
    Dim strReturn
    
    StrSplit = Split(pstrCell, pstrCode)
    For i = 0 To UBound(StrSplit)
        strReturn = strReturn & pstrAdd & StrSplit(i)
    Next i
    
    SplitPlus = strReturn
End Function

Sub StrSplit( _
    Optional pstrSeparater = " " _
    )
    Dim blnTab As Boolean
    Dim blnSemicolon As Boolean
    Dim blnComma As Boolean
    Dim blnSpace As Boolean
    
    Select Case pstrSeparater
        Case "tab": blnTab = True
        Case ";": blnSemicolon = True
        Case ",": blnComma = True
        Case " ": blnSpace = True
    End Select
    Selection.TextToColumns _
        DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=True, _
        Tab:=blnTab, _
        Semicolon:=blnSemicolon, _
        Comma:=blnComma, _
        Space:=blnSpace, _
        Other:=False, _
        FieldInfo:=Array( _
            Array(1, 2), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
            Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1) _
        ), _
        TrailingMinusNumbers:=True
End Sub

Sub StrSplit2()
    Application.CommandBars.FindControl(ID:=806).Execute
End Sub

Function StrUnion( _
    pobjTarget As Excel.Range, _
    Optional pstrSep As String _
    ) As String
    Dim objCell As Excel.Range
    Dim strBuffer As String
    Dim lngPos As Long
    Dim lngAddLength As Long
    
    On Error Resume Next
        strBuffer = String(4096, 0)
        lngPos = 0
        
        For Each objCell In pobjTarget
            lngAddLength = Len(objCell.Text) + Len(pstrSep)
            If (Len(strBuffer) < lngPos + lngAddLength) Then
                strBuffer = strBuffer & String(((lngAddLength Mod 4096) + 1) * 4096, 0)
            End If
            Mid$(strBuffer, lngPos + 1, lngAddLength) = objCell.Text & pstrSep
            lngPos = lngPos + lngAddLength
        Next
           
        If (Len(pstrSep) > 0) Then
            StrUnion = Mid(strBuffer, 1, lngPos - Len(pstrSep))
        Else
            StrUnion = Mid(strBuffer, 1, lngPos)
        End If
End Function

Sub StrYank()
    On Error Resume Next
        Application.CommandBars.FindControl(ID:=755).Execute
    On Error GoTo 0
End Sub

Sub StrYank2()
    On Error Resume Next
        ActiveSheet.Paste
    On Error GoTo 0
End Sub

Sub StrKillRingSave()
    On Error Resume Next
        Selection.Copy
    On Error GoTo 0
End Sub

Sub StrKillRegion()
    Selection.Cut
End Sub

Sub StrKillCurrentRegion()
    Selection.CurrentRegion.Select
    Selection.Copy
End Sub

Sub StrKillLine()
    Dim x As String
    On Error Resume Next
        x = ActiveCell.Address
        Rows(Selection.Row).Select
        Selection.Delete Shift:=xlUp
        Range(x).Select
    On Error GoTo 0
End Sub

Sub StrKillVerticalLine()
    Dim x As String
    On Error Resume Next
        x = ActiveCell.Address
        Columns(Selection.Column).Select
        Selection.Delete Shift:=xlUp
        Range(x).Select
    On Error GoTo 0
End Sub

Sub StrDeleteBackwardChar()
    Selection.ClearContents
End Sub

Function StrRnd( _
    Optional pintLength As Integer = 8, _
    Optional pblnUpper As Boolean = False, _
    Optional pblnSign As Boolean = False _
    ) As String
    Dim pw As String
    Dim lw As String
    Dim up As String
    Dim no As String
    Dim sg As String
    Dim ch As String
    Dim intMax As Integer
    Dim i As Integer
    lw = "qwertyuiopasdfghjklzxcvbnm"
    up = "QWERTYUIOPASDFGHJKLZXCVBNM"
    no = "123456789"
    sg = "-^@[;:],./!""#$%&'()=~|`{+*}<>?_"
    ch = lw & no

    If pblnUpper = True Then ch = ch & up
    If pblnSign = True Then ch = ch & sg
    intMax = Len(ch)
    For i = 1 To pintLength: pw = Mid(ch, Int((intMax - 1 + 1) * Rnd + 1), 1) & pw: Next i
    StrRnd = pw
End Function
 
Function StrRepeat( _
    no As Integer, _
    ch As String _
    )
    StrRepeat = String(no, ch)
End Function

