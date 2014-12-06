Attribute VB_Name = "test_data"
Sub DataCreateNames()
    Dim i As Integer
    Dim j As Integer
    Dim x As String
    Dim y As String

    j = 100
    x = ActiveSheet.Name
    y = ActiveCell.Address
    
    Call DataCreateNameMakeRef
    For i = 0 To j - 1
        Call DataCreateName(Cells(Range(y).Row + i, Range(y).Column), x)
    Next i
    Sheets(x).Select: Cells(Range(y).Row + j + 2, Range(y).Column + 2).Select
    Application.DisplayAlerts = False: Sheets("hiragana").Delete: Application.DisplayAlerts = True
End Sub

Sub DataCreateName(index, target_sheet_name)
    sex_type = DataCreateNameGetRandomSexType()
    myouji = DataCreateNameGetRandomMyouji
    namae = DataCreateNameGetRandomNamae(sex_type)
    initial_char = DataCreateNameGetInitialChar(myouji)
    DataCreateNameWriteOnePerson index, target_sheet_name, myouji, namae, initial_char
End Sub

Sub DataCreateNameMakeRef()
    Dim strData(1 To 47) As String
    Dim varData As Variant
    Dim strData2(46, 2) As String
    
    strData(1) = "五十音,へボン式ローマ字,頭文字"
    strData(2) = "あ,a,a"
    strData(3) = "い,i,i"
    strData(4) = "う,u,u"
    strData(5) = "え,e,e"
    strData(6) = "お,o,o"
    strData(7) = "か,ka,k"
    strData(8) = "き,ki,k"
    strData(9) = "く,ku,k"
    strData(10) = "け,ke,k"
    strData(11) = "こ,ko,k"
    strData(12) = "さ,sa,s"
    strData(13) = "し,shi,s"
    strData(14) = "す,su,s"
    strData(15) = "せ,se,s"
    strData(16) = "そ,so,s"
    strData(17) = "た,ta,t"
    strData(18) = "ち,chi,c"
    strData(19) = "つ,tsu,t"
    strData(20) = "て,te,t"
    strData(21) = "と,to,t"
    strData(22) = "な,na,n"
    strData(23) = "に,ni,n"
    strData(24) = "ぬ,nu,n"
    strData(25) = "ね,ne,n"
    strData(26) = "の,no,n"
    strData(27) = "は,ha,h"
    strData(28) = "ひ,hi,h"
    strData(29) = "ふ,fu,f"
    strData(30) = "へ,he,h"
    strData(31) = "ほ,ho,h"
    strData(32) = "ま,ma,m"
    strData(33) = "み,mi,m"
    strData(34) = "む,mu,m"
    strData(35) = "め,me,m"
    strData(36) = "も,mo,m"
    strData(37) = "や,ya,y"
    strData(38) = "ゆ,yu,y"
    strData(39) = "よ,yo,y"
    strData(40) = "ら,ra,r"
    strData(41) = "り,ri,r"
    strData(42) = "る,ru,r"
    strData(43) = "れ,re,r"
    strData(44) = "ろ,ro,r"
    strData(45) = "わ,wa,w"
    strData(46) = "を,wo,w"
    strData(47) = "ん,n,n"
    
    For i = 1 To UBound(strData)
        varData = Split(strData(i), ",")
        For j = 0 To 2
            strData2(i - 1, j) = varData(j)
        Next
    Next
    Sheets.Add
    ActiveSheet.Name = "hiragana"
    Range("A1:C47") = strData2
End Sub

Function DataCreateNameGetRandomSexType()
    If Rnd >= 0.5 Then
        ret = "m"
    Else
        ret = "f"
    End If
    DataCreateNameGetRandomSexType = ret
End Function

Function DataCreateNameGetRandomMyouji()
    first_char = DataCreateNameGetRandomHiragana
    
    random_value = Rnd
    If random_value < 0.25 Then
        tail_part = "山"
    ElseIf random_value < 0.5 Then
        tail_part = "川"
    ElseIf random_value < 0.75 Then
        tail_part = "田"
    Else
        tail_part = "沢"
    End If
    
    myouji = first_char & tail_part
    
    DataCreateNameGetRandomMyouji = myouji
End Function

Function DataCreateNameGetRandomNamae(sex_type)
    first_char = DataCreateNameGetRandomHiragana
    second_char = DataCreateNameGetRandomHiragana
    
    random_value = Rnd
    If sex_type = "m" Then
        If random_value < 0.25 Then
            tail_part = "男"
        ElseIf random_value < 0.5 Then
            tail_part = "人"
        ElseIf random_value < 0.75 Then
            tail_part = "郎"
        Else
            tail_part = "夫"
        End If
    Else
        If random_value < 0.34 Then
            tail_part = "子"
        ElseIf random_value < 0.67 Then
            tail_part = "代"
        Else
            tail_part = "美"
        End If
    End If
    
    namae = first_char & second_char & tail_part
    
    DataCreateNameGetRandomNamae = namae
End Function

Function DataCreateNameGetRandomHiragana()
    min_index = 2
    max_index = 47
    
    random_index = Int((max_index - min_index + 1) * Rnd + min_index)
    
    ch = Sheets("hiragana").Cells(random_index, 1).Value
    
    DataCreateNameGetRandomHiragana = ch
End Function

Function DataCreateNameGetInitialChar(myouji)
    first_char = Left(myouji, 1)
    
    alphabet = Application.WorksheetFunction.VLookup( _
        first_char, _
        Sheets("hiragana").Range("A2:C47"), _
        3, False _
    )
    
    DataCreateNameGetInitialChar = alphabet
End Function

Sub DataCreateNameWriteOnePerson(index, target_sheet_name, myouji, namae, initial_char)
    Sheets(target_sheet_name).Cells(index.Row, index.Column) = initial_char
    Sheets(target_sheet_name).Cells(index.Row, index.Column + 1) = myouji
    Sheets(target_sheet_name).Cells(index.Row, index.Column + 2) = namae
End Sub


