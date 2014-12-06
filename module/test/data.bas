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
    
    strData(1) = "�܏\��,�փ{�������[�}��,������"
    strData(2) = "��,a,a"
    strData(3) = "��,i,i"
    strData(4) = "��,u,u"
    strData(5) = "��,e,e"
    strData(6) = "��,o,o"
    strData(7) = "��,ka,k"
    strData(8) = "��,ki,k"
    strData(9) = "��,ku,k"
    strData(10) = "��,ke,k"
    strData(11) = "��,ko,k"
    strData(12) = "��,sa,s"
    strData(13) = "��,shi,s"
    strData(14) = "��,su,s"
    strData(15) = "��,se,s"
    strData(16) = "��,so,s"
    strData(17) = "��,ta,t"
    strData(18) = "��,chi,c"
    strData(19) = "��,tsu,t"
    strData(20) = "��,te,t"
    strData(21) = "��,to,t"
    strData(22) = "��,na,n"
    strData(23) = "��,ni,n"
    strData(24) = "��,nu,n"
    strData(25) = "��,ne,n"
    strData(26) = "��,no,n"
    strData(27) = "��,ha,h"
    strData(28) = "��,hi,h"
    strData(29) = "��,fu,f"
    strData(30) = "��,he,h"
    strData(31) = "��,ho,h"
    strData(32) = "��,ma,m"
    strData(33) = "��,mi,m"
    strData(34) = "��,mu,m"
    strData(35) = "��,me,m"
    strData(36) = "��,mo,m"
    strData(37) = "��,ya,y"
    strData(38) = "��,yu,y"
    strData(39) = "��,yo,y"
    strData(40) = "��,ra,r"
    strData(41) = "��,ri,r"
    strData(42) = "��,ru,r"
    strData(43) = "��,re,r"
    strData(44) = "��,ro,r"
    strData(45) = "��,wa,w"
    strData(46) = "��,wo,w"
    strData(47) = "��,n,n"
    
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
        tail_part = "�R"
    ElseIf random_value < 0.5 Then
        tail_part = "��"
    ElseIf random_value < 0.75 Then
        tail_part = "�c"
    Else
        tail_part = "��"
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
            tail_part = "�j"
        ElseIf random_value < 0.5 Then
            tail_part = "�l"
        ElseIf random_value < 0.75 Then
            tail_part = "�Y"
        Else
            tail_part = "�v"
        End If
    Else
        If random_value < 0.34 Then
            tail_part = "�q"
        ElseIf random_value < 0.67 Then
            tail_part = "��"
        Else
            tail_part = "��"
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


