Attribute VB_Name = "object_time"
Sub TimeWait( _
    pstrWaittime _
    )
    Dim cnt
    
    cnt = timeGetTime()
    Do While timeGetTime() - cnt < pstrWaittime
        DoEvents
    Loop
End Sub

Function TimeGetDate( _
    pstrFormat _
    )
    Select Case pstrFormat
        Case 1
            TimeGetDate = Right(Year(Now()), 2)
            TimeGetDate = TimeGetDate & Right("0" & Month(Now()), 2)
            TimeGetDate = TimeGetDate & Right("0" & Day(Now()), 2)
            TimeGetDate = TimeGetDate & Right("0" & Hour(Now()), 2)
            TimeGetDate = TimeGetDate & Right("0" & Minute(Now()), 2)
            TimeGetDate = TimeGetDate & Right("0" & Second(Now()), 2)
        Case 2
            TimeGetDate = Right(Year(Now()), 4)
            TimeGetDate = TimeGetDate & Right("0" & Month(Now()), 2)
            TimeGetDate = TimeGetDate & Right("0" & Day(Now()), 2)
        Case Else
    End Select
End Function



