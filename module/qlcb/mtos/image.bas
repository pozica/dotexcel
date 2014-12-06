Attribute VB_Name = "qlcb_mtos_image"
Sub QLCBMTOSImage()
    If WorksheetFunction.IsError(Range(cstrMediaNameCell)) = True Then Call QLCBErr003
    Dim strTmpFilePath As String
    Dim strTmpFileRelativePath As String
    Dim strMediaName As String
    Dim strMediaNames As Variant
    Dim intMediaNamesTimes As Integer
    Dim strEntryBasename As String
    Dim arrImageNameCell As String
    Dim arrImageName As Variant
    Dim arrImageNames As Variant
    Dim arrImageNamesTimes As Integer
    Dim strImageName As String
    Dim strImageNames As Variant
    Dim intImageStat As Integer
    Dim strImageInputName As String
    Dim strImageOutputName As String
    Dim strImageFirst As String
    Dim strImageSecond As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim strNow As String
    
    strNow = TimeGetDate(1)
    strTmpFilePath = ActiveWorkbook.Path & "\_tmp"
    If SheetCheck(cstrWSName1) = True Then strEntryBasename = Range(cstrEntryBasenameCell)
    strMediaName = Range(cstrMediaNameCell)
    strMediaNames = Split(strMediaName, ",")
    intMediaNamesTimes = UBound(Split(strMediaName, ","))
    For i = 0 To intMediaNamesTimes
        strMediaName = strMediaNames(i)
        If strMediaName <> "" Then
            strTmpFilePath = strTmpFilePath & "\" & "img_" & strMediaName & "_" & strEntryBasename & "_" & strNow
            Call FileCheck2(strTmpFilePath)
            strTmpFileRelativePath = cstrTmpFileRelativePath & "\" & "img_" & strMediaName & "_" & strEntryBasename & "_" & strNow
            k = carrImageNameCell2 + i
            arrImageNameCell = carrImageNameCell1 & k
            arrImageNameCell = Application.ConvertFormula(arrImageNameCell, xlR1C1, xlA1)
            arrImageName = Range(arrImageNameCell)
        
            ''' B-1-2. Parse array of image elements
            arrImageNames = Split(arrImageName, ";")
            arrImageNamesTimes = UBound(Split(arrImageName, ";"))
            For j = 0 To arrImageNamesTimes
                strImageName = arrImageNames(j)
                
                '''' B-1-3. Parse sub-array of image elements
                strImageNames = Split(strImageName, ",")
                intImageStat = strImageNames(0)
                
                '''' B-1-4. Process images
                If intImageStat = 1 Then
                    strImageInputName = strImageNames(1)
                    strImageOutputName = strTmpFileRelativePath & "\" & strImageNames(2) & strEntryBasename & strImageNames(3)
                    strImageFirst = strImageNames(4)
                    strImageSecond = strImageNames(5)
                    Call ImageResize(strImageInputName, strImageOutputName, strImageFirst, strImageSecond, ActiveWorkbook.Path)
                ElseIf intImageStat = 2 Then
                    strImageInputName = Range(strImageNames(1))
                    Call FileMake("UTF-8N", strImageInputName, ActiveWorkbook.Path & "\temp")
                    strImageInputName = "temp"
                    strImageOutputName = strTmpFileRelativePath & "\" & strImageNames(2) & strEntryBasename & strImageNames(3)
                    strImageFirst = ThisWorkbook.Path & "\tmpl\mtos\" & strMediaName & "\" & strImageNames(4)
                    strImageSecond = strImageNames(5)
                    Call ImageFromText(strImageInputName, strImageOutputName, strImageFirst, strImageSecond, ActiveWorkbook.Path)
                ElseIf intImageStat = 3 Then
                    strImageFirst = ThisWorkbook.Path & "\tmpl\mtos\" & strMediaName & "\" & strImageNames(1)
                    strImageOutputName = strTmpFileRelativePath & "\" & strImageNames(2) & strEntryBasename & strImageNames(3)
                    Call ImageCopy(strImageFirst, strImageOutputName, ActiveWorkbook.Path)
                End If
            Next
        End If
    Next
End Sub
    
Sub QLCBMTOSImageUpload()
    If WorksheetFunction.IsError(Range(cstrMediaNameCell)) = True Then Call QLCBErr003
    Dim strSetPutimgPath As String
    
    strSetPutimgPath = FileRead(ThisWorkbook.Path & "\tmpl\mtos\" & Range(cstrMediaNameCell) & "\putimg", 1)
    Select Case UBound(Split(strMediaName, ","))
        Case 0 To 10
            Call TimeWait(2000)
        Case 11 To 15
            Call TimeWait(3000)
        Case 16 To 20
            Call TimeWait(4000)
        Case 21 To 25
            Call TimeWait(5000)
        Case 26 To 30
            Call TimeWait(6000)
        Case 31 To 35
            Call TimeWait(7000)
        Case 36 To 40
            Call TimeWait(8000)
        Case 41 To 45
            Call TimeWait(9000)
        Case Else
            Call TimeWait(20000)
    End Select
    Call RemoteUpload(ActiveWorkbook.Path & "\_tmp\*", strSetPutimgPath)
End Sub



