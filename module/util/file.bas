Attribute VB_Name = "util_file"
Function FileCheck( _
    pstrFile As String _
    ) As Boolean
    With CreateObject("Scripting.FileSystemObject")
        If .FileExists(pstrFile) = True Then
            FileCheck = True
        Else
            FileCheck = False
        End If
    End With
End Function

Function FileCheck2( _
    pstrPath, _
    Optional pstrFlag = "" _
    ) As Boolean
    
    With CreateObject("Scripting.FileSystemObject")
        If .FolderExists(pstrPath) = False Then
            If pstrFlag <> 1 Then MkDir pstrPath
            If .FileExists(pstrPath) = False Then
                FileCheck2 = False
            Else
              FileCheck2 = True
            End If
        Else
            FileCheck2 = True
        End If
    End With
End Function

Function FileRead( _
    pstrFile, _
    Optional pstrFlag = "" _
    )
    Dim cnt
    Dim rc
    
    cnt = FreeFile
    Open pstrFile For Input As #cnt
    Do While Not EOF(cnt)
        Line Input #cnt, rc
        If RegexBASPMatch("/^#/", rc) = False Then
            If pstrFlag = 1 Then
                FileRead = rc
            Else
                FileRead = FileRead & vbCrLf & rc
            End If
        End If
    Loop
    Close #cnt
End Function

Sub FileMake( _
    pstrCharset, _
    pstrData, _
    pstrFilePath _
    )
    Dim fileNum As Long
    Dim inputd As Integer
    Dim buf As String
    Dim bytData() As Byte
    
    'for UTF-8N
    With CreateObject("ADODB.Stream")
        If pstrCharset = "UTF-8N" Then
            .Type = 2
            .Charset = "UTF-8"
            .Open
            .WriteText pstrData
            .SaveToFile pstrFilePath, 2
            .Close
            Call FileConvertUTF8N(pstrFilePath)
        Else
            .Type = 2
            .Charset = pstrCharset
            .Open
            .WriteText pstrData
            .SaveToFile pstrFilePath, 2
            .Close
        End If
    End With
End Sub

Sub FileConvertUTF8N( _
    pstrFilePath _
    )
    Dim strFilePath As String
    Dim objReadStream As Object
    Dim objWriteStream As Object
    Dim bytData() As Byte
    Const adTypeText = 2
    Const adTypeBinary = 1
    Const adReadLine = -2
    Const adWriteLine = 0
    Const adLF = 10
    Const adCRLF = -1
    Const adSaveCreateOverWrite = 2
    
    Set objReadStream = CreateObject("ADODB.Stream")
    Set objWriteStream = CreateObject("ADODB.Stream")
    With objReadStream
        .Open
        .Type = adTypeText
        .Charset = "UTF-8"
        .LoadFromFile pstrFilePath
    End With
    With objWriteStream
        .Open
        .Type = adTypeText
        .Charset = "UTF-8"
    End With
    'Convert a data per line
    Do Until objReadStream.EOS
        Call objWriteStream.WriteText(objReadStream.ReadText(adReadLine), adWriteLine)
    Loop
    Call objReadStream.Close
    'UTF-8 to UTF-8N
    With objWriteStream
        .Position = 0
        .Type = adTypeBinary
        .Position = 3
        bytData = .Read
        .Close
        .Open
        .Position = 0
        .Type = adTypeBinary
        .Write bytData
        .SaveToFile pstrFilePath, adSaveCreateOverWrite
        .Close
    End With
End Sub

Sub FileKill( _
    pstrPath As String _
    )
    With CreateObject("Scripting.FileSystemObject")
        If .FileExists(pstrPath) = True Then
            Call Kill(pstrPath)
        End If
    End With
End Sub

Function FileAddData( _
    pstrPath, _
    pstrTmpDataCell, _
    pstrSheetName, _
    Optional pstrCom = "", _
    Optional blnVisible = False _
    ) As Integer
    
    On Error GoTo Error
        If SheetCheck(pstrSheetName) = False Then
            ActiveWorkbook.Worksheets.Add
            ActiveSheet.Name = pstrSheetName
        End If
        With ActiveWorkbook.Worksheets(pstrSheetName).QueryTables.Add( _
            Connection:="TEXT;" & pstrPath, _
            Destination:=Range(pstrTmpDataCell) _
            )
            .Name = pstrSheetName
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .TextFilePromptOnRefresh = False
            .TextFilePlatform = 932
            '.TextFilePlatform = 1252
            .TextFileStartRow = 1
            .TextFileParseType = xlDelimited
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFileConsecutiveDelimiter = False
            .TextFileTabDelimiter = True
            .TextFileSemicolonDelimiter = False
            .TextFileCommaDelimiter = False
            .TextFileSpaceDelimiter = False
            ' .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
            .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1)
            .TextFileTrailingMinusNumbers = True
            .Refresh BackgroundQuery:=False
        End With
        Call Application.CalculateFullRebuild
        ActiveWorkbook.Worksheets(pstrSheetName).Visible = blnVisible
        FileAddData = 1
        Exit Function
Error:
    If MsgBox( _
        "*************" & vbCrLf & _
        " Error " & Err.Number & vbCrLf & _
        "*************" & vbCrLf & _
        vbCrLf & _
        " " & Err.Description & vbCrLf & _
        vbCrLf & _
        "※エラーの原因:" & vbCrLf & _
        "  1. 登録シートのフォーマットが違う" & vbCrLf & _
        "     => 修正してください" & vbCrLf & _
        "  2. 設定シート[main]の内容が壊れている" & vbCrLf & _
        "     => 非表示を解除して自力で解決" & vbCrLf & _
        "     => 再試行して、設定ファイルを作り直す" & vbCrLf & _
        "", _
        vbCritical + vbMsgBoxSetForeground + vbRetryCancel, _
        cstrMacroName & " " & cstrMacroVer) = vbRetry Then
        Call ActiveWorkbook.Worksheets(pstrSheetName).Delete
        FileAddData = 2
        Call QLCBMain(pstrCom)
    Else
        FileAddData = 3
        Exit Function
    End If
End Function

Sub FileList( _
    pstrDir _
    )
    Dim fsoFiles As Object
    Dim i As Integer
    Dim varData As Variant

    Set fsoFiles = CreateObject("Scripting.FileSystemObject").GetFolder(pstrDir).Files
        i = 3
        For Each varData In fsoFiles
            Cells(i, 2) = varData.Name
            Cells(i, 3) = varData.Type
            Cells(i, 4) = varData.DateLastModified
            i = i + 1
        Next
End Sub

Function FileList2( _
    pstrDir _
    ) As String
    Dim fsoFiles As Object
    Dim fsoFolders As Object
    Dim varData1 As Variant
    Dim varData2 As Variant
    
    Set fsoFiles = CreateObject("Scripting.FileSystemObject").GetFolder(pstrDir).Files
        For Each varData1 In fsoFiles
            If FileList2 = "" Then
                FileList2 = pstrDir & "\" & varData1.Name
            Else
                FileList2 = FileList2 & vbTab & pstrDir & "\" & varData1.Name
            End If
        Next
        Set fsoFolders = CreateObject("Scripting.FileSystemObject").GetFolder(pstrDir).SubFolders
            For Each varData2 In fsoFolders
                If FileList2 = "" Then
                    FileList2 = FileList2(pstrDir & "\" & varData2.Name)
                Else
                    FileList2 = FileList2 & vbTab & FileList2(pstrDir & "\" & varData2.Name)
                End If
            Next
        Set fsoFolders = Nothing
    Set fsoFiles = Nothing
    FileList2 = FileList2
End Function

