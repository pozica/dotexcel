Attribute VB_Name = "util_shell"
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = &H400&
Private Const STILL_ACTIVE = &H103&

Private Sub ShellCmdEnd( _
    ProcessID As Long _
    )
    Dim hProcess As Long
    Dim EndCode As Long
    Dim EndRet  As Long

    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 1, ProcessID)
        Do
            EndRet = GetExitCodeProcess(hProcess, EndCode)
            DoEvents
        Loop While (EndCode = STILL_ACTIVE)
    EndRet = CloseHandle(hProcess)
End Sub

Sub ShellCmd( _
    pstrCmd _
    )
    Dim x As Long
    Dim y As String
    Dim z As String

    x = Shell("cmd /c " & pstrCmd & " > " & ThisWorkbook.Path & "\tmp")
    Call ShellCmdEnd(x)
    y = ActiveCell.Address
    z = ActiveSheet.Name
    Call FileAddData(ThisWorkbook.Path & "\tmp", y, z, , True)
End Sub

Sub ShellCmd2( _
    pstrCmd _
    )
    Dim WSH
    Dim wExec
    Dim strData As String
    Dim varData As Variant
    Dim i As Integer
        
    Set WSH = CreateObject("WScript.Shell")
        Set wExec = WSH.Run("%ComSpec% /c " & pstrCmd, 0)
            Do While wExec.Status = 0
                DoEvents
            Loop
            strData = wExec.StdOut.ReadAll
            varData = Split(strData, vbCrLf)
            On Error Resume Next
                For i = 0 To UBound(varData)
                    ActiveCell.Offset(i, 0).Value = varData(i)
                Next
                ActiveCell.Offset(UBound(varData) + 1, 0).Select
        Set wExec = Nothing
    Set WSH = Nothing
End Sub

Sub ShellCmd3( _
    pstrCmd As String _
    )
    Dim strData As String
    Dim varData As Variant
    
    strData = ShellExecCmd(pstrCmd)
    varData = Split(strData, vbCrLf)
    On Error Resume Next
        For i = 0 To UBound(varData)
            ActiveCell.Offset(i, 0).Value = varData(i)
        Next
        ActiveCell.Offset(UBound(varData) + 1, 0).Select
        EmacsModeRebottom
    On Error GoTo 0
End Sub

Function ShellExecCmd( _
    pstrCmd As String _
    ) As String
    Const TemporaryFolder = 2
    Dim oShell As Object
    Dim fso As Object
    Dim fdr As Object
    Dim ts As Object
    Dim strFileName As String

    Set oShell = CreateObject("WScript.Shell")
        Set fso = CreateObject("Scripting.FileSystemObject")
            Set fdr = fso.GetSpecialFolder(TemporaryFolder)
                
                Do
                    strFileName = fso.BuildPath(fdr.Path, fso.GetTempName)
                Loop While fso.FileExists(strFileName)
            
                Call oShell.Run("%ComSpec% /c " & pstrCmd & ">" & strFileName & " 2<&1", 0, True)
            
                If fso.FileExists(strFileName) Then
                    Set ts = fso.OpenTextFile(strFileName)
                        ShellExecCmd = ts.ReadAll
                        ts.Close
                        Kill strFileName
                    Set ts = Nothing
                End If
            Set fdr = Nothing
        Set fso = Nothing
    Set oShell = Nothing
End Function

Sub ShellZsh2( _
    pstrCmd _
    )
    Dim WSH
    Dim wExec
    Dim strData As String
    Dim varData As Variant
    Dim i As Integer
    
    Set WSH = CreateObject("WScript.Shell")
        Set wExec = WSH.Exec("zsh -c '" & pstrCmd & "'")
            Do While wExec.Status = 0
                DoEvents
            Loop
            strData = wExec.StdOut.ReadAll
            varData = Split(strData, vbLf)
            On Error Resume Next
                For i = 0 To UBound(varData)
                    ActiveCell.Offset(i, 0).Value = varData(i)
                Next
                ActiveCell.Offset(UBound(varData) + 1, 0).Select
        Set wExec = Nothing
    Set WSH = Nothing
End Sub

Sub ShellZsh3( _
    pstrCmd As String _
    )
    Dim strData As String
    Dim varData As Variant
    
    strData = ShellExecZsh(pstrCmd)
    varData = Split(strData, vbLf)
    On Error Resume Next
        For i = 0 To UBound(varData)
            ActiveCell.Offset(i, 0).Value = varData(i)
        Next
        ActiveCell.Offset(UBound(varData) + 1, 0).Select
        EmacsModeRebottom
    On Error GoTo 0
End Sub

Function ShellExecZsh( _
    pstrCmd As String _
    ) As String
    Const TemporaryFolder = 2
    Dim oShell As Object
    Dim fso As Object
    Dim fdr As Object
    Dim ts As Object
    Dim strFileName As String

    Set oShell = CreateObject("WScript.Shell")
        Set fso = CreateObject("Scripting.FileSystemObject")
            Set fdr = fso.GetSpecialFolder(TemporaryFolder)
                
                Do
                    strFileName = fso.BuildPath(fdr.Path, fso.GetTempName)
                Loop While fso.FileExists(strFileName)
            
                Call oShell.Run("zsh -c """ & pstrCmd & " > `cygpath -u '" & strFileName & "'` 2<&1""", 0, True)
            
                If fso.FileExists(strFileName) Then
                    Set ts = fso.OpenTextFile(strFileName)
                        ShellExecZsh = ts.ReadAll
                        ts.Close
                        Kill strFileName
                    Set ts = Nothing
                End If
            Set fdr = Nothing
        Set fso = Nothing
    Set oShell = Nothing
End Function

Function ShellVBA( _
    pstrCmd _
    )
    pstrCmd = Application.Evaluate(pstrCmd)
    ActiveCell.Value = pstrCmd
    ActiveCell.Offset(2, 0).Select
    EmacsModeRebottom
End Function
