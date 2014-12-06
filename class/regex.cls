VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "regex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private re As Object

Private Sub Class_Initialize()
    Set re = CreateObject("VBScript.RegExp")
End Sub

Private Sub Class_Terminate()
    Set re = Nothing
End Sub

Public Sub SetOption(ByVal isGlobal As Boolean, ByVal isIgnoreCase As Boolean, ByVal isMultiLine As Boolean)
    With re
        .Global = isGlobal
        .IgnoreCase = isIgnoreCase
        .MultiLine = isMultiLine
    End With
End Sub

Public Function replace(ByVal SourceString As String, ByVal Pattern As String, ByVal ReplaceString As String) As String
    With re
        .Pattern = Pattern
        replace = .replace(SourceString, ReplaceString)
    End With
End Function

Public Function test(ByVal SourceString As String, ByVal Pattern As String) As Boolean
    With re
        .Pattern = Pattern
        test = .test(SourceString)
    End With
End Function

Public Function Execute(ByVal SourceString As String, ByVal Pattern As String, ByVal index As Long) As String
    With re
        .Pattern = Pattern
        Execute = .Execute(SourceString)(index - 1)
    End With
End Function

Public Function count(ByVal SourceString As String, ByVal Pattern As String) As Long
    With re
        .Pattern = Pattern
        count = .Execute(SourceString).count
    End With
End Function




