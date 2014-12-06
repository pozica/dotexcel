Attribute VB_Name = "object_regex"
Option Explicit

Function RegexBASPReplace( _
    pstrRegstr, _
    pstrTarget _
    )
    With CreateObject("basp21")
        RegexBASPReplace = .replace(pstrRegstr, pstrTarget)
    End With
End Function

Function RegexBASPMatch( _
    pstrRegstr, _
    pstrTarget _
    )
    With CreateObject("basp21")
        RegexBASPMatch = .Match(pstrRegstr, pstrTarget)
    End With
End Function


