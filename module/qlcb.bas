Attribute VB_Name = "qlcb"
Option Explicit

Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function SendMail Lib "bsmtp" (szServer As String, szTo As String, szFrom As String, szSubject As String, szBody As String, szFile As String) As String
Public Declare Function FlushMail Lib "bsmtp" (szServer As String, szDir As String, szLogfile As String) As Long
Public Const cstrMacroName = "QLCB"
Public Const cstrMacroVer = "0.3.0.01"
Public Const cstrServer = "\\192.168.1.100\"
Public Const cstrMailServer = "example.com"
Public Const cstrWSName1 = "main"
Public Const cstrWSName2 = "relay"
Public Const cstrEntryIDCell1 = "R1C"
Public Const cstrEntryIDCell2 = 27
Public Const cstrEntryIDCell3 = 52
Public Const cstrEntryBasenameCell = cstrWSName1 & "!B1"
Public Const cstrEntryDateCell = cstrWSName1 & "!C1"
Public Const cstrMediaNameCell = cstrWSName1 & "!M1"
Public Const carrImageNameCell1 = cstrWSName1 & "!R1C"
Public Const carrImageNameCell2 = 53
Public Const cstrTmpDataCell = cstrWSName2 & "!A"
Public Const cstrTmpFileName = "import"
Public Const cstrTmpFileRelativePath = ".\_tmp"
Public hdcQLCBMTOS As String
Public hdcQLCB As String

Sub QLCBMain( _
    Optional pstrCom = "", _
    Optional pstrCommand = "" _
    )
    If pstrCommand = "" Then
        pstrCom = QLCBDisplay(pstrCom, hdcQLCB)
    Else
        pstrCom = QLCBDisplay(pstrCom, pstrCommand)
    End If
    If pstrCom = "" Then Exit Sub
    If Workbooks.count > 0 Then Call FileCheck2(ActiveWorkbook.Path & "\_tmp")
    Call QLCBAlias(pstrCom)
End Sub

