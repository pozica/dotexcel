Attribute VB_Name = "object_excel_msgbox"
Option Explicit

Public Declare Function SetWindowsHookEx Lib "user32.dll" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32.dll" (ByVal hhk As Long) As Long
Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
Public Const GWL_HINSTANCE = (-6)
Public Const WH_CBT = 5
Public Const HCBT_ACTIVATE = 5
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOACTIVATE = &H10
Public glngHook As Long
Public glngX As Long
Public glngY As Long

Public Function MsgBoxHookProc(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = HCBT_ACTIVATE Then
        Call SetWindowPos(wParam, 0, glngX, glngY, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE)
        UnhookWindowsHookEx glngHook
    End If
    MsgBoxHookProc = False
End Function

Public Sub MsgBoxHookTest()
    glngX = 10 'per pixel
    glngY = 10
    glngHook = SetWindowsHookEx(WH_CBT, AddressOf MsgBoxHookProc, GetWindowLong(FindWindow("XLMAIN", Application.Caption), GWL_HINSTANCE), 0)
    MsgBox "Set done", vbInformation
End Sub

