Attribute VB_Name = "ModMenu"
'******************************************************************
'***************Copyright PSST 2003********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive
Option Explicit
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_CLOSE = &H10
Private Const GWL_WNDPROC = (-4)
Public gOldProc As Long
Public NewMenu As Long
Public Sub SubClass(mhwnd As Long)
    'Start subclassing our form so we can respond to menu activity
    gOldProc& = GetWindowLong(mhwnd, GWL_WNDPROC)
    Call SetWindowLong(mhwnd, GWL_WNDPROC, AddressOf MenuProc)
End Sub
Public Sub UnSubClass(mhwnd As Long)
    'Stop subclassing - this is done automatically when calling form unloads
    Call SetWindowLong(mhwnd, GWL_WNDPROC, gOldProc&)
End Sub
Private Function MenuProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case wMsg&
        'Form closing - stop subclassing
        Case WM_CLOSE
            UnSubClass hwnd
        Case Else
            'respond to clicks on menus
            If wParam = NewMenu Then
                frmAbout.Show , Form1
            End If
    End Select
    MenuProc = CallWindowProc(gOldProc&, hwnd&, wMsg&, wParam&, lParam&)
End Function

